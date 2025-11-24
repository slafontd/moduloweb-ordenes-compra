using ModuloWeb.BROKER;
using ModuloWeb.ENTITIES;
using MySql.Data.MySqlClient;
using ClosedXML.Excel;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // =====================================================
        // 0. Helper: conexión (Railway o local según entorno)
        // =====================================================
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");
            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);

            // modo local
            return ConexionBD.Conectar();
        }

        // =======================================
        // 1. Crear ORDEN + DETALLES + EXCEL + MAIL
        // =======================================
        public int CrearOrden(
            int idProveedor,
            decimal total,
            List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // 1. Inserta encabezado
            int idOrden = broker.InsertarOrden(idProveedor, total);

            // 2. Inserta detalle
            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // 3. Generar Excel leyendo la info desde BD
            string rutaExcel = GenerarExcel(idOrden);

            // 4. Enviar correo al proveedor
            EnviarCorreo(idOrden, idProveedor, rutaExcel);

            return idOrden;
        }

        // ===========================
        // 2. Listar órdenes desde BD
        // ===========================
        public List<OrdenCompra> ObtenerOrdenes()
        {
            return broker.ObtenerOrdenes();
        }

        // ==========================================
        // 3. Generar EXCEL a partir de lo guardado
        //    (usa la plantilla PlantillaOrdenes.xlsx)
        // ==========================================
        private string GenerarExcel(int idOrden)
        {
            // ---------- 3.1 Encabezado + datos proveedor ----------
            string proveedorNombre   = "";
            string proveedorNit      = "";
            string proveedorCiudad   = "";
            string proveedorDireccion= "";
            string proveedorTelefono = "";
            DateTime fecha           = DateTime.Now;
            string condicionesPago   = "30 días";   // si luego quieres, lo sacas de la BD
            decimal totalOrden       = 0;

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  o.id_orden,
                            o.id_proveedor,
                            o.total,
                            o.fecha,
                            p.nombre       AS proveedor_nombre,
                            p.nit          AS proveedor_nit,
                            p.direccion    AS proveedor_direccion,
                            p.telefono     AS proveedor_telefono
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new Exception($"No se encontró la orden {idOrden} en la base de datos.");

                    proveedorNombre    = reader.GetString("proveedor_nombre");
                    proveedorNit       = reader.IsDBNull(reader.GetOrdinal("proveedor_nit"))
                                            ? "" : reader.GetString("proveedor_nit");
                    proveedorDireccion = reader.IsDBNull(reader.GetOrdinal("proveedor_direccion"))
                                            ? "" : reader.GetString("proveedor_direccion");
                    proveedorTelefono  = reader.IsDBNull(reader.GetOrdinal("proveedor_telefono"))
                                            ? "" : reader.GetString("proveedor_telefono");

                    // ciudad la sacamos “por partes” de la dirección si no tienes campo ciudad
                    proveedorCiudad = "";

                    totalOrden = reader.GetDecimal("total");
                    fecha      = reader.GetDateTime("fecha");
                }
            }

            // ---------- 3.2 Detalles con nombre de producto ----------
            var detalles = new List<(int idProducto,
                                     string nombre,
                                     int cantidad,
                                     decimal precio,
                                     decimal subtotal)>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT d.id_producto,
                           pr.nombre,
                           d.cantidad,
                           d.precio,
                           d.subtotal
                    FROM detalle_orden d
                    JOIN productos pr ON pr.id = d.id_producto
                    WHERE d.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        detalles.Add((
                            reader.GetInt32("id_producto"),
                            reader.GetString("nombre"),
                            reader.GetInt32("cantidad"),
                            reader.GetDecimal("precio"),
                            reader.GetDecimal("subtotal")
                        ));
                    }
                }
            }

            // Subtotal (suma de líneas)
            decimal subtotal = detalles.Sum(d => d.subtotal);

            // ---------- 3.3 Rutas ----------
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
                throw new Exception($"No se encuentra la plantilla PlantillaOrdenes.xlsx en {rutaPlantilla}.");

            // ---------- 3.4 Abrir plantilla y rellenar ----------
            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var ws = wb.Worksheet(1);

                // ===== BLOQUE IZQUIERDO: DATOS PROVEEDOR =====
                // (dejamos las etiquetas en la col B y escribimos en la col C)
                ws.Cell("C8").Value  = proveedorNombre;      // Proveedor
                ws.Cell("C9").Value  = proveedorNit;         // NIT/CC
                ws.Cell("C10").Value = proveedorCiudad;      // Ciudad
                ws.Cell("C11").Value = proveedorDireccion;   // Dirección
                ws.Cell("C12").Value = proveedorTelefono;    // Teléfono

                // Fecha orden / moneda / comprador / condiciones
                ws.Cell("C14").Value = fecha;
                ws.Cell("C14").Style.DateFormat.Format = "dd/MM/yyyy";

                ws.Cell("F14").Value = "COP";                // Moneda
                ws.Cell("H14").Value = "ANNI CARMONA";       // Comprador fijo (o lo haces variable)
                ws.Cell("C16").Value = $"Condiciones de pago: {condicionesPago}";

                // ===== DETALLES DE LÍNEA =====
                int fila = 19;  // primera fila de detalle según la plantilla

                int linea = 1;
                foreach (var d in detalles)
                {
                    ws.Cell(fila, "B").Value = linea;           // Ln
                    ws.Cell(fila, "D").Value = d.idProducto;    // Item (puedes poner código)
                    // Catalogo (E) y Modelo (F) los dejamos vacíos por ahora
                    ws.Cell(fila, "G").Value = d.nombre;        // Descripción
                    ws.Cell(fila, "J").Value = d.cantidad;      // Cantidad
                    ws.Cell(fila, "L").Value = d.precio;        // Precio Unit.
                    ws.Cell(fila, "N").Value = d.subtotal;      // Valor Total línea

                    fila++;
                    linea++;
                }

                // ===== RESUMEN: SUBTOTAL / DESCUENTO / TOTAL =====
                ws.Cell("N22").Value = subtotal;    // Subtotal
                ws.Cell("N23").Value = 0m;          // Descuento
                ws.Cell("N24").Value = totalOrden;  // Total de la orden (de BD)

                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // =====================================
        // 4. Enviar correo con SendGrid (API)
        // =====================================
        private void EnviarCorreo(int idOrden, int idProveedor, string rutaExcel)
        {
            // 1. Correo del proveedor
            string correoDestino = broker.ObtenerCorreoProveedor(idProveedor);

            if (string.IsNullOrWhiteSpace(correoDestino))
            {
                Console.WriteLine("Proveedor sin correo, no se envía email.");
                return;
            }

            // 2. From y API key desde variables de entorno
            string fromEmail = Environment.GetEnvironmentVariable("FROM_EMAIL");
            string apiKey    = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");

            if (string.IsNullOrWhiteSpace(fromEmail))
            {
                Console.WriteLine("FROM_EMAIL no está configurado.");
                return;
            }

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                Console.WriteLine("SENDGRID_API_KEY no está configurado.");
                return;
            }

            var client = new SendGridClient(apiKey);

            var from = new EmailAddress(fromEmail, "Sistema de Órdenes");
            var to   = new EmailAddress(correoDestino);

            string subject   = $"Orden de Compra #{idOrden}";
            string plainText = "Adjunto la orden de compra generada automáticamente.";

            var msg = MailHelper.CreateSingleEmail(from, to, subject, plainText, null);

            // 3. Adjuntar el Excel generado
            byte[] bytes  = File.ReadAllBytes(rutaExcel);
            string base64 = Convert.ToBase64String(bytes);

            msg.AddAttachment(
                Path.GetFileName(rutaExcel),
                base64,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            // 4. Enviar y LOG
            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");
            var body = response.Body.ReadAsStringAsync().Result;
            Console.WriteLine($"SENDGRID BODY: {body}");
        }
    }
}
