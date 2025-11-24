using ModuloWeb.BROKER;
using ModuloWeb.ENTITIES;
using MySql.Data.MySqlClient;
using ClosedXML.Excel;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // =====================================================
        // 0. Helper: conexión (Railway o local según el entorno)
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

            // 3. Generar Excel leyendo TODO desde BD (encabezado + detalle)
            string rutaExcel = GenerarExcel(idOrden);

            // 4. Enviar correo al proveedor con el Excel adjunto
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
        // ==========================================
        private string GenerarExcel(int idOrden)
        {
            // ----- 3.1. Leer encabezado + datos del proveedor -----
            string proveedorNombre = "";
            string proveedorNit = "";
            string proveedorCiudad = "";
            string proveedorDireccion = "";
            string proveedorTelefono = "";
            decimal total = 0;
            DateTime fecha = DateTime.Now;

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  o.id_orden,
                            o.id_proveedor,
                            o.total,
                            o.fecha,
                            p.nombre    AS proveedor_nombre,
                            p.nit       AS proveedor_nit,
                            p.correo    AS proveedor_correo,
                            p.telefono  AS proveedor_telefono,
                            p.direccion AS proveedor_direccion
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new Exception($"No se encontró la orden {idOrden} en la base de datos.");

                    proveedorNombre    = reader.GetString(reader.GetOrdinal("proveedor_nombre"));
                    proveedorNit       = reader.IsDBNull(reader.GetOrdinal("proveedor_nit")) ? "" : reader.GetString(reader.GetOrdinal("proveedor_nit"));
                    proveedorCiudad    = ""; // si luego agregas ciudad en la tabla, la lees aquí
                    proveedorDireccion = reader.IsDBNull(reader.GetOrdinal("proveedor_direccion")) ? "" : reader.GetString(reader.GetOrdinal("proveedor_direccion"));
                    proveedorTelefono  = reader.IsDBNull(reader.GetOrdinal("proveedor_telefono")) ? "" : reader.GetString(reader.GetOrdinal("proveedor_telefono"));

                    total = reader.GetDecimal(reader.GetOrdinal("total"));
                    fecha = reader.GetDateTime(reader.GetOrdinal("fecha"));
                }
            }

            // ----- 3.2. Leer detalles con nombre de producto -----
            var detalles = new List<(int idProducto, string nombre, int cantidad, decimal precio, decimal subtotal)>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  d.id_producto,
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
                            reader.GetInt32(reader.GetOrdinal("id_producto")),
                            reader.GetString(reader.GetOrdinal("nombre")),
                            reader.GetInt32(reader.GetOrdinal("cantidad")),
                            reader.GetDecimal(reader.GetOrdinal("precio")),
                            reader.GetDecimal(reader.GetOrdinal("subtotal"))
                        ));
                    }
                }
            }

            // ----- 3.3. Rutas -----
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
                throw new Exception($"No se encuentra la plantilla PlantillaOrdenes.xlsx en {rutaPlantilla}.");

            // ----- 3.4. Abrir plantilla y rellenar -----
            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                // Hoja principal (la que ves en la captura)
                var hoja = wb.Worksheet("Hoja1");

                // Hoja de soporte para fórmulas de fecha, condiciones, etc.
                var instancia = wb.Worksheet("Instancia");

                // ===== ENCABEZADO EN LA HOJA VISIBLE =====
                // (ajusta si en tu plantilla estas celdas cambian)
                hoja.Cell("B5").Value = proveedorNombre;       // Proveedor
                hoja.Cell("B6").Value = proveedorNit;          // NIT/CC
                hoja.Cell("B7").Value = proveedorCiudad;       // Ciudad
                hoja.Cell("B8").Value = proveedorDireccion;    // Dirección
                hoja.Cell("B9").Value = proveedorTelefono;     // Contacto / Teléfono

                // Número de orden si quieres mostrarlo arriba a la izquierda
                hoja.Cell("A1").Value = idOrden;

                // ===== DATOS DE LA ORDEN EN LA HOJA INSTANCIA =====
                // Estas celdas se usan en fórmulas de Hoja1:
                instancia.Cell("E2").Value = fecha;        // Fecha de la orden  -> Hoja1!B15
                instancia.Cell("F2").Value = "COP";        // Moneda              -> Hoja1!E15
                instancia.Cell("I2").Value = "30 días";    // Condiciones pago    -> Hoja1!B16
                instancia.Cell("P2").Value = "ANNI CARMONA"; // Comprador (fijo o de BD si luego lo agregas)

                // ===== DETALLES (Líneas de la orden) =====
                // En tu plantilla la tabla comienza en la fila 19:
                //  B: Ln,  D: Item,  G: Descripción,  J: Cantidad,  L: Precio Unit,  N: Valor Total
                int fila = 19;
                int linea = 1;

                foreach (var d in detalles)
                {
                    hoja.Cell(fila, "B").Value = linea;        // Ln
                    hoja.Cell(fila, "D").Value = d.idProducto; // Item / código
                    hoja.Cell(fila, "G").Value = d.nombre;     // Descripción
                    hoja.Cell(fila, "J").Value = d.cantidad;   // Cantidad
                    hoja.Cell(fila, "L").Value = d.precio;     // Precio unitario
                    hoja.Cell(fila, "N").Value = d.subtotal;   // Valor total de la línea

                    fila++;
                    linea++;
                }

                // Si quieres dejar también el total de la orden en una celda específica, por ejemplo:
                // hoja.Cell("N25").Value = total;

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

            // 3. Adjuntar el Excel
            byte[] bytes  = File.ReadAllBytes(rutaExcel);
            string base64 = Convert.ToBase64String(bytes);

            msg.AddAttachment(
                Path.GetFileName(rutaExcel),
                base64,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            // 4. Enviar y loguear respuesta
            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");

            var body = response.Body.ReadAsStringAsync().Result;
            Console.WriteLine($"SENDGRID BODY: {body}");
        }
    }
}
