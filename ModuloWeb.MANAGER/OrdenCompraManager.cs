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
                return new MySqlConnection(cs);   // para Railway

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

            // 3. Generar Excel leyendo la orden desde BD
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
        // ==========================================
        private string GenerarExcel(int idOrden)
        {
            // ----- 3.1. Leer encabezado + proveedor -----
            int idProveedor = 0;
            string provNombre = "";
            string provNit = "";
            string provCiudad = "";
            string provDireccion = "";
            string provTelefono = "";
            decimal total = 0;
            DateTime fecha = DateTime.Now;
            string condiciones = "30 días"; // si luego lo quieres guardar en BD, se cambia aquí

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  o.id_orden,
                            o.id_proveedor,
                            o.total,
                            o.fecha,
                            p.nombre      AS prov_nombre,
                            p.nit         AS prov_nit,
                            p.direccion   AS prov_direccion,
                            p.telefono    AS prov_telefono
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new Exception($"No se encontró la orden {idOrden} en la base de datos.");

                    idProveedor   = reader.GetInt32("id_proveedor");
                    provNombre    = reader.GetString("prov_nombre");
                    provNit       = reader.IsDBNull(reader.GetOrdinal("prov_nit")) ? "" : reader.GetString("prov_nit");
                    provDireccion = reader.IsDBNull(reader.GetOrdinal("prov_direccion")) ? "" : reader.GetString("prov_direccion");
                    provTelefono  = reader.IsDBNull(reader.GetOrdinal("prov_telefono")) ? "" : reader.GetString("prov_telefono");
                    // Ciudad no está en la tabla, la dejamos vacía por ahora
                    provCiudad    = "";

                    total = reader.GetDecimal("total");
                    fecha = reader.GetDateTime("fecha");
                }
            }

            // ----- 3.2. Leer detalles con nombre de producto -----
            var detalles = new List<(int idProducto, string nombre, int cantidad, decimal precio, decimal subtotal)>();

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

            // ----- 3.3. Preparar rutas -----
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
                var ws = wb.Worksheet(1);

                // === ENCABEZADO PROVEEDOR (panel izquierdo) ===
                // NO tocamos nada a la derecha (Facturar a: SUPLINDUSTRIA...)
                ws.Cell("D8").Value  = provNombre;
                ws.Cell("D9").Value  = provNit;
                ws.Cell("D10").Value = provCiudad;
                ws.Cell("D11").Value = provDireccion;
                ws.Cell("D13").Value = provTelefono;

                // Fecha de la orden (fila 15, bajo "Fecha Orden")
                ws.Cell("B15").Value = fecha;
                ws.Cell("B15").Style.DateFormat.Format = "dd/MM/yyyy";

                // Moneda (ejemplo fijo)
                ws.Cell("E15").Value = "COP";

                // Condiciones de pago (fila 16)
                ws.Cell("B16").Value = $"Condiciones de pago: {condiciones}";

                // === DETALLES DE LÍNEA ===
                int fila = 19; // La primera línea de detalle en la plantilla
                int linea = 1;

                foreach (var d in detalles)
                {
                    // Columna B: Ln
                    ws.Cell(fila, "B").Value = linea;

                    // Columna D: Item (puedes usar el id del producto)
                    ws.Cell(fila, "D").Value = d.idProducto;

                    // Columna G: Descripcion
                    ws.Cell(fila, "G").Value = d.nombre;

                    // Columna J: Cantidad
                    ws.Cell(fila, "J").Value = d.cantidad;

                    // Columna L: Precio Unit.
                    ws.Cell(fila, "L").Value = d.precio;
                    ws.Cell(fila, "L").Style.NumberFormat.Format = "#,##0";

                    // Columna N: Valor Total
                    ws.Cell(fila, "N").Value = d.subtotal;
                    ws.Cell(fila, "N").Style.NumberFormat.Format = "#,##0";

                    fila++;
                    linea++;
                }

                // === SUBTOTAL / TOTAL EN LA PARTE INFERIOR ===
                ws.Cell("N22").Value = total; // Subtotal
                ws.Cell("N22").Style.NumberFormat.Format = "#,##0";

                // Si no manejas descuento, el total es el mismo
                ws.Cell("N24").Value = total; // Total
                ws.Cell("N24").Style.NumberFormat.Format = "#,##0";

                // Guarda el archivo rellenado
                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // =====================================
        // 4. Enviar correo con SendGrid (API)
        // =====================================
        private void EnviarCorreo(int idOrden, int idProveedor, string rutaExcel)
        {
            // 1. Correo del proveedor (desde BROKER)
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

            // 4. Enviar y LOG COMPLETO
            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");

            var body = response.Body.ReadAsStringAsync().Result;
            Console.WriteLine($"SENDGRID BODY: {body}");
        }
    }
}
