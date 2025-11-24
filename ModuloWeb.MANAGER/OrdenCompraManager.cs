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
            // 1. Inserta encabezado en BD
            int idOrden = broker.InsertarOrden(idProveedor, total);

            // 2. Inserta detalle en BD
            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // 3. Generar Excel LEYENDO lo que quedó en la BD
            string rutaExcel = GenerarExcel(idOrden);

            // 4. Enviar correo (desde SendGrid) al proveedor
            EnviarCorreo(idOrden, rutaExcel);

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
            string proveedorNombre = "";
            string proveedorNit = "";
            string proveedorDireccion = "";
            string proveedorTelefono = "";
            decimal totalOrden = 0;
            DateTime fechaOrden = DateTime.Now;

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  o.id_orden,
                            o.id_proveedor,
                            o.total,
                            o.fecha,
                            p.nombre   AS proveedor_nombre,
                            p.nit      AS proveedor_nit,
                            p.direccion AS proveedor_direccion,
                            p.telefono  AS proveedor_telefono
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new Exception($"No se encontró la orden {idOrden} en la base de datos.");

                    proveedorNombre = reader.GetString("proveedor_nombre");
                    proveedorNit = reader.IsDBNull(reader.GetOrdinal("proveedor_nit"))
                        ? ""
                        : reader.GetString("proveedor_nit");
                    proveedorDireccion = reader.IsDBNull(reader.GetOrdinal("proveedor_direccion"))
                        ? ""
                        : reader.GetString("proveedor_direccion");
                    proveedorTelefono = reader.IsDBNull(reader.GetOrdinal("proveedor_telefono"))
                        ? ""
                        : reader.GetString("proveedor_telefono");

                    totalOrden = reader.GetDecimal("total");
                    fechaOrden = reader.GetDateTime("fecha");
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
                var ws = wb.Worksheet(1); // Hoja1

                // =============== ENCABEZADO ===============
                // Celdas según tu plantilla (ajusta si cambian):
                // Proveedor (nombre)       -> fila 5, col 2 (B5)
                // NIT/CC                   -> fila 6, col 2 (B6)
                // Dirección                -> fila 8, col 2 (B8)
                // Teléfono                 -> fila 7, col 2 (B7)
                // Fecha Orden              -> fila 11, col 3 (C11)
                // Condiciones de pago      -> fila 13, col 3 (C13)
                // Moneda (COP)             -> fila 11, col 6 (F11)

                ws.Cell(5, 2).Value = proveedorNombre;
                ws.Cell(6, 2).Value = proveedorNit;
                ws.Cell(7, 2).Value = proveedorTelefono;
                ws.Cell(8, 2).Value = proveedorDireccion;

                ws.Cell(11, 3).Value = fechaOrden;
                ws.Cell(11, 6).Value = "COP";
                ws.Cell(13, 3).Value = "30 días";

                // =============== DETALLE ===============
                // Fila inicial para ítems (según plantilla)
                int filaInicio = 19;

                // Columnas (según encabezados de tu plantilla):
                const int COL_LN          = 2;   // B  -> Ln
                const int COL_ITEM        = 4;   // D  -> Item
                const int COL_DESCRIPCION = 10;  // J  -> Descripcion
                const int COL_PRECIO_UNIT = 12;  // L  -> Precio Unit.
                const int COL_CANTIDAD    = 13;  // M  -> Cantidad
                const int COL_VALOR_TOTAL = 14;  // N  -> Valor Total

                int fila = filaInicio;
                int linea = 1;
                decimal subtotalAcumulado = 0;

                foreach (var d in detalles)
                {
                    // línea e item
                    ws.Cell(fila, COL_LN).Value = linea;
                    ws.Cell(fila, COL_ITEM).Value = linea; // o d.idProducto si prefieres

                    // Descripción
                    ws.Cell(fila, COL_DESCRIPCION).Value = d.nombre;

                    // Precio unitario, cantidad y total
                    ws.Cell(fila, COL_PRECIO_UNIT).Value = d.precio;
                    ws.Cell(fila, COL_CANTIDAD).Value    = d.cantidad;
                    ws.Cell(fila, COL_VALOR_TOTAL).Value = d.subtotal;

                    subtotalAcumulado += d.subtotal;

                    fila++;
                    linea++;
                }

                // =============== SUBTOTAL / TOTAL ===============
                // Según la plantilla:
                // "Subtotal:"   -> fila 22, col 13 (M22), valor en col 14 (N22)
                // "Descuento:"  -> fila 23, col 13 (M23), valor en col 14 (N23)
                // "Total:"      -> fila 24, col 13 (M24), valor en col 14 (N24)
                const int FILA_SUBTOTAL = 22;
                const int FILA_DESCTO   = 23;
                const int FILA_TOTAL    = 24;
                const int COL_VALOR_RESUMEN = 14; // N

                ws.Cell(FILA_SUBTOTAL, COL_VALOR_RESUMEN).Value = subtotalAcumulado;

                // Por ahora no aplicamos descuentos
                ws.Cell(FILA_DESCTO, COL_VALOR_RESUMEN).Value = 0m;

                // Total final = subtotal - descuento
                ws.Cell(FILA_TOTAL, COL_VALOR_RESUMEN).Value = subtotalAcumulado;

                // Si quieres fijar un zoom cómodo (para que no se vea "gigante"):
                // ws.SheetView.ZoomScale = 90;

                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // =====================================
        // 4. Enviar correo con SendGrid (API)
        // =====================================
        private void EnviarCorreo(int idOrden, string rutaExcel)
        {
            // Buscar correo del proveedor desde la orden
            string correoDestino = "";
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT p.correo
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);
                object result = cmd.ExecuteScalar();
                correoDestino = result?.ToString() ?? "";
            }

            if (string.IsNullOrWhiteSpace(correoDestino))
            {
                Console.WriteLine("Proveedor sin correo, no se envía email.");
                return;
            }

            // From y API key desde variables de entorno
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

            // Adjuntar Excel
            byte[] bytes  = File.ReadAllBytes(rutaExcel);
            string base64 = Convert.ToBase64String(bytes);

            msg.AddAttachment(
                Path.GetFileName(rutaExcel),
                base64,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");

            var body = response.Body.ReadAsStringAsync().Result;
            Console.WriteLine($"SENDGRID BODY: {body}");
        }
    }
}
