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

            // 3. Generar Excel LEYENDO la orden desde BD
            string rutaExcel = GenerarExcel(idOrden);

            // 4. Enviar correo al proveedor
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
            decimal total = 0;
            DateTime fecha = DateTime.Now;

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT o.id_orden,
                           o.id_proveedor,
                           o.total,
                           o.fecha,
                           p.nombre   AS proveedor_nombre,
                           p.nit      AS proveedor_nit
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

                // ENCABEZADO – AJUSTA CELDAS SEGÚN TU PLANTILLA
                ws.Cell("B2").Value = idOrden;
                ws.Cell("B3").Value = proveedorNombre;
                ws.Cell("B4").Value = proveedorNit;
                ws.Cell("B5").Value = fecha.ToString("dd/MM/yyyy HH:mm");
                ws.Cell("B6").Value = total;

                // DETALLES – fila inicial (ajústala a tu plantilla)
                int fila = 10;   // por ejemplo, fila 10

                foreach (var d in detalles)
                {
                    ws.Cell(fila, 1).Value = d.idProducto;
                    ws.Cell(fila, 2).Value = d.nombre;
                    ws.Cell(fila, 3).Value = d.cantidad;
                    ws.Cell(fila, 4).Value = d.precio;
                    ws.Cell(fila, 5).Value = d.subtotal;
                    fila++;
                }

                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // =====================================
        // 4. Enviar correo con SendGrid (API)
        // =====================================
        private void EnviarCorreo(int idOrden, string archivoExcel)
        {
            string apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY") ?? "";
            string fromEmail = Environment.GetEnvironmentVariable("FROM_EMAIL") ?? "";

            if (string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(fromEmail))
            {
                Console.WriteLine("SENDGRID_API_KEY o FROM_EMAIL no configurados. No se envía correo.");
                return;
            }

            // correo del proveedor desde la BD
            string correoDestino = broker.ObtenerCorreoProveedor(idOrden);

            var client = new SendGridClient(apiKey);
            var from = new EmailAddress(fromEmail, "Sistema de Órdenes");
            var to = new EmailAddress(correoDestino);

            string subject = $"Orden de Compra #{idOrden}";
            string textoPlano = "Adjunto la orden de compra generada automáticamente.";

            var msg = MailHelper.CreateSingleEmail(from, to, subject, textoPlano, null);

            // Adjuntar Excel
            byte[] bytes = File.ReadAllBytes(archivoExcel);
            string base64 = Convert.ToBase64String(bytes);
            msg.AddAttachment(Path.GetFileName(archivoExcel), base64);

            var response = client.SendEmailAsync(msg).Result;

            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");
        }
    }
}
