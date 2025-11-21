using ModuloWeb.BROKER;
using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ModuloWeb.ENTITIES;
using MySql.Data.MySqlClient;
using System.Collections.Generic;
using SendGrid;
using SendGrid.Helpers.Mail;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // Helper: crea la conexión a MySQL
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);

            // Para desarrollo local
            return ConexionBD.Conectar();
        }

        // Crear nueva orden de compra
        public int CrearOrden(int idProveedor, decimal total, List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            int idOrden = broker.InsertarOrden(idProveedor, total);

            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            string rutaPDF = GenerarPDF(idOrden, idProveedor, total, detalles);

            EnviarCorreo(idOrden, idProveedor, rutaPDF);

            return idOrden;
        }

        // Obtener órdenes
        public List<OrdenCompra> ObtenerOrdenes()
        {
            List<OrdenCompra> lista = new List<OrdenCompra>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id_orden, id_proveedor, total, fecha, estado FROM ordenes_compra",
                    con
                );

                var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lista.Add(new OrdenCompra
                    {
                        IdOrden = reader.GetInt32("id_orden"),
                        IdProveedor = reader.GetInt32("id_proveedor"),
                        Total = reader.GetDecimal("total"),
                        Fecha = reader.GetDateTime("fecha"),
                        Estado = reader.GetString("estado")
                    });
                }
            }

            return lista;
        }

        // Generar PDF dentro de /tmp (para funcionar en Railway)
        private string GenerarPDF(int idOrden, int idProveedor, decimal total, List<(int, int, decimal)> detalles)
        {
            string carpeta = "/tmp/Ordenes";
            Directory.CreateDirectory(carpeta);

            string ruta = Path.Combine(carpeta, $"orden_{idOrden}.pdf");

            using (FileStream fs = new FileStream(ruta, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (Document doc = new Document(PageSize.A4, 50, 50, 50, 50))
                {
                    PdfWriter.GetInstance(doc, fs);
                    doc.Open();

                    var titulo = new Paragraph(
                        $"ORDEN DE COMPRA #{idOrden}",
                        FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16)
                    )
                    {
                        Alignment = Element.ALIGN_CENTER
                    };

                    doc.Add(titulo);
                    doc.Add(new Paragraph("\n"));
                    doc.Add(new Paragraph($"Proveedor: {idProveedor}"));
                    doc.Add(new Paragraph($"Fecha: {DateTime.Now:dd/MM/yyyy HH:mm}"));
                    doc.Add(new Paragraph(" "));

                    PdfPTable tabla = new PdfPTable(3);
                    tabla.WidthPercentage = 100;
                    tabla.AddCell("Producto");
                    tabla.AddCell("Cantidad");
                    tabla.AddCell("Precio");

                    foreach (var d in detalles)
                    {
                        tabla.AddCell(d.Item1.ToString());
                        tabla.AddCell(d.Item2.ToString());
                        tabla.AddCell($"{d.Item3:C}");
                    }

                    doc.Add(tabla);
                    doc.Add(new Paragraph("\n"));
                    doc.Add(new Paragraph($"Total: {total:C}"));

                    doc.Close();
                }
            }

            return ruta;
        }

        // ========== ENVÍO DE CORREO CON SENDGRID ========== //
        private async void EnviarCorreo(int idOrden, int idProveedor, string rutaPDF)
        {
            try
            {
                // API key de SendGrid desde Railway
                string apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    Console.WriteLine("ERROR: SENDGRID_API_KEY no está configurada.");
                    return;
                }

                // Correo del proveedor desde la BD
                string proveedorCorreo = ObtenerCorreoProveedor(idProveedor);
                if (string.IsNullOrWhiteSpace(proveedorCorreo))
                {
                    Console.WriteLine("ERROR: El proveedor no tiene correo.");
                    return;
                }

                // Correo remitente verificado en SendGrid
                string fromEmail = Environment.GetEnvironmentVariable("FROM_EMAIL");
                if (string.IsNullOrWhiteSpace(fromEmail))
                {
                    Console.WriteLine("ERROR: FROM_EMAIL no está configurado.");
                    return;
                }

                var client = new SendGridClient(apiKey);

                var from = new EmailAddress(fromEmail, "Sistema de Órdenes");
                var to = new EmailAddress(proveedorCorreo);

                string subject = $"Orden de Compra #{idOrden}";
                string plainTextContent = "Adjunto la orden de compra generada automáticamente.";
                string htmlContent = "<p>Adjunto la orden de compra generada automáticamente.</p>";

                var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);

                // Adjuntar PDF
                if (File.Exists(rutaPDF))
                {
                    byte[] fileBytes = File.ReadAllBytes(rutaPDF);
                    string fileBase64 = Convert.ToBase64String(fileBytes);
                    msg.AddAttachment($"orden_{idOrden}.pdf", fileBase64);
                }

                var response = await client.SendEmailAsync(msg);

                string responseBody = await response.Body.ReadAsStringAsync();

                Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");
                Console.WriteLine("RESPUESTA SENDGRID:");
                Console.WriteLine(responseBody);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al enviar correo con SendGrid: " + ex.Message);
            }
        }

        // Obtener correo del proveedor
        private string ObtenerCorreoProveedor(int idProveedor)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT correo FROM proveedores WHERE id=@id",
                    con
                );

                cmd.Parameters.AddWithValue("@id", idProveedor);

                return cmd.ExecuteScalar()?.ToString() ?? "";
            }
        }
    }
}
