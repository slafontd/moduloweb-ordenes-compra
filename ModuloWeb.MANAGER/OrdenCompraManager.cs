using ModuloWeb.BROKER;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using iTextSharp.text;
using iTextSharp.text.pdf;
using ModuloWeb.ENTITIES;
using MySql.Data.MySqlClient;
using System.Collections.Generic;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // Helper para crear la conexión a MySQL
        // - En producción (Railway): usa la variable de entorno ConnectionStrings__DefaultConnection
        // - En desarrollo local: si no está definida, usa ConexionBD.Conectar()
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
            {
                return new MySqlConnection(cs);
            }

            // Fallback para cuando trabajas local con tu MySQL de siempre
            return ConexionBD.Conectar();
        }

        // Crea una nueva orden de compra
        public int CrearOrden(int idProveedor, decimal total, List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // Inserta encabezado de la orden
            int idOrden = broker.InsertarOrden(idProveedor, total);

            // Inserta detalle por cada producto
            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // Genera el PDF y envía correo
            string rutaPDF = GenerarPDF(idOrden, idProveedor, total, detalles);
            EnviarCorreo(idOrden, idProveedor, rutaPDF);

            return idOrden;
        }

        // Obtiene todas las órdenes de compra
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

        // Genera el PDF con los datos de la orden
        private string GenerarPDF(int idOrden, int idProveedor, decimal total, List<(int, int, decimal)> detalles)
        {
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
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

        // Enviar correo con Gmail (requiere contraseña de aplicación)
        private void EnviarCorreo(int idOrden, int idProveedor, string rutaPDF)
        {
            string proveedorCorreo = ObtenerCorreoProveedor(idProveedor);
            string remitente = "lafontdiazsantiago@gmail.com"; // tu correo Gmail
            string claveApp = "jeae szgh fkff fzyz";            // contraseña de aplicación de Gmail

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(remitente, "Sistema de Órdenes");
                mail.To.Add(proveedorCorreo);
                mail.Subject = $"Orden de Compra #{idOrden}";
                mail.Body = "Adjunto la orden de compra generada automáticamente.";
                mail.Attachments.Add(new Attachment(rutaPDF));

                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(remitente, claveApp);

                    try
                    {
                        smtp.Send(mail);
                        Console.WriteLine("Correo enviado correctamente a " + proveedorCorreo);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error al enviar correo: " + ex.Message);
                    }
                }
            }
        }

        // Obtiene el correo del proveedor
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

                return cmd.ExecuteScalar()?.ToString() ?? "sin_correo@empresa.com";
            }
        }
    }
}

