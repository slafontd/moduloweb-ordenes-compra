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
using ClosedXML.Excel;


namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // ================================
        // 1. MÉTODO PARA CREAR ORDEN
        // ================================
        public int CrearOrden(int idProveedor, decimal total, List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            int idOrden = broker.InsertarOrden(idProveedor, total);

            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // 👉 Generar archivo Excel
            string rutaExcel = GenerarExcel(idOrden, idProveedor, total, detalles);

            // 👉 Enviar por correo
            EnviarCorreo(idOrden, idProveedor, rutaExcel);

            return idOrden;
        }

        // ================================
        // 2. MÉTODO PARA GENERAR EXCEL
        // ================================
        private string GenerarExcel(int idOrden, int idProveedor, decimal total,
                                   List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            // Ruta de tu plantilla original
            string rutaPlantilla = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                                "Plantillas", "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
                throw new Exception("La plantilla PlantillaOrdenes.xlsx no existe en el servidor.");

            // Abrir plantilla
            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var ws = wb.Worksheet(1);

                // ===== RELLENAR ENCABEZADO =====
                ws.Cell("B2").Value = idOrden;
                ws.Cell("B3").Value = idProveedor;
                ws.Cell("B4").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                // ===== RELLENAR DETALLES =====
                int fila = 8;

                foreach (var d in detalles)
                {
                    ws.Cell(fila, 1).Value = d.idProducto;
                    ws.Cell(fila, 2).Value = d.cantidad;
                    ws.Cell(fila, 3).Value = d.precio;
                    ws.Cell(fila, 4).Value = d.cantidad * d.precio;
                    fila++;
                }

                // TOTAL
                ws.Cell("C5").Value = total;

                // Guardar archivo final
                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // ================================
        // 3. MÉTODO PARA ENVIAR CORREO
        // ================================
        private void EnviarCorreo(int idOrden, int idProveedor, string archivo)
        {
            string correoDestino = broker.ObtenerCorreoProveedor(idProveedor);

            string remitente = "ordenes@moduloweb.com"; 
            string apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");

            var client = new SendGrid.SendGridClient(apiKey);
            var from = new SendGrid.Helpers.Mail.EmailAddress(remitente, "Sistema de Órdenes");
            var to = new SendGrid.Helpers.Mail.EmailAddress(correoDestino);
            var subject = $"Orden de Compra #{idOrden}";
            var plainText = "Adjunto la orden de compra generada automáticamente.";

            var msg = SendGrid.Helpers.Mail.MailHelper.CreateSingleEmail(from, to, subject, plainText, null);

            // Adjuntar archivo Excel
            byte[] archivoBytes = File.ReadAllBytes(archivo);
            string archivoBase64 = Convert.ToBase64String(archivoBytes);

            msg.AddAttachment($"Orden_{idOrden}.xlsx", archivoBase64);

            var response = client.SendEmailAsync(msg).Result;

            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");
        }
    }
}
