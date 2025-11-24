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

        // ================================
        // 1. CREAR ORDEN
        // ================================
        public int CrearOrden(
            int idProveedor,
            decimal total,
            List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // Insertar encabezado
            int idOrden = broker.InsertarOrden(idProveedor, total);

            // Insertar detalles
            foreach (var d in detalles)
            {
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);
            }

            // Generar Excel a partir de la plantilla
            string rutaExcel = GenerarExcel(idOrden, idProveedor, total, detalles);

            // Enviar correo
            EnviarCorreo(idOrden, idProveedor, rutaExcel);

            return idOrden;
        }

        // ================================
        // 2. GENERAR EXCEL DESDE PLANTILLA
        // ================================
        private string GenerarExcel(
            int idOrden,
            int idProveedor,
            decimal total,
            List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // Carpeta donde se guardan las órdenes
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            // Archivo de salida
            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            // Plantilla (asegúrate de que está marcada como Content / Copy to Output)
            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
            {
                throw new Exception($"No se encuentra la plantilla PlantillaOrdenes.xlsx en: {rutaPlantilla}");
            }

            // Abrir la plantilla y rellenar
            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var ws = wb.Worksheet(1); // primera hoja

                // ===== Encabezado (ajusta celdas a tu plantilla) =====
                ws.Cell("B2").Value = idOrden;
                ws.Cell("B3").Value = idProveedor;
                ws.Cell("B4").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                ws.Cell("C5").Value = total;

                // ===== Detalles =====
                int fila = 8; // primera fila donde empiezan los productos

                foreach (var d in detalles)
                {
                    ws.Cell(fila, 1).Value = d.idProducto;            // Columna A
                    ws.Cell(fila, 2).Value = d.cantidad;              // Columna B
                    ws.Cell(fila, 3).Value = d.precio;                // Columna C
                    ws.Cell(fila, 4).Value = d.cantidad * d.precio;   // Columna D (subtotal)
                    fila++;
                }

                // Guardar el archivo final
                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // ================================
        // 3. ENVIAR CORREO CON SENDGRID
        // ================================
        private void EnviarCorreo(int idOrden, int idProveedor, string rutaExcel)
        {
            // Correo del proveedor desde BD
            string correoDestino = broker.ObtenerCorreoProveedor(idProveedor);
            if (string.IsNullOrWhiteSpace(correoDestino))
            {
                Console.WriteLine("Proveedor sin correo, no se envía email.");
                return;
            }

            // From y API key desde variables de entorno
            string remitente = Environment.GetEnvironmentVariable("FROM_EMAIL");
            if (string.IsNullOrWhiteSpace(remitente))
                throw new Exception("La variable de entorno FROM_EMAIL no está configurada.");

            string apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new Exception("La variable de entorno SENDGRID_API_KEY no está configurada.");

            var client = new SendGridClient(apiKey);

            var from = new EmailAddress(remitente, "Sistema de Órdenes");
            var to = new EmailAddress(correoDestino);

            string subject = $"Orden de Compra #{idOrden}";
            string plainText = "Adjunto la orden de compra generada automáticamente.";

            var msg = MailHelper.CreateSingleEmail(from, to, subject, plainText, null);

            // Leer el Excel generado
            byte[] bytes = File.ReadAllBytes(rutaExcel);
            string base64 = Convert.ToBase64String(bytes);

            // Adjuntar con tipo MIME de Excel
            msg.AddAttachment(
                $"Orden_{idOrden}.xlsx",
                base64,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");
        }
    }
}
