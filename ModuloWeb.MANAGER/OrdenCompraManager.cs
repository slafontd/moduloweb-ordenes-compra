using ModuloWeb.BROKER;
using System;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using SendGrid;
using SendGrid.Helpers.Mail;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // ==============================
        // 1. Crear orden
        // ==============================
        public int CrearOrden(int idProveedor, decimal total, List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // 1. Guarda en BD
            int idOrden = broker.InsertarOrden(idProveedor, total);

            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // 2. Genera Excel desde la plantilla
            string rutaExcel = GenerarExcel(idOrden, idProveedor, total, detalles);

            // 3. Envía por correo
            EnviarCorreo(idOrden, idProveedor, rutaExcel);

            return idOrden;
        }

        // ==============================
        // 2. Generar Excel desde plantilla
        // ==============================
        private string GenerarExcel(
            int idOrden,
            int idProveedor,
            decimal total,
            List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida    = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");
            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
                throw new Exception($"No se encuentra la plantilla PlantillaOrdenes.xlsx en {rutaPlantilla}.");

            var proveedor = broker.ObtenerProveedorPorId(idProveedor);

            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var hoja      = wb.Worksheet("Hoja1");
                var instancia = wb.Worksheet("Instancia");

                // ==== Encabezado visible en Hoja1 ====
                hoja.Cell("H6").Value = idOrden;                       // Nº de orden
                hoja.Cell("H8").Value = DateTime.Now;                 // Fecha
                hoja.Cell("C5").Value = total;                        // Total general (si usas esa celda)

                // ==== Encabezado en hoja Instancia ====
                instancia.Cell("J2").Value = idOrden;                 // Consecutivo
                instancia.Cell("H2").Value = total;                   // Total
                instancia.Cell("I2").Value = "30 días";               // Condición de pago (ejemplo)

                if (proveedor != null)
                {
                    instancia.Cell("L2").Value = proveedor.Nombre;    // Proveedor
                    instancia.Cell("M2").Value = proveedor.Nit;       // NIT
                    instancia.Cell("O2").Value = proveedor.Direccion; // Dirección
                    instancia.Cell("P2").Value = proveedor.Correo;    // Contacto / correo
                    instancia.Cell("Q2").Value = proveedor.Telefono;  // Teléfono
                    instancia.Cell("B2").Value = proveedor.Nombre;    // Texto auxiliar
                }

                // ==== Detalles en la tabla de Hoja1 (filas 20 en adelante) ====
                int fila = 20;
                int item = 1;

                foreach (var d in detalles)
                {
                    var prod = broker.ObtenerProductoPorId(d.idProducto);

                    hoja.Cell(fila, 2).Value = item;                                // B: ítem
                    hoja.Cell(fila, 3).Value = prod?.Id ?? d.idProducto;            // C: código
                    hoja.Cell(fila, 4).Value = prod?.Nombre ?? $"Producto {d.idProducto}"; // D: descripción
                    hoja.Cell(fila, 5).Value = d.cantidad;                          // E: cantidad
                    hoja.Cell(fila, 6).Value = "UND";                               // F: unidad
                    hoja.Cell(fila, 7).Value = d.precio;                             // G: valor unitario
                    hoja.Cell(fila, 8).Value = d.cantidad * d.precio;               // H: valor total

                    fila++;
                    item++;
                }

                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // ==============================
        // 3. Enviar correo con SendGrid
        // ==============================
        private void EnviarCorreo(int idOrden, int idProveedor, string archivo)
        {
            string? correoDestino = broker.ObtenerCorreoProveedor(idProveedor);

            if (string.IsNullOrWhiteSpace(correoDestino))
            {
                Console.WriteLine("El proveedor no tiene correo configurado. No se envía correo.");
                return;
            }

            // Tomamos remitente y API key desde variables de entorno
            string? fromEmail = Environment.GetEnvironmentVariable("FROM_EMAIL");
            string? apiKey    = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");

            if (string.IsNullOrWhiteSpace(fromEmail) || string.IsNullOrWhiteSpace(apiKey))
            {
                Console.WriteLine("Falta FROM_EMAIL o SENDGRID_API_KEY en variables de entorno. No se envía correo.");
                return;
            }

            var client  = new SendGridClient(apiKey);
            var from    = new EmailAddress(fromEmail, "Sistema de Órdenes");
            var to      = new EmailAddress(correoDestino);
            var subject = $"Orden de Compra #{idOrden}";
            var plain   = "Adjunto la orden de compra generada automáticamente.";

            var msg = MailHelper.CreateSingleEmail(from, to, subject, plain, null);

            // Adjuntamos el Excel generado
            byte[] archivoBytes  = File.ReadAllBytes(archivo);
            string archivoBase64 = Convert.ToBase64String(archivoBytes);

            msg.AddAttachment(Path.GetFileName(archivo), archivoBase64);

            var response = client.SendEmailAsync(msg).Result;

            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");
        }
    }
}

