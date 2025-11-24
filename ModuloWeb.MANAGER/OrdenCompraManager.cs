using ModuloWeb.BROKER;
using System;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using SendGrid;
using SendGrid.Helpers.Mail;
using ModuloWeb.ENTITIES;
using MySql.Data.MySqlClient;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        // Helper para conexión directa (solo lo usa ObtenerOrdenes)
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);

            // Local
            return ConexionBD.Conectar();
        }

        // =====================================================
        // 1. CREAR ORDEN
        // =====================================================
        public int CrearOrden(int idProveedor, decimal total,
                              List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // Guardar encabezado
            int idOrden = broker.InsertarOrden(idProveedor, total);

            // Guardar detalles
            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // Generar EXCEL
            string rutaExcel = GenerarExcel(idOrden, idProveedor, total, detalles);

            // Convertir a PDF sencillo
            string rutaPDF = ConvertirExcelAPdf(rutaExcel);

            // Enviar correo con PDF
            EnviarCorreo(idOrden, idProveedor, rutaPDF);

            return idOrden;
        }

        // =====================================================
        // 2. GENERAR EXCEL DESDE PLANTILLA
        // =====================================================
        private string GenerarExcel(int idOrden, int idProveedor, decimal total,
                                   List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx"
            );

            if (!File.Exists(rutaPlantilla))
                throw new Exception("No se encuentra la plantilla PlantillaOrdenes.xlsx en /Plantillas.");

            // Datos del proveedor
            var proveedor = broker.ObtenerProveedorPorId(idProveedor);

            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var ws = wb.Worksheet(1);

                // ENCABEZADO (ajusta celdas a tu plantilla)
                ws.Cell("B2").Value = idOrden;
                ws.Cell("B3").Value = proveedor?.Nombre ?? idProveedor.ToString();
                ws.Cell("B4").Value = proveedor?.Nit ?? "";
                ws.Cell("B5").Value = proveedor?.Correo ?? "";
                ws.Cell("B6").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                // DETALLES
                int fila = 10;

                foreach (var d in detalles)
                {
                    ws.Cell(fila, 1).Value = d.idProducto;
                    ws.Cell(fila, 2).Value = d.cantidad;
                    ws.Cell(fila, 3).Value = d.precio;
                    ws.Cell(fila, 4).Value = d.cantidad * d.precio;
                    fila++;
                }

                // TOTAL (ajusta la celda según tu diseño)
                ws.Cell("D7").Value = total;

                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // =====================================================
        // 3. CONVERTIR EXCEL A PDF SIMPLE
        // =====================================================
        private string ConvertirExcelAPdf(string rutaExcel)
        {
            string carpeta = Path.GetDirectoryName(rutaExcel)!;
            string rutaPDF = Path.Combine(
                carpeta,
                Path.GetFileNameWithoutExtension(rutaExcel) + ".pdf"
            );

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var ws = workbook.Worksheet(1);

                using (var fs = new FileStream(rutaPDF, FileMode.Create))
                {
                    var doc = new iTextSharp.text.Document();
                    iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);

                    doc.Open();

                    var tabla = new iTextSharp.text.pdf.PdfPTable(4);
                    tabla.WidthPercentage = 100;

                    // Encabezados de la tabla
                    tabla.AddCell("Producto");
                    tabla.AddCell("Cantidad");
                    tabla.AddCell("Precio");
                    tabla.AddCell("Subtotal");

                    int fila = 10;
                    while (!string.IsNullOrEmpty(ws.Cell(fila, 1).GetString()))
                    {
                        tabla.AddCell(ws.Cell(fila, 1).GetString());
                        tabla.AddCell(ws.Cell(fila, 2).GetString());
                        tabla.AddCell(ws.Cell(fila, 3).GetString());
                        tabla.AddCell(ws.Cell(fila, 4).GetString());
                        fila++;
                    }

                    doc.Add(tabla);
                    doc.Close();
                }
            }

            return rutaPDF;
        }

        // =====================================================
        // 4. ENVIAR CORREO (SENDGRID)
        // =====================================================
        private void EnviarCorreo(int idOrden, int idProveedor, string archivo)
        {
            string correoDestino = broker.ObtenerCorreoProveedor(idProveedor);
            
            // Tomar el remitente SOLO desde variables de entorno  
            string remitente = Environment.GetEnvironmentVariable("FROM_EMAIL");
            
            if (string.IsNullOrWhiteSpace(remitente))
            throw new Exception("La variable de entorno FROM_EMAIL no está configurada.");
            
            // API Key desde variables
            string apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");
            
            if (string.IsNullOrWhiteSpace(apiKey))
            throw new Exception("La variable de entorno SENDGRID_API_KEY no está configurada.");
            
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


        // =====================================================
        // 5. OBTENER ÓRDENES (USADO POR EL CONTROLLER)
        // =====================================================
        public List<OrdenCompra> ObtenerOrdenes()
        {
            var lista = new List<OrdenCompra>();

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
                        IdOrden     = reader.GetInt32("id_orden"),
                        IdProveedor = reader.GetInt32("id_proveedor"),
                        Total       = reader.GetDecimal("total"),
                        Fecha       = reader.GetDateTime("fecha"),
                        Estado      = reader.GetString("estado")
                    });
                }
            }

            return lista;
        }
    }
}
