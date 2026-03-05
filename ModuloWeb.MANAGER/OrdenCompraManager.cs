using ModuloWeb.BROKER;
using ModuloWeb.ENTITIES;
using MySql.Data.MySqlClient;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.IO;

namespace ModuloWeb.MANAGER
{
    public class OrdenCompraManager
    {
        private readonly OrdenCompraBroker broker = new OrdenCompraBroker();

        public int CrearOrdenConPDF(
            int idProveedor,
            decimal total,
            string condiciones,
            List<(int? idProducto, string? nombreManual, int cantidad, decimal precio)> detalles)
        {
            int idOrden = broker.InsertarOrden(idProveedor, total);

            foreach (var d in detalles)
            {
                if (d.idProducto.HasValue)
                {
                    broker.InsertarDetalle(idOrden, d.idProducto.Value, d.cantidad, d.precio);
                }
                else
                {
                    InsertarDetalleManual(idOrden, d.nombreManual ?? "Producto sin nombre", d.cantidad, d.precio);
                }
            }

            GenerarPDF(idOrden, condiciones);
            return idOrden;
        }

        private void InsertarDetalleManual(int idOrden, string nombreProducto, int cantidad, decimal precio)
        {
            using (var con = broker.CrearConexion())
            {
                con.Open();
                var subtotal = cantidad * precio;
                
                var cmd = new MySqlCommand(
                    "INSERT INTO detalle_orden (id_orden, id_producto, cantidad, precio, subtotal, nombre_producto_manual) " +
                    "VALUES (@orden, NULL, @cant, @precio, @sub, @nombre);", con
                );
                
                cmd.Parameters.AddWithValue("@orden", idOrden);
                cmd.Parameters.AddWithValue("@cant", cantidad);
                cmd.Parameters.AddWithValue("@precio", precio);
                cmd.Parameters.AddWithValue("@sub", subtotal);
                cmd.Parameters.AddWithValue("@nombre", nombreProducto);
                
                cmd.ExecuteNonQuery();
            }
        }

        private void GenerarPDF(int idOrden, string condiciones)
        {
            string proveedorNombre = "";
            string proveedorNit = "";
            string proveedorDireccion = "";
            string proveedorTelefono = "";
            decimal total = 0;
            DateTime fecha = DateTime.Now;

            using (var con = broker.CrearConexion())
            {
                con.Open();
                var cmd = new MySqlCommand(@"
                    SELECT o.id_orden, o.total, o.fecha,
                           p.nombre, p.nit, p.direccion, p.telefono
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id", con);
                
                cmd.Parameters.AddWithValue("@id", idOrden);
                
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        proveedorNombre = reader.GetString("nombre");
                        proveedorNit = reader["nit"] != DBNull.Value ? reader.GetString("nit") : "";
                        proveedorDireccion = reader["direccion"] != DBNull.Value ? reader.GetString("direccion") : "";
                        proveedorTelefono = reader["telefono"] != DBNull.Value ? reader.GetString("telefono") : "";
                        total = reader.GetDecimal("total");
                        fecha = reader.GetDateTime("fecha");
                    }
                }
            }

            var detalles = new List<(string nombre, int cantidad, decimal precio, decimal subtotal)>();
            
            using (var con = broker.CrearConexion())
            {
                con.Open();
                var cmd = new MySqlCommand(@"
                    SELECT 
                        COALESCE(pr.nombre, d.nombre_producto_manual) as nombre,
                        d.cantidad, d.precio, d.subtotal
                    FROM detalle_orden d
                    LEFT JOIN productos pr ON pr.id = d.id_producto
                    WHERE d.id_orden = @id", con);
                
                cmd.Parameters.AddWithValue("@id", idOrden);
                
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        detalles.Add((
                            reader.GetString("nombre"),
                            reader.GetInt32("cantidad"),
                            reader.GetDecimal("precio"),
                            reader.GetDecimal("subtotal")
                        ));
                    }
                }
            }

            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);
            string rutaPDF = Path.Combine(carpeta, $"Orden_{idOrden}.pdf");

            using (var writer = new PdfWriter(rutaPDF))
            using (var pdf = new PdfDocument(writer))
            {
                var document = new Document(pdf);

                // TÍTULO
                var titulo = new Paragraph($"ORDEN DE COMPRA #{idOrden}")
                    .SetFontSize(20)
                    .SetTextAlignment(TextAlignment.CENTER);
                document.Add(titulo);

                document.Add(new Paragraph($"Fecha: {fecha:dd/MM/yyyy HH:mm}")
                    .SetTextAlignment(TextAlignment.RIGHT)
                    .SetFontSize(10));

                document.Add(new Paragraph("\n"));

                // PROVEEDOR
                var tablaProveedor = new Table(2).UseAllAvailableWidth();
                tablaProveedor.AddCell(new Cell().Add(new Paragraph("PROVEEDOR:")));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph(proveedorNombre)));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph("NIT:")));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph(proveedorNit)));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph("Direccion:")));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph(proveedorDireccion)));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph("Telefono:")));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph(proveedorTelefono)));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph("Condiciones:")));
                tablaProveedor.AddCell(new Cell().Add(new Paragraph(condiciones)));
                
                document.Add(tablaProveedor);
                document.Add(new Paragraph("\n"));

                // PRODUCTOS
                var tablaProductos = new Table(new float[] { 1, 5, 2, 3, 3 }).UseAllAvailableWidth();
                
                tablaProductos.AddHeaderCell(new Cell().Add(new Paragraph("#")));
                tablaProductos.AddHeaderCell(new Cell().Add(new Paragraph("Producto")));
                tablaProductos.AddHeaderCell(new Cell().Add(new Paragraph("Cantidad")));
                tablaProductos.AddHeaderCell(new Cell().Add(new Paragraph("Precio Unit.")));
                tablaProductos.AddHeaderCell(new Cell().Add(new Paragraph("Subtotal")));

                int linea = 1;
                foreach (var d in detalles)
                {
                    tablaProductos.AddCell(new Cell().Add(new Paragraph(linea.ToString())));
                    tablaProductos.AddCell(new Cell().Add(new Paragraph(d.nombre)));
                    tablaProductos.AddCell(new Cell().Add(new Paragraph(d.cantidad.ToString())));
                    tablaProductos.AddCell(new Cell().Add(new Paragraph($"${d.precio:N2}")));
                    tablaProductos.AddCell(new Cell().Add(new Paragraph($"${d.subtotal:N2}")));
                    linea++;
                }

                document.Add(tablaProductos);
                document.Add(new Paragraph("\n"));

                // TOTAL
                var parrafoTotal = new Paragraph($"TOTAL: ${total:N2}")
                    .SetFontSize(16)
                    .SetTextAlignment(TextAlignment.RIGHT);
                
                document.Add(parrafoTotal);

                document.Add(new Paragraph("\n\n"));
                document.Add(new Paragraph("_________________________")
                    .SetTextAlignment(TextAlignment.CENTER));
                document.Add(new Paragraph("Firma y Sello")
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetFontSize(10));

                document.Close();
            }

            Console.WriteLine($"✅ PDF generado: {rutaPDF}");
        }

        public List<OrdenCompra> ObtenerOrdenes()
        {
            return broker.ObtenerOrdenes();
        }
    }
}