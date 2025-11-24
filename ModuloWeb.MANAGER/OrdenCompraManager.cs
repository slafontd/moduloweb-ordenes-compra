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

        // =====================================================================
        // 0. Helper: crea conexión (usa cadena de Railway o la local de siempre)
        // =====================================================================
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);

            // modo local
            return ConexionBD.Conectar();
        }

        // =====================================================================
        // 1. CREA ORDEN: inserta encabezado + detalles + genera Excel + envía mail
        // =====================================================================
        public int CrearOrden(
            int idProveedor,
            decimal total,
            List<(int idProducto, int cantidad, decimal precio)> detalles)
        {
            // 1. Insertar encabezado
            int idOrden = broker.InsertarOrden(idProveedor, total);

            // 2. Insertar detalles
            foreach (var d in detalles)
                broker.InsertarDetalle(idOrden, d.idProducto, d.cantidad, d.precio);

            // 3. Generar el Excel a partir de lo que quedó en BD
            string rutaExcel = GenerarExcel(idOrden, idProveedor);

            // 4. Enviar por correo
            EnviarCorreo(idOrden, idProveedor, rutaExcel);

            return idOrden;
        }

        // Lista de órdenes para la pantalla de trazabilidad
        public List<OrdenCompra> ObtenerOrdenes()
        {
            return broker.ObtenerOrdenes();
        }

        // =====================================================================
        // 2. GENERAR EXCEL USANDO LA PLANTILLA
        //    - Lee encabezado y detalles desde BD
        //    - Llena la hoja Hoja1 en las posiciones que vimos en la plantilla
        //    - Calcula totales por ítem y totales de la orden
        // =====================================================================
        private string GenerarExcel(int idOrden, int idProveedor)
        {
            // --------- 2.1. Traer datos del proveedor y la orden ----------------
            string proveedorNombre = "";
            string proveedorNit = "";
            string proveedorDireccion = "";
            string proveedorTelefono = "";
            decimal totalOrden = 0;
            DateTime fechaOrden = DateTime.Now;
            string condicionesPago = "30 días";   // por ahora fijo

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  o.id_orden,
                            o.total,
                            o.fecha,
                            p.nombre      AS proveedor_nombre,
                            p.nit         AS proveedor_nit,
                            p.direccion   AS proveedor_direccion,
                            p.telefono    AS proveedor_telefono
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id AND p.id = @prov;
                ", con);

                cmd.Parameters.AddWithValue("@id", idOrden);
                cmd.Parameters.AddWithValue("@prov", idProveedor);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new Exception($"No se encontró la orden {idOrden} en la base de datos.");

                    proveedorNombre    = reader.GetString("proveedor_nombre");
                    proveedorNit       = reader.IsDBNull(reader.GetOrdinal("proveedor_nit"))
                                            ? ""
                                            : reader.GetString("proveedor_nit");
                    proveedorDireccion = reader.GetString("proveedor_direccion");
                    proveedorTelefono  = reader.GetString("proveedor_telefono");
                    totalOrden         = reader.GetDecimal("total");
                    fechaOrden         = reader.GetDateTime("fecha");
                }
            }

            // --------- 2.2. Traer detalles de la orden -------------------------
            var detalles = new List<(int idProducto, string nombre, int cantidad, decimal precio, decimal subtotal)>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  d.id_producto,
                            pr.nombre,
                            d.cantidad,
                            d.precio,
                            d.subtotal
                    FROM detalle_orden d
                    JOIN productos pr ON pr.id = d.id_producto
                    WHERE d.id_orden = @id;
                ", con);

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

            // --------- 2.3. Preparar rutas ------------------------------------
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
                throw new Exception($"No se encuentra la plantilla PlantillaOrdenes.xlsx en {rutaPlantilla}.");

            // --------- 2.4. Abrir plantilla y rellenar -------------------------
            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var ws = wb.Worksheet("Hoja1");

                // ========== ENCABEZADO (ESQUINA IZQUIERDA) =====================
                // Estas celdas son las que en tu archivo se ven:
                //  A1: Nombre proveedor
                //  A2: NIT
                //  A3: Dirección
                //  A4: Teléfono
                ws.Cell("A1").Value = proveedorNombre;
                ws.Cell("A2").Value = proveedorNit;
                ws.Cell("A3").Value = proveedorDireccion;
                ws.Cell("A4").Value = proveedorTelefono;

                // Fecha de orden (celda debajo de "Fecha Orden")
                // en la plantilla actual está en B7 (si se movió, ajusta solo esta referencia)
                ws.Cell("B7").Value = fechaOrden;

                // Condiciones de pago (debajo de "Condiciones de pago")
                ws.Cell("A9").Value = $"Condiciones de pago: {condicionesPago}";

                // ========== DETALLES (TABLA PRINCIPAL) ==========================
                // Encabezados en la fila 18:
                //  B: Ln, D: Item, G: Descripcion, J: Cantidad, L: Precio Unit., N: Valor Total
                int fila = 19;
                int linea = 1;
                decimal sumaSubtotales = 0;

                foreach (var d in detalles)
                {
                    // Ln
                    ws.Cell(fila, "B").Value = linea;

                    // Item → usamos el id del producto
                    ws.Cell(fila, "D").Value = d.idProducto;

                    // Descripción
                    ws.Cell(fila, "G").Value = d.nombre;

                    // Cantidad
                    ws.Cell(fila, "J").Value = d.cantidad;

                    // Precio unitario
                    ws.Cell(fila, "L").Value = d.precio;

                    // Valor total de la línea
                    ws.Cell(fila, "N").Value = d.subtotal;

                    sumaSubtotales += d.subtotal;

                    fila++;
                    linea++;
                }

                // ========== TOTALES (SUBTOTAL / DESC / TOTAL) ===================
                // En la plantilla:
                //   N22 → Subtotal
                //   N23 → Descuento
                //   N24 → Total
                ws.Cell("N22").Value = sumaSubtotales; // Subtotal
                ws.Cell("N23").Value = 0;              // Descuento (por ahora 0)
                ws.Cell("N24").Value = sumaSubtotales; // Total final de la orden

                // IMPORTANTE: Guardia para que si algún total venía distinto en BD,
                // lo respetes: si quieres forzar a lo que viene en BD, cambia por:
                // ws.Cell("N24").Value = totalOrden;

                // Guardar el archivo final
                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        // =====================================================================
        // 3. Enviar correo con SENDGRID (usando FROM_EMAIL y SENDGRID_API_KEY)
        // =====================================================================
        private void EnviarCorreo(int idOrden, int idProveedor, string rutaExcel)
        {
            // 1. Correo del proveedor desde la tabla proveedores
            string correoDestino = broker.ObtenerCorreoProveedor(idProveedor);

            if (string.IsNullOrWhiteSpace(correoDestino))
            {
                Console.WriteLine("Proveedor sin correo, no se envía email.");
                return;
            }

            // 2. From y API key desde variables de entorno de Railway
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

            // 3. Adjuntar el Excel generado
            byte[] bytes  = File.ReadAllBytes(rutaExcel);
            string base64 = Convert.ToBase64String(bytes);

            msg.AddAttachment(
                Path.GetFileName(rutaExcel),
                base64,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            // 4. Enviar y loguear respuesta de SendGrid
            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");

            var body = response.Body.ReadAsStringAsync().Result;
            Console.WriteLine($"SENDGRID BODY: {body}");
        }
    }
}
