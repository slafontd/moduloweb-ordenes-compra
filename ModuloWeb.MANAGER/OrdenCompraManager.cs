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

        
        // 0. Helper: conexión (Railway)
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);

            // modo local
            return ConexionBD.Conectar();
        }

    
        // 1. Crear ORDEN + DETALLES + EXCEL + MAIL 
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

            // 3. Generar Excel leyendo TODO desde BD
            string rutaExcel = GenerarExcel(idOrden);

            // 4. Enviar correo al proveedor con el Excel adjunto
            EnviarCorreo(idOrden, idProveedor, rutaExcel);

            return idOrden;
        }

       
        // 2. Listar órdenes desde BD
        public List<OrdenCompra> ObtenerOrdenes()
        {
            return broker.ObtenerOrdenes();
        }

        
        // 3. Generar EXCEL a partir de lo guardado      
        private string GenerarExcel(int idOrden)
        {
            //3.1. Leer encabezado + datos del proveedor 
            string proveedorNombre = "";
            string proveedorNit = "";
            string proveedorDireccion = "";
            string proveedorTelefono = "";
            decimal total = 0;
            DateTime fecha = DateTime.Now;

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(@"
                    SELECT  o.id_orden,
                            o.total,
                            o.fecha,
                            p.nombre     AS proveedor_nombre,
                            p.nit        AS proveedor_nit,
                            p.direccion  AS proveedor_direccion,
                            p.telefono   AS proveedor_telefono
                    FROM ordenes_compra o
                    JOIN proveedores p ON p.id = o.id_proveedor
                    WHERE o.id_orden = @id;", con);

                cmd.Parameters.AddWithValue("@id", idOrden);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        throw new Exception($"No se encontró la orden {idOrden} en la base de datos.");

                    proveedorNombre    = reader.GetString("proveedor_nombre");
                    proveedorNit       = reader.IsDBNull(reader.GetOrdinal("proveedor_nit"))
                                            ? ""
                                            : reader.GetString("proveedor_nit");
                    proveedorDireccion = reader.IsDBNull(reader.GetOrdinal("proveedor_direccion"))
                                            ? ""
                                            : reader.GetString("proveedor_direccion");
                    proveedorTelefono  = reader.IsDBNull(reader.GetOrdinal("proveedor_telefono"))
                                            ? ""
                                            : reader.GetString("proveedor_telefono");

                    total = reader.GetDecimal("total");
                    fecha = reader.GetDateTime("fecha");
                }
            }

            //3.2. Leer detalles con nombre de producto
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

            //  3.3. Preparar rutas 
            string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
            Directory.CreateDirectory(carpeta);

            string rutaSalida = Path.Combine(carpeta, $"Orden_{idOrden}.xlsx");

            string rutaPlantilla = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Plantillas",
                "PlantillaOrdenes.xlsx");

            if (!File.Exists(rutaPlantilla))
                throw new Exception($"No se encuentra la plantilla PlantillaOrdenes.xlsx en {rutaPlantilla}.");

            //3.4. Abrir plantilla y rellenar
            using (var wb = new XLWorkbook(rutaPlantilla))
            {
                var ws = wb.Worksheet(1);

                // 3.4.1 ENCABEZADO IZQUIERDO (bloque “Proveedor”)
                ws.Cell("B8").Value  = "Proveedor: " + proveedorNombre;
                ws.Cell("B9").Value  = "NIT/CC: " + proveedorNit;
                ws.Cell("B10").Value = "Ciudad:"; 
                ws.Cell("B11").Value = "Dirección: " + proveedorDireccion;
                ws.Cell("B12").Value = "Contacto:"; 
                ws.Cell("B13").Value = proveedorTelefono;

                // 3.4.2 ENCABEZADO DERECHA (caja gris superior derecha)
                // ejemplo plantilla:
                //  J8:  "Facturar a:  SUPLINDUSTRIA S.A.S."
                //  J9:  "Nit/CC:    901.130.635-2     "
                //  J10: "Ciudad:      Medellin       Depto:    Antioquia"
                //  J11: "Direccion:  CRR. 56 # 29-60"
                //  J12: "Telefonos:  1)604 4445669   2) 6859378 "

                ws.Cell("J8").Value  = "Facturar a:  " + proveedorNombre;
                ws.Cell("J9").Value  = "Nit/CC:    " + proveedorNit;
                ws.Cell("J10").Value = "Ciudad: "; 
                ws.Cell("J11").Value = "Direccion:  " + proveedorDireccion;
                ws.Cell("J12").Value = "Telefonos:  " + proveedorTelefono;

                // 3.4.3 OTROS CAMPOS (fecha, total, etc.)
                // Fecha de la orden (celda donde se ve la fecha en tu plantilla)
                ws.Cell("B15").Value = fecha;       
                ws.Cell("C15").Value = "COP";       // Moneda fija

                ws.Cell("B16").Value = "Condiciones de pago: 30 días"; // o lo que venga del formulario
                // total (celda de “Total” en la parte baja)
                ws.Cell("L27").Value = total;

                // 3.4.4 DETALLES (líneas de productos)
                int fila = 20;  // primera fila de la tabla de ítems 

                int linea = 1;
                foreach (var d in detalles)
                {
                    // Ln
                    ws.Cell(fila, 1).Value = linea;
                    // Item (se puede usar id de producto)
                    ws.Cell(fila, 2).Value = d.idProducto;
                    // Catalogo
                    ws.Cell(fila, 3).Value = ""; // si no tienes catálogo queda vacio
                    // Modelo
                    ws.Cell(fila, 4).Value = ""; // idem
                    // Descripción
                    ws.Cell(fila, 5).Value = d.nombre;
                    // F. Entrega, % IVA, UM
                    // Cantidad
                    ws.Cell(fila, 8).Value = d.cantidad;
                    // Precio Unit.
                    ws.Cell(fila, 10).Value = d.precio;
                    // Valor Total (subtotal)
                    ws.Cell(fila, 12).Value = d.subtotal;

                    fila++;
                    linea++;
                }

                wb.SaveAs(rutaSalida);
            }

            return rutaSalida;
        }

        
        // 4. Enviar correo con SendGrid (API)
        
        private void EnviarCorreo(int idOrden, int idProveedor, string rutaExcel)
        {
            // 1. Correo del proveedor
            string correoDestino = broker.ObtenerCorreoProveedor(idProveedor);

            if (string.IsNullOrWhiteSpace(correoDestino))
            {
                Console.WriteLine("Proveedor sin correo, no se envía email.");
                return;
            }

            // 2. From y API key desde variables de entorno (Railway → VARIABLES)
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

            // 3. Adjuntar el Excel
            byte[] bytes  = File.ReadAllBytes(rutaExcel);
            string base64 = Convert.ToBase64String(bytes);

            msg.AddAttachment(
                Path.GetFileName(rutaExcel),
                base64,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            );

            // 4. Enviar y LOG de la respuesta
            var response = client.SendEmailAsync(msg).Result;
            Console.WriteLine($"STATUS SENDGRID: {response.StatusCode}");

            var body = response.Body.ReadAsStringAsync().Result;
            Console.WriteLine($"SENDGRID BODY: {body}");
        }
    }
}
