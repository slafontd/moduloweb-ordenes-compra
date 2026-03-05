using Microsoft.AspNetCore.Mvc;
using ModuloWeb.MANAGER;
using ModuloWeb.BROKER;
using ModuloWeb.ENTITIES;
using ModuloWeb1.Models;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ModuloWeb1.Controllers
{
    public class OrdenCompraController : Controller
    {
        OrdenCompraManager manager = new OrdenCompraManager();
        OrdenCompraBroker  broker  = new OrdenCompraBroker();

        private string RutaPlantilla =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plantillas", "PlantillaOrdenes.xlsx");

        private string RutaCredenciales =>
    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Credenciales", "oauth-client.json");

        // ── Helpers ──────────────────────────────────────────────────────────

        /// Genera el número de orden: PrimerPalabra-Consecutivo  (ej: "Suministros-0")
        private string GenerarNumeroOrden(string nombreProveedor, int consecutivo)
        {
            // Tomar la primera palabra (separador: espacio, guión bajo, guión)
            string primera = Regex.Split(nombreProveedor.Trim(), @"[\s_\-]+")[0];
            // Limpiar caracteres no alfanuméricos
            primera = Regex.Replace(primera, @"[^a-zA-Z0-9]", "");
            if (string.IsNullOrEmpty(primera)) primera = "ORD";
            return $"{primera}-{consecutivo}";
        }

        // ── GET: Formulario ──────────────────────────────────────────────────
        public IActionResult Crear()
        {
            ViewBag.Proveedores = broker.ObtenerProveedores();
            return View();
        }

        // ── POST: Crear orden ────────────────────────────────────────────────
        [HttpPost]
        public IActionResult Crear([FromForm] string datosOrden)
        {
            try
            {
                var model = JsonSerializer.Deserialize<OrdenCompraViewModel>(datosOrden,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                if (model == null || !model.Productos.Any())
                {
                    ViewBag.Error = "⚠️ Debe agregar al menos un producto";
                    ViewBag.Proveedores = broker.ObtenerProveedores();
                    return View();
                }

                // Calcular total
                decimal total = model.Productos
                    .Where(p => p.Cantidad > 0 && p.PrecioUnitario > 0)
                    .Sum(p => p.Cantidad * p.PrecioUnitario
                              * (1 - p.Descuento / 100)
                              * (1 + p.Iva / 100));

                // ── 1. Número de orden consecutivo por proveedor ──────────────
                int consecutivo = broker.ContarOrdenesPorProveedor(model.IdProveedor);
                var proveedor   = broker.ObtenerProveedorPorId(model.IdProveedor)
                    ?? new Proveedor { Nombre = "SinProveedor" };
                string numeroOrden = GenerarNumeroOrden(proveedor.Nombre, consecutivo);

                // ── 2. Guardar en BD (método existente del Manager) ───────────
                var detallesBD = model.Productos
                    .Where(p => p.Cantidad > 0 && p.PrecioUnitario > 0)
                    .Select(p => ((int?)null, p.NombreManual, p.Cantidad, p.PrecioUnitario))
                    .ToList();

                int idOrden = manager.CrearOrdenConPDF(
                    model.IdProveedor, total, model.Condiciones, detallesBD);

                // Persistir el número de orden legible
                broker.GuardarNumeroOrden(idOrden, numeroOrden);

                // ── 3. Generar Excel ──────────────────────────────────────────
                var cabezalDto = new OrdenExcelDto
                {
                    Condiciones     = model.Condiciones    ?? "",
                    Moneda          = model.Moneda         ?? "COP",
                    Comprador       = model.Comprador      ?? "",
                    EntregarA       = model.EntregarA      ?? "SUPLINDUSTRIA S.A.S.",
                    EntregarAlterno = model.EntregarAlterno ?? "NA"
                };

                var detallesDto = model.Productos
                    .Where(p => p.Cantidad > 0 && p.PrecioUnitario > 0)
                    .Select(p => new DetalleExcelDto
                    {
                        NombreManual   = p.NombreManual   ?? "",
                        Item           = p.Item           ?? "",
                        Catalogo       = p.Catalogo       ?? "",
                        Modelo         = p.Modelo         ?? "",
                        Descripcion    = p.Descripcion    ?? "",
                        FechaEntrega   = p.FechaEntrega   ?? "",
                        Iva            = p.Iva,
                        Cantidad       = p.Cantidad,
                        Um             = p.Um             ?? "UND",
                        PrecioUnitario = p.PrecioUnitario,
                        Descuento      = p.Descuento
                    }).ToList();

                var excelService = new ExcelOrdenService(RutaPlantilla);
                byte[] excelBytes = excelService.GenerarExcel(
                    idOrden, numeroOrden, proveedor, DateTime.Now, cabezalDto, detallesDto);

                string carpeta    = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
                string xlsxPath   = Path.Combine(carpeta, $"{numeroOrden}.xlsx");
                Directory.CreateDirectory(carpeta);
                System.IO.File.WriteAllBytes(xlsxPath, excelBytes);

                // ── 4. Convertir a PDF ────────────────────────────────────────
                string pdfPath = PdfConverterService.ConvertirAPdf(xlsxPath);

                // ── 5. Subir PDF a Google Drive ───────────────────────────────
                string driveLink = "";
                try
                {
                    var driveService = new GoogleDriveService(RutaCredenciales);
                    driveLink = driveService.SubirPdfAsync(
                        pdfPath, $"{numeroOrden}.pdf").GetAwaiter().GetResult();
                }
                catch (Exception exDrive)
                {
                    // No bloquear si Drive falla; solo notificar
                    ViewBag.WarningDrive = $"⚠️ PDF generado pero no se pudo subir a Drive: {exDrive.Message}";
                }

                ViewBag.Mensaje    = $"✅ Orden {numeroOrden} creada. Total: ${total:N2}";
                ViewBag.IdOrden    = idOrden;
                ViewBag.NumOrden   = numeroOrden;
                ViewBag.DriveLink  = driveLink;
            }
            catch (Exception ex)
            {
                ViewBag.Error = $"❌ Error: {ex.Message}";
            }

            ViewBag.Proveedores = broker.ObtenerProveedores();
            return View();
        }

        // ── POST: Crear proveedor (AJAX) ─────────────────────────────────────
        [HttpPost]
        public IActionResult CrearProveedor([FromBody] ProveedorViewModel vm)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(vm.Nombre))
                    return BadRequest("El nombre es obligatorio");

                var p = new Proveedor
                {
                    Nombre    = vm.Nombre.Trim(),
                    Nit       = vm.Nit?.Trim()       ?? "",
                    Correo    = vm.Correo?.Trim()     ?? "",
                    Telefono  = vm.Telefono?.Trim()   ?? "",
                    Direccion = vm.Direccion?.Trim()  ?? "",
                    Ciudad    = vm.Ciudad?.Trim()     ?? "",
                    Contacto  = vm.Contacto?.Trim()   ?? ""
                };

                int id = broker.InsertarProveedor(p);
                return Ok(new {
                    id, nombre=p.Nombre, nit=p.Nit, direccion=p.Direccion,
                    telefono=p.Telefono, ciudad=p.Ciudad, contacto=p.Contacto, correo=p.Correo
                });
            }
            catch (Exception ex) { return BadRequest(ex.Message); }
        }

        // ── DELETE: Eliminar proveedor (AJAX) ────────────────────────────────
        [HttpPost]
        public IActionResult EliminarProveedor([FromBody] int id)
        {
            try
            {
                bool ok = broker.EliminarProveedor(id);
                if (!ok)
                    return BadRequest("No se puede eliminar: el proveedor tiene órdenes asociadas.");
                return Ok();
            }
            catch (Exception ex) { return BadRequest(ex.Message); }
        }

        // ── GET: Lista de órdenes ────────────────────────────────────────────
        public IActionResult Lista()
        {
            var ordenes = broker.ObtenerOrdenes();
            return View(ordenes);
        }

        // ── GET: Descargar PDF ───────────────────────────────────────────────
        public IActionResult DescargarExcel(int id)
        {
            try
            {
                // Buscar por id (nombre de archivo puede variar, buscar el xlsx)
                string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
                // Intentar pdf primero, luego xlsx
                string[] pdfs = Directory.GetFiles(carpeta, "*.pdf")
                    .Where(f => Path.GetFileNameWithoutExtension(f).Contains($"-"))
                    .ToArray();

                // Fallback: buscar por nombre numérico antiguo
                string xlsxFallback = Path.Combine(carpeta, $"Orden_{id}.xlsx");
                string pdfFallback  = Path.Combine(carpeta, $"Orden_{id}.pdf");

                // Buscar PDF del consecutivo (guardado con numero_orden como nombre)
                string? archivo = pdfs.FirstOrDefault()  // simplificado: mejorar si se quiere por id
                    ?? (System.IO.File.Exists(pdfFallback) ? pdfFallback : null)
                    ?? (System.IO.File.Exists(xlsxFallback) ? xlsxFallback : null);

                if (archivo == null) return NotFound("Archivo no encontrado.");

                bool esPdf = archivo.EndsWith(".pdf");
                byte[] bytes = System.IO.File.ReadAllBytes(archivo);
                string mime  = esPdf
                    ? "application/pdf"
                    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                return File(bytes, mime, Path.GetFileName(archivo));
            }
            catch (Exception ex) { return BadRequest($"Error: {ex.Message}"); }
        }

        // ── GET: Descargar por nombre de orden ───────────────────────────────
        public IActionResult DescargarPorNumero(string numero)
        {
            try
            {
                string carpeta = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ordenes");
                string pdf  = Path.Combine(carpeta, $"{numero}.pdf");
                string xlsx = Path.Combine(carpeta, $"{numero}.xlsx");

                if (System.IO.File.Exists(pdf))
                    return File(System.IO.File.ReadAllBytes(pdf), "application/pdf", $"{numero}.pdf");
                if (System.IO.File.Exists(xlsx))
                    return File(System.IO.File.ReadAllBytes(xlsx),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"{numero}.xlsx");

                return NotFound("Archivo no encontrado.");
            }
            catch (Exception ex) { return BadRequest($"Error: {ex.Message}"); }
        }
    }
}