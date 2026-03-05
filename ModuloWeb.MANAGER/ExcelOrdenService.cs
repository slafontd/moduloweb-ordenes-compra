using ClosedXML.Excel;
using ModuloWeb.ENTITIES;

namespace ModuloWeb.MANAGER
{
    public class OrdenExcelDto
    {
        public string Condiciones     { get; set; } = "";
        public string Moneda          { get; set; } = "COP";
        public string Comprador       { get; set; } = "";
        public string EntregarA       { get; set; } = "";
        public string EntregarAlterno { get; set; } = "NA";
    }

    public class DetalleExcelDto
    {
        public string  NombreManual   { get; set; } = "";
        public string  Item           { get; set; } = "";
        public string  Catalogo       { get; set; } = "";
        public string  Modelo         { get; set; } = "";
        public string  Descripcion    { get; set; } = "";
        public string  FechaEntrega   { get; set; } = "";
        public decimal Iva            { get; set; } = 0;
        public int     Cantidad       { get; set; } = 1;
        public string  Um             { get; set; } = "UND";
        public decimal PrecioUnitario { get; set; } = 0;
        public decimal Descuento      { get; set; } = 0;
    }

    public class ExcelOrdenService
    {
        private readonly string _rutaPlantilla;
        private const int IMG_W = 165; // 4.37 cm a 96 dpi
        private const int IMG_H = 94;  // 2.50 cm a 96 dpi

        public ExcelOrdenService(string rutaPlantilla)
        {
            _rutaPlantilla = rutaPlantilla;
        }

        public byte[] GenerarExcel(
            int    idOrden,
            string numeroOrden,
            Proveedor proveedor,
            DateTime  fecha,
            OrdenExcelDto cabezal,
            List<DetalleExcelDto> detalles)
        {
            using var wb = new XLWorkbook(_rutaPlantilla);

            FixImageSize(wb);
            LlenarInstancia(wb, idOrden, numeroOrden, proveedor, fecha, cabezal);
            LlenarProductos(wb, detalles);

            // Ocultar todas las hojas excepto Hoja1 → el PDF solo mostrará la orden
            foreach (var hoja in wb.Worksheets)
                if (hoja.Name != "Hoja1")
                    hoja.Visibility = XLWorksheetVisibility.Hidden;

            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return stream.ToArray();
        }

        // ── Logo ─────────────────────────────────────────────────────────────
        private void FixImageSize(XLWorkbook wb)
        {
            foreach (var pic in wb.Worksheet("Hoja1").Pictures)
                pic.WithSize(IMG_W, IMG_H);
        }

        // ── Cabezal ──────────────────────────────────────────────────────────
        private void LlenarInstancia(
            XLWorkbook wb, int idOrden, string numeroOrden,
            Proveedor proveedor, DateTime fecha, OrdenExcelDto cab)
        {
            var ws = wb.Worksheet("Instancia");

            ws.Cell("B2").Value = numeroOrden;          // Número de orden legible
            ws.Cell("C2").Value = proveedor.Contacto;
            ws.Cell("D2").Value = proveedor.Correo;
            ws.Cell("E2").Value = fecha.ToString("dd/MM/yyyy");
            ws.Cell("F2").Value = cab.Moneda;
            ws.Cell("G2").Value = cab.EntregarA;
            ws.Cell("H2").Value = cab.EntregarAlterno;
            ws.Cell("I2").Value = cab.Condiciones;
            ws.Cell("J2").Value = "NA";
            ws.Cell("K2").Value = "NA";
            ws.Cell("L2").Value = proveedor.Nombre;
            ws.Cell("M2").Value = proveedor.Nit;
            ws.Cell("N2").Value = proveedor.Ciudad;
            ws.Cell("O2").Value = proveedor.Direccion;
            ws.Cell("P2").Value = cab.Comprador;
        }

        // ── Productos con bordes ──────────────────────────────────────────────
        private void LlenarProductos(XLWorkbook wb, List<DetalleExcelDto> detalles)
        {
            var ws = wb.Worksheet("Hoja1");
            const int FILA_BASE  = 19;
            const int COL_INICIO = 2;   // B
            const int COL_FIN    = 14;  // N

            for (int i = 0; i < detalles.Count; i++)
            {
                int fila = FILA_BASE + i;
                var d = detalles[i];

                ws.Cell(fila, 2).Value  = i + 1;
                ws.Cell(fila, 4).Value  = d.Item;
                ws.Cell(fila, 5).Value  = d.Catalogo;
                ws.Cell(fila, 6).Value  = d.Modelo;
                ws.Cell(fila, 7).Value  = string.IsNullOrWhiteSpace(d.Descripcion)
                                            ? d.NombreManual : d.Descripcion;
                ws.Cell(fila, 8).Value  = d.FechaEntrega;
                ws.Cell(fila, 9).Value  = d.Iva;
                ws.Cell(fila, 10).Value = d.Cantidad;
                ws.Cell(fila, 11).Value = d.Um;
                ws.Cell(fila, 12).Value = d.PrecioUnitario;
                ws.Cell(fila, 13).Value = d.Descuento;
                ws.Cell(fila, 14).FormulaA1 =
                    $"=J{fila}*L{fila}*(1-(M{fila}/100))*(1+(I{fila}/100))";

                // Formato números
                ws.Cell(fila, 12).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(fila, 14).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(fila, 2).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
                ws.Cell(fila, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // ── Bordes de la fila ───────────────────────────────────────
                AplicarBordesFila(ws, fila, COL_INICIO, COL_FIN);
            }

            // Fila TOTAL
            int filaTot = FILA_BASE + detalles.Count;
            ws.Cell(filaTot, 13).Value = "TOTAL:";
            ws.Cell(filaTot, 13).Style.Font.Bold = true;
            ws.Cell(filaTot, 14).FormulaA1 = $"=SUM(N{FILA_BASE}:N{filaTot - 1})";
            ws.Cell(filaTot, 14).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(filaTot, 14).Style.Font.Bold = true;

            // Borde alrededor del bloque completo (cabezal fila 18 + datos)
            var rangoTabla = ws.Range(18, COL_INICIO, filaTot, COL_FIN);
            rangoTabla.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        }

        private static void AplicarBordesFila(IXLWorksheet ws, int fila, int colIni, int colFin)
        {
            var rango = ws.Range(fila, colIni, fila, colFin);
            rango.Style.Border.TopBorder       = XLBorderStyleValues.Thin;
            rango.Style.Border.BottomBorder    = XLBorderStyleValues.Thin;
            rango.Style.Border.LeftBorder      = XLBorderStyleValues.Thin;
            rango.Style.Border.RightBorder     = XLBorderStyleValues.Thin;
            rango.Style.Border.InsideBorder    = XLBorderStyleValues.Thin;
        }
    }
}