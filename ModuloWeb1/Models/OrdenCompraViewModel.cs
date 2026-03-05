namespace ModuloWeb1.Models
{
    public class OrdenCompraViewModel
    {
        public int IdProveedor { get; set; }

        // Cabezal de la orden
        public string Condiciones   { get; set; } = "30 días";
        public string Moneda        { get; set; } = "COP";
        public string Comprador     { get; set; } = "";
        public string EntregarA     { get; set; } = "SUPLINDUSTRIA S.A.S.";
        public string EntregarAlterno { get; set; } = "NA";

        public List<DetalleProductoViewModel> Productos { get; set; } = new();
    }

    public class DetalleProductoViewModel
    {
        // Solo modo manual
        public string NombreManual  { get; set; } = "";

        // Columnas de la tabla de la plantilla
        public string Item          { get; set; } = "";
        public string Catalogo      { get; set; } = "";
        public string Modelo        { get; set; } = "";
        public string Descripcion   { get; set; } = "";
        public string FechaEntrega  { get; set; } = "";
        public decimal Iva          { get; set; } = 0;
        public int    Cantidad      { get; set; } = 1;
        public string Um            { get; set; } = "UND";
        public decimal PrecioUnitario { get; set; } = 0;
        public decimal Descuento    { get; set; } = 0;
    }

    // Para crear un proveedor nuevo desde el formulario
    public class ProveedorViewModel
    {
        public string Nombre    { get; set; } = "";
        public string Nit       { get; set; } = "";
        public string Correo    { get; set; } = "";
        public string Telefono  { get; set; } = "";
        public string Direccion { get; set; } = "";
        public string Ciudad    { get; set; } = "";
        public string Contacto  { get; set; } = "";
    }
}