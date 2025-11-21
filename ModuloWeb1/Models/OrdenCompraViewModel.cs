namespace ModuloWeb1.Models
{
    public class OrdenCompraViewModel
    {
        public int IdProveedor { get; set; }
        public List<int> IdProducto { get; set; } = new();
        public List<int> Cantidad { get; set; } = new();
    }
}
