namespace ModuloWeb.ENTITIES
{
    public class DetalleOrden
    {
        public int Id { get; set; }
        public int IdOrden { get; set; }
        public int IdProducto { get; set; }
        public int Cantidad { get; set; }
        public decimal Precio { get; set; }
        public decimal Subtotal { get; set; }

        // Propiedad opcional para mostrar el nombre del producto
        public string NombreProducto { get; set; }
    }
}
