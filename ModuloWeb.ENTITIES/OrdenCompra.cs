using System;

namespace ModuloWeb.ENTITIES
{
    public class OrdenCompra
    {
        public int IdOrden { get; set; }          // ← Esta propiedad es la que faltaba
        public int IdProveedor { get; set; }
        public decimal Total { get; set; }
        public DateTime Fecha { get; set; }
        public string Estado { get; set; } = string.Empty;
    }
}
