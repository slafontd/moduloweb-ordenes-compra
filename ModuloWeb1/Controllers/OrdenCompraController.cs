using Microsoft.AspNetCore.Mvc;
using ModuloWeb.MANAGER;
using ModuloWeb.BROKER;
using ModuloWeb.ENTITIES;
using ModuloWeb1.Models;
using System.Collections.Generic;
using System.Linq;

namespace ModuloWeb1.Controllers
{
    public class OrdenCompraController : Controller
    {
        OrdenCompraManager manager = new OrdenCompraManager();
        OrdenCompraBroker broker = new OrdenCompraBroker(); // Aquí agregas esta línea

        // Muestra el formulario
        public IActionResult Crear()
        {
            ViewBag.Proveedores = broker.ObtenerProveedores();
            ViewBag.Productos = broker.ObtenerProductos();
            return View();
        }

        // Recibe el formulario (POST)
        [HttpPost]
        public IActionResult Crear(OrdenCompraViewModel model)
        {
            var detalles = new List<(int idProducto, int cantidad, decimal precio)>();

            foreach (var idProd in model.IdProducto)
            {
                decimal precio = broker.ObtenerPrecioProducto(idProd);
                int cantidad = model.Cantidad[model.IdProducto.IndexOf(idProd)];
                detalles.Add((idProd, cantidad, precio));
            }

            var total = detalles.Sum(d => d.cantidad * d.precio);
            manager.CrearOrden(model.IdProveedor, total, detalles);

            ViewBag.Mensaje = " Orden creada correctamente.";

            // Volvemos a llenar los combos
            ViewBag.Proveedores = broker.ObtenerProveedores();
            ViewBag.Productos = broker.ObtenerProductos();

            return View();
        }

        // Lista de órdenes
        public IActionResult Lista()
        {
            var ordenes = manager.ObtenerOrdenes();
            return View(ordenes);
        }
    }
}
