using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;
using ModuloWeb.ENTITIES;

namespace ModuloWeb.BROKER
{
    public class OrdenCompraBroker
    {
        // --------------------------------------------------------------------
        //  Helper: crea la conexión a MySQL
        //  - En producción (Railway): usa la variable de entorno
        //      ConnectionStrings__DefaultConnection
        //  - En desarrollo local: si no existe esa variable,
        //      usa ConexionBD.Conectar() como antes
        // --------------------------------------------------------------------
        private MySqlConnection CrearConexion()
        {
            // Intenta leer la cadena de conexión desde las variables de entorno
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
            {
                return new MySqlConnection(cs);
            }

            // Fallback para desarrollo local (localhost, etc.)
            return ConexionBD.Conectar();
        }

        // --------------------------------------------------------------------
        // Inserta la orden
        // --------------------------------------------------------------------
        public int InsertarOrden(int idProveedor, decimal total)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "INSERT INTO ordenes_compra (id_proveedor, total) " +
                    "VALUES (@prov, @total); SELECT LAST_INSERT_ID();",
                    con
                );

                cmd.Parameters.AddWithValue("@prov", idProveedor);
                cmd.Parameters.AddWithValue("@total", total);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        // --------------------------------------------------------------------
        // Inserta los detalles de la orden
        // --------------------------------------------------------------------
        public void InsertarDetalle(int idOrden, int idProducto, int cantidad, decimal precio)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var subtotal = cantidad * precio;

                var cmd = new MySqlCommand(
                    "INSERT INTO detalle_orden (id_orden, id_producto, cantidad, precio, subtotal) " +
                    "VALUES (@orden, @prod, @cant, @precio, @sub);",
                    con
                );

                cmd.Parameters.AddWithValue("@orden", idOrden);
                cmd.Parameters.AddWithValue("@prod", idProducto);
                cmd.Parameters.AddWithValue("@cant", cantidad);
                cmd.Parameters.AddWithValue("@precio", precio);
                cmd.Parameters.AddWithValue("@sub", subtotal);

                cmd.ExecuteNonQuery();
            }
        }

        // --------------------------------------------------------------------
        // Obtiene todos los proveedores
        // --------------------------------------------------------------------
        public List<Proveedor> ObtenerProveedores()
        {
            var lista = new List<Proveedor>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id, nombre, correo, telefono, direccion FROM proveedores",
                    con
                );

                var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lista.Add(new Proveedor
                    {
                        Id = reader.GetInt32("id"),
                        Nombre = reader.GetString("nombre"),
                        Correo = reader.GetString("correo"),
                        Telefono = reader.GetString("telefono"),
                        Direccion = reader.GetString("direccion")
                    });
                }
            }

            return lista;
        }

        // --------------------------------------------------------------------
        // Obtiene todos los productos con su proveedor
        // --------------------------------------------------------------------
        public List<Producto> ObtenerProductos()
        {
            var lista = new List<Producto>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id, nombre, precio, id_proveedor FROM productos",
                    con
                );

                var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lista.Add(new Producto
                    {
                        Id = reader.GetInt32("id"),
                        Nombre = reader.GetString("nombre"),
                        Precio = reader.GetDecimal("precio"),
                        IdProveedor = reader.GetInt32("id_proveedor")
                    });
                }
            }

            return lista;
        }

        // --------------------------------------------------------------------
        // Obtiene el precio real del producto desde BD
        // --------------------------------------------------------------------
        public decimal ObtenerPrecioProducto(int idProducto)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT precio FROM productos WHERE id = @id",
                    con
                );

                cmd.Parameters.AddWithValue("@id", idProducto);

                object result = cmd.ExecuteScalar();

                if (result == null || result == DBNull.Value)
                    return 0;

                return Convert.ToDecimal(result);
            }
        }
    }
}
