using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;
using ModuloWeb.ENTITIES;

namespace ModuloWeb.BROKER
{
    public class OrdenCompraBroker
    {
        // ==========================
        //  Helper: crear conexión
        // ==========================
        private MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");

            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);

            // Desarrollo local
            return ConexionBD.Conectar();
        }

        // ==========================
        // Insertar ENCABEZADO orden
        // ==========================
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

        // ==========================
        // Insertar DETALLE orden
        // ==========================
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

        // ==========================
        // Listar proveedores
        // ==========================
        public List<Proveedor> ObtenerProveedores()
        {
            var lista = new List<Proveedor>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id, nombre, nit, correo, telefono, direccion FROM proveedores",
                    con
                );

                var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    lista.Add(new Proveedor
                    {
                        Id        = reader.GetInt32("id"),
                        Nombre    = reader.GetString("nombre"),
                        Nit       = reader["nit"] != DBNull.Value ? reader.GetString("nit") : "",
                        Correo    = reader.GetString("correo"),
                        Telefono  = reader.GetString("telefono"),
                        Direccion = reader.GetString("direccion")
                    });
                }
            }

            return lista;
        }

        // ==========================
        // Obtener UN proveedor
        // ==========================
        public Proveedor? ObtenerProveedorPorId(int idProveedor)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id, nombre, nit, correo, telefono, direccion " +
                    "FROM proveedores WHERE id = @id",
                    con
                );

                cmd.Parameters.AddWithValue("@id", idProveedor);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        return null;

                    return new Proveedor
                    {
                        Id        = reader.GetInt32("id"),
                        Nombre    = reader.GetString("nombre"),
                        Nit       = reader["nit"] != DBNull.Value ? reader.GetString("nit") : "",
                        Correo    = reader.GetString("correo"),
                        Telefono  = reader.GetString("telefono"),
                        Direccion = reader.GetString("direccion")
                    };
                }
            }
        }

        // ==========================
        // Listar productos
        // ==========================
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
                        Id          = reader.GetInt32("id"),
                        Nombre      = reader.GetString("nombre"),
                        Precio      = reader.GetDecimal("precio"),
                        IdProveedor = reader.GetInt32("id_proveedor")
                    });
                }
            }

            return lista;
        }

        // ==========================
        // Obtener UN producto
        // ==========================
        public Producto? ObtenerProductoPorId(int idProducto)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id, nombre, precio, id_proveedor " +
                    "FROM productos WHERE id = @id",
                    con
                );

                cmd.Parameters.AddWithValue("@id", idProducto);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        return null;

                    return new Producto
                    {
                        Id          = reader.GetInt32("id"),
                        Nombre      = reader.GetString("nombre"),
                        Precio      = reader.GetDecimal("precio"),
                        IdProveedor = reader.GetInt32("id_proveedor")
                    };
                }
            }
        }

        // ==========================
        // Precio de producto
        // ==========================
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

        // ==========================
        // Correo de proveedor (para enviar)
        // ==========================
        public string? ObtenerCorreoProveedor(int idProveedor)
        {
            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT correo FROM proveedores WHERE id = @id",
                    con
                );

                cmd.Parameters.AddWithValue("@id", idProveedor);

                return cmd.ExecuteScalar()?.ToString();
            }
        }


        // --------------------------------------------------------------
        // LISTAR TODAS LAS ÓRDENES (para trazabilidad)
        // --------------------------------------------------------------
        public List<OrdenCompra> ObtenerOrdenes()
        {
            var lista = new List<OrdenCompra>();

            using (var con = CrearConexion())
            {
                con.Open();

                var cmd = new MySqlCommand(
                    "SELECT id_orden, id_proveedor, total, fecha, estado " +
                    "FROM ordenes_compra ORDER BY fecha DESC",
                    con
                );

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        lista.Add(new OrdenCompra
                        {
                            IdOrden     = reader.GetInt32("id_orden"),
                            IdProveedor = reader.GetInt32("id_proveedor"),
                            Total       = reader.GetDecimal("total"),
                            Fecha       = reader.GetDateTime("fecha"),
                            Estado      = reader.GetString("estado")
                        });
                    }
                }
            }

            return lista;
        }
    }
}
