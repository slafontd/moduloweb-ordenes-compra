using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;
using ModuloWeb.ENTITIES;

namespace ModuloWeb.BROKER
{
    public class OrdenCompraBroker
    {
        public MySqlConnection CrearConexion()
        {
            var cs = Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection");
            if (!string.IsNullOrWhiteSpace(cs))
                return new MySqlConnection(cs);
            return ConexionBD.Conectar();
        }

        // ══════════════════════════════════════════════════════
        //  PROVEEDORES
        // ══════════════════════════════════════════════════════

        public int InsertarProveedor(Proveedor p)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "INSERT INTO proveedores (nombre, nit, correo, telefono, direccion, ciudad, contacto) " +
                "VALUES (@n,@nit,@c,@t,@d,@ciu,@cont); SELECT LAST_INSERT_ID();", con);
            cmd.Parameters.AddWithValue("@n",    p.Nombre);
            cmd.Parameters.AddWithValue("@nit",  p.Nit);
            cmd.Parameters.AddWithValue("@c",    p.Correo);
            cmd.Parameters.AddWithValue("@t",    p.Telefono);
            cmd.Parameters.AddWithValue("@d",    p.Direccion);
            cmd.Parameters.AddWithValue("@ciu",  p.Ciudad);
            cmd.Parameters.AddWithValue("@cont", p.Contacto);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        public bool EliminarProveedor(int id)
        {
            using var con = CrearConexion();
            con.Open();
            // Solo elimina si no tiene órdenes asociadas
            var check = new MySqlCommand(
                "SELECT COUNT(*) FROM ordenes_compra WHERE id_proveedor = @id", con);
            check.Parameters.AddWithValue("@id", id);
            int ordenes = Convert.ToInt32(check.ExecuteScalar());
            if (ordenes > 0) return false;

            var cmd = new MySqlCommand("DELETE FROM proveedores WHERE id = @id", con);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();
            return true;
        }

        public List<Proveedor> ObtenerProveedores()
        {
            var lista = new List<Proveedor>();
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "SELECT id, nombre, nit, correo, telefono, direccion, " +
                "IFNULL(ciudad,'') AS ciudad, IFNULL(contacto,'') AS contacto " +
                "FROM proveedores ORDER BY nombre", con);
            using var reader = cmd.ExecuteReader();
            while (reader.Read()) lista.Add(MapProveedor(reader));
            return lista;
        }

        public Proveedor? ObtenerProveedorPorId(int id)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "SELECT id, nombre, nit, correo, telefono, direccion, " +
                "IFNULL(ciudad,'') AS ciudad, IFNULL(contacto,'') AS contacto " +
                "FROM proveedores WHERE id = @id", con);
            cmd.Parameters.AddWithValue("@id", id);
            using var reader = cmd.ExecuteReader();
            return reader.Read() ? MapProveedor(reader) : null;
        }

        private Proveedor MapProveedor(MySqlDataReader r) => new Proveedor
        {
            Id        = r.GetInt32("id"),
            Nombre    = r.GetString("nombre"),
            Nit       = r["nit"] != DBNull.Value ? r.GetString("nit") : "",
            Correo    = r.GetString("correo"),
            Telefono  = r.GetString("telefono"),
            Direccion = r.GetString("direccion"),
            Ciudad    = r.GetString("ciudad"),
            Contacto  = r.GetString("contacto")
        };

        // ══════════════════════════════════════════════════════
        //  ÓRDENES
        // ══════════════════════════════════════════════════════

        /// Cuenta cuántas órdenes tiene el proveedor (para el consecutivo)
        public int ContarOrdenesPorProveedor(int idProveedor)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "SELECT COUNT(*) FROM ordenes_compra WHERE id_proveedor = @id", con);
            cmd.Parameters.AddWithValue("@id", idProveedor);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        /// Guarda el número de orden legible en la fila recién insertada
        public void GuardarNumeroOrden(int idOrden, string numeroOrden)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "UPDATE ordenes_compra SET numero_orden = @num WHERE id_orden = @id", con);
            cmd.Parameters.AddWithValue("@num", numeroOrden);
            cmd.Parameters.AddWithValue("@id",  idOrden);
            cmd.ExecuteNonQuery();
        }

        public int InsertarOrden(int idProveedor, decimal total)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "INSERT INTO ordenes_compra (id_proveedor, total) " +
                "VALUES (@prov, @total); SELECT LAST_INSERT_ID();", con);
            cmd.Parameters.AddWithValue("@prov",  idProveedor);
            cmd.Parameters.AddWithValue("@total", total);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        public void InsertarDetalle(int idOrden, int idProducto, int cantidad, decimal precio)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "INSERT INTO detalle_orden (id_orden, id_producto, cantidad, precio, subtotal) " +
                "VALUES (@orden,@prod,@cant,@precio,@sub);", con);
            cmd.Parameters.AddWithValue("@orden",  idOrden);
            cmd.Parameters.AddWithValue("@prod",   idProducto);
            cmd.Parameters.AddWithValue("@cant",   cantidad);
            cmd.Parameters.AddWithValue("@precio", precio);
            cmd.Parameters.AddWithValue("@sub",    cantidad * precio);
            cmd.ExecuteNonQuery();
        }

        public List<OrdenCompra> ObtenerOrdenes()
        {
            var lista = new List<OrdenCompra>();
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "SELECT id_orden, id_proveedor, total, fecha, estado " +
                "FROM ordenes_compra ORDER BY fecha DESC", con);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                lista.Add(new OrdenCompra
                {
                    IdOrden     = reader.GetInt32("id_orden"),
                    IdProveedor = reader.GetInt32("id_proveedor"),
                    Total       = reader.GetDecimal("total"),
                    Fecha       = reader.GetDateTime("fecha"),
                    Estado      = reader.GetString("estado")
                });
            return lista;
        }

        // ══════════════════════════════════════════════════════
        //  PRODUCTOS
        // ══════════════════════════════════════════════════════

        public List<Producto> ObtenerProductos()
        {
            var lista = new List<Producto>();
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "SELECT id, nombre, precio, id_proveedor FROM productos", con);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                lista.Add(new Producto
                {
                    Id          = reader.GetInt32("id"),
                    Nombre      = reader.GetString("nombre"),
                    Precio      = reader.GetDecimal("precio"),
                    IdProveedor = reader.GetInt32("id_proveedor")
                });
            return lista;
        }

        public string? ObtenerCorreoProveedor(int idProveedor)
        {
            using var con = CrearConexion();
            con.Open();
            var cmd = new MySqlCommand(
                "SELECT correo FROM proveedores WHERE id = @id", con);
            cmd.Parameters.AddWithValue("@id", idProveedor);
            return cmd.ExecuteScalar()?.ToString();
        }
    }
}