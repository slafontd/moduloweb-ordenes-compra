using MySql.Data.MySqlClient;
using System;

namespace ModuloWeb.BROKER
{
    public class ConexionBD
    {
        private static string cadena =
            "Server=localhost;Database=moduloweb;Uid=root;Pwd=3816Sa810&;";

        public static MySqlConnection Conectar()
        {
            return new MySqlConnection(cadena);
        }

        public static bool ProbarConexion()
        {
            try
            {
                using var con = new MySqlConnection(cadena);
                con.Open();
                Console.WriteLine(" Conexión exitosa a MySQL.");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(" Error de conexión: " + ex.Message);
                return false;
            }
        }
    }
}

