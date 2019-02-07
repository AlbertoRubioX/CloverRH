using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace Datos
{
    public class ConexionT
    {

        public static string cadenaConexion = ConfigurationManager.ConnectionStrings["Tress_Connection"].ToString();

        private static void Cadena()
        {
            if (string.IsNullOrEmpty(cadenaConexion))
                cadenaConexion = "Data Source = MXAPP6\\MXTRESS; Initial Catalog = Dataproducts; Persist Security Info = True; User ID = tress20; Password = Tress20";
        }
        public static string CadenaConexion()
        {
            Cadena();
            return cadenaConexion;
        }
    }
}
