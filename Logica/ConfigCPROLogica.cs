using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class ConfigCPROLogica
    {
        public string Kanban { get; set; }
        public string KanPath { get; set; }
        public string KanFile { get; set; }
        public string KanStart { get; set; }
        public string KanEnd { get; set; }
        public int KanMins { get; set; }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_config WHERE clave = '01'");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        

    }
}
