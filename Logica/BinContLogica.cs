using Datos;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logica
{
    public class BinContLogica
    {
        public string folio { get; set; }
        public string hora { get; set; }
        public DateTime fecha { get; set; }
        public string planta { get; set; }
        public string bincode { get; set; }
        public string item { get; set; }
        public string descrip { get; set; }
        public double cantidad { get; set; }
        public double contador { get; set; }
        public double diferencia { get; set; }
        public string um { get; set; }


        public static void guardar(BinContLogica bincont)
        {
            string[] parametros = {"@Hora","@Planta","@Bincode","@Item","@Descrip","@UM","@Cantidad"};
            AccesoDatos.ActualizarPRO("sp_mant_bincont", parametros, bincont.hora, bincont.planta, bincont.bincode, bincont.item, bincont.descrip, bincont.um, bincont.cantidad);

        }

        public static bool VerificarRegistros( BinContLogica bincont)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_bincont where cast(fecha as date) = cast(GETDATE() as date) AND hora = '" + bincont.hora+ "'";
                DataTable datos = AccesoDatos.ConsultarPRO(sQuery);
                if (datos.Rows.Count != 0)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }

        }

        public static DataTable obtenerBinCont(BinContLogica bincont)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_bincont where cast(fecha as date) = cast(GETDATE() as date) AND hora = '" + bincont.hora + "'";
                DataTable datos = AccesoDatos.ConsultarPRO(sQuery);
                return datos;
            }
            catch
            {
                return null;
            }
        }

        public static void ActualizarBinCont( BinContLogica bincon,double contador,double diferencia)
        {
            string sQuery= "UPDATE t_bincont SET contador="+contador+", diferencia="+diferencia+" WHERE folio="+bincon.folio;
            DataTable datos = AccesoDatos.ConsultarPRO(sQuery);
        }


    }
}
