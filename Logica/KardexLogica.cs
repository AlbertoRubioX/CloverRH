using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using Datos;

namespace Logica
{
    public class KardexLogica
    {
        public DateTime Fecha { get; set; }
        public string Proceso { get; set; }
        public string Descrip { get; set; }
        public string Ubicacion { get; set; }
        public string Hora { get; set; }
        public static int Guardar(KardexLogica kar)
        {
            string[] parametros = { "@Proceso", "@Descrip", "@Ubicacion" };
            return AccesoDatos.Actualizar("sp_mant_kardex", parametros, kar.Proceso,kar.Descrip,kar.Ubicacion );
        }

        public static int AsisMinDif(KardexLogica kar)
        {
            string[] parametros = { "@Proceso"};
            return AccesoDatos.Actualizar("sf_asist_mindif", parametros, kar.Proceso);
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_kardex");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultaDia(KardexLogica kar)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_kardex where cast(fecha as date) = cast('"+kar.Fecha+"' as date)");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        public static DateTime ConsultaDiaGen(KardexLogica kar)
        {
            DataTable datos = new DataTable();
            DateTime dtFecha = DateTime.Now;
            try
            {
                datos = AccesoDatos.Consultar("SELECT MAX(fecha) FROM t_kardex where proceso = '"+kar.Proceso+"' AND cast(fecha as date) = cast('" + kar.Fecha + "' as date)");
                if (!string.IsNullOrEmpty(datos.Rows[0][0].ToString()))
                    dtFecha = Convert.ToDateTime(datos.Rows[0][0].ToString());
                else
                {
                    dtFecha = dtFecha.AddDays(-1);
                    datos = AccesoDatos.Consultar("SELECT MAX(fecha) FROM t_kardex where proceso = '" + kar.Proceso + "' AND cast(fecha as date) = cast('" + dtFecha + "' as date)");
                    if (!string.IsNullOrEmpty(datos.Rows[0][0].ToString()))
                        dtFecha = Convert.ToDateTime(datos.Rows[0][0].ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dtFecha;
        }

        public static bool ValidaDiaHoraGen(KardexLogica kar)
        {
            DataTable datos = new DataTable();
            int iCant = 0;
            try
            {
                datos = AccesoDatos.Consultar("SELECT COUNT(*) FROM t_kardex where proceso = '" + kar.Proceso + "' AND cast(fecha as date) = cast('" + kar.Fecha + "' as date) AND SUBSTRING(hora,1,2) = "+kar.Hora+"");
                if (datos.Rows.Count > 0)
                {
                    iCant = Convert.ToInt16(datos.Rows[0][0].ToString());
                    if (iCant > 0)
                        return false;
                }       
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return true;
        }

        public static DataTable ListarDia(KardexLogica kar)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT CONVERT(VARCHAR,fecha,105) AS FECHA,descrip AS ARCHIVO,ubicacion AS UBICACION,hora AS [HORA GENERADO] from t_kardex where CAST(fecha as date) = cast('" + kar.Fecha + "' as date)");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        public static bool Verifica(string _asCodigo)
        {
            try
            {
                string sQuery;
                if (_asCodigo == "ACT") //ACTIVOS
                    sQuery = "SELECT * FROM t_config WHERE cargar_actorbis = '1'";
                else//ASISTENCIA
                    sQuery = "SELECT * FROM t_config WHERE cargar_asisorbis = '1'";
                DataTable datos = AccesoDatos.Consultar(sQuery);
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
    }
}
