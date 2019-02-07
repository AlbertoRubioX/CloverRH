using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class ActuserLogicaIT
    {
        public string Usuario { get; set; }
        public DateTime Fecha { get; set; }
        public long Consec { get; set; }
        public int Cant { get; set; }
        public string Sistema { get; set; }
        public string Actividad { get; set; }
        public string Solicita { get; set; }
        public long Minutos { get; set; }
        public string Catego { get; set; }
        public string Depto { get; set; }
        public int Axo { get; set; }
        public int Mes { get; set; }
        public int Semana { get; set; }
        public static int Guardar(ActuserLogicaIT act)
        {
            string[] parametros = { "@Usuario", "@Fecha", "@Consec", "@Cant", "@Sistema", "@Actividad", "@Solicita", "@Minutos", "@Catego", "@Depto", "@Axo", "@Mes", "@Semana" };
            return AccesoDatos.ActualizarIT("sp_mant_actuser", parametros, act.Usuario, act.Fecha, act.Consec, act.Cant, act.Sistema, act.Actividad, act.Solicita, act.Minutos, act.Catego, act.Depto, act.Axo, act.Mes, act.Semana);
        }

        public static bool Verificar(ActuserLogicaIT act)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_actuser where usuario = '"+act.Usuario+"' and axo = '" + act.Axo + "' and mes = '" + act.Mes + "' ";
                DataTable datos = AccesoDatos.ConsultarIT(sQuery);
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

        public static DataTable ActividadMensualCat(ActuserLogicaIT act)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Usuario", "@Axo", "@Mes", "@Catego" };
                datos = AccesoDatos.ConsultaSPIT("sp_rep_actuser_cat", parametros, act.Usuario, act.Axo, act.Mes, act.Catego);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }
        public static DataTable ActividadMensualDepto(ActuserLogicaIT act)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Usuario", "@Axo", "@Mes", "@Depto" };
                datos = AccesoDatos.ConsultaSPIT("sp_rep_actuser_depto", parametros, act.Usuario, act.Axo, act.Mes, act.Depto);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }

        public static DataTable ActividadSemanal(ActuserLogicaIT act)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Usuario", "@Axo", "@Mes", "@Semana" };
                datos = AccesoDatos.ConsultaSPIT("sp_rep_actuser_sem", parametros, act.Usuario, act.Axo, act.Mes, act.Semana);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }

        public static DataTable ActividadProyectoSem(ActuserLogicaIT act)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Usuario", "@Axo", "@Mes", "@Semana" };
                datos = AccesoDatos.ConsultaSPIT("sp_actuser_sem", parametros, act.Usuario, act.Axo, act.Mes, act.Semana);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }
        public static DataTable ActividadProyectoCat(ActuserLogicaIT act)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Usuario", "@Axo", "@Mes", "@Semana" };
                datos = AccesoDatos.ConsultaSPIT("sp_actuser_cat", parametros, act.Usuario, act.Axo, act.Mes, act.Semana);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }
        public static DataTable ActividadMesCat(ActuserLogicaIT act)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Usuario", "@Axo", "@Mes"};
                datos = AccesoDatos.ConsultaSPIT("sp_rep_actuser_mescat", parametros, act.Usuario, act.Axo, act.Mes);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }
    }
}

