using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using Datos;

namespace Logica
{
    public class ConfigLogica
    {
        public string Activos { get; set; }
        public string DirecAct { get; set; }
        public string FileAct { get; set; }
        public string HrAct1 { get; set; }
        public string HrAct2 { get; set; }
        public string CargarAct { get; set; }
        public string Asistencia { get; set; }
        public string DirecAsis { get; set; }
        public string FileAsis { get; set; }
        public string HrAsis1 { get; set; }
        public string HrAsis2 { get; set; }
        public string CargarAsis { get; set; }
        public string Server { get; set; }
        public string Tipo { get; set; }
        public string Based { get; set; }
        public string User { get; set; }
        public string Passwd { get; set; }
        public string ServerOrb { get; set; }
        public string TipoOrb { get; set; }
        public string BasedOrb { get; set; }
        public string UserOrb { get; set; }
        public string PasswdOrb { get; set; }
        public int PuertoOrb { get; set; }
        public int AsisGenMin { get; set; }
        public string Kanban { get; set; }
        public string KanPath { get; set; }
        public string KanFile { get; set; }
        public string KanStart { get; set; }
        public string KanEnd { get; set; }
        public int KanMins { get; set; }

        public static int Guardar(ConfigLogica config)
        {
            string[] parametros = { "@Activos", "@DirecAct", "@FileAct", "@HrAct1", "@HrAct2", "@CargaAct", "@Asistencia", "@DirecAsis", "@FileAsis", "@HrAsis1", "@HrAsis2", "@CargaAsis", "@Server3", "@Tipo3", "@Based3", "@User3", "@Passwd3", "@ServerOrb", "@TipoOrb", "@BasedOrb", "@UserOrb", "@PasswdOrb", "@PuertoOrb", "@AsisGenMin", "@Kanban", "@KanDirec", "@KanFile", "@KanStart", "@KanEnd", "@KanMins" };
            return AccesoDatos.Actualizar("sp_mant_config", parametros, config.Activos, config.DirecAct, config.FileAct, config.HrAct1, config.HrAct2, config.CargarAct, config.Asistencia, config.DirecAsis, config.FileAsis, config.HrAsis1, config.HrAsis2, config.CargarAsis, config.Server, config.Tipo, config.Based, config.User, config.Passwd, config.ServerOrb, config.TipoOrb, config.BasedOrb, config.UserOrb, config.PasswdOrb, config.PuertoOrb,config.AsisGenMin, config.Kanban, config.KanPath, config.KanFile, config.KanStart, config.KanEnd, config.KanMins );
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.Consultar("SELECT * FROM t_config");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ListarReportes()
        {
            DataTable datos = new DataTable();
            try
            {
                string sQuery = "select nombre_act AS REPORTE,ind_genact AS ACTIVO,direc_act AS DIRECTORIO,hr_1t AS [HORA 1T],hr_2t AS [HORA 2T] from t_config union " +
                                "select nombre_asis, ind_genasis, direc_asis, hr_1tasis, hr_2tasis from t_config";
                datos = AccesoDatos.Consultar(sQuery);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        public static bool VerificaCargaAuto(string _asCodigo)
        {
            try
            {
                string sQuery;
                if(_asCodigo == "ACT") //ACTIVOS
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
