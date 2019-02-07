using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class GlobalRpoLogica
    {
        public long Folio { get; set; }
        
        public string RPO { get; set; }
        public string Modelo { get; set; }
        public DateTime Print { get; set; }
        public DateTime Register { get; set; }
        public DateTime Kanban { get; set; }
        public DateTime Start { get; set; }        
        public double Finish { get; set; }
        public double Shipped { get; set; }
        public string HorKan { get; set; }
        public string Truck { get; set; }
        public string Transfer { get; set; }
        public string Pallet { get; set; }
        public DateTime Ftransfer { get; set; }
        public DateTime Fscanned { get; set; }
        public DateTime Fposted { get; set; }
        public static int Guardar(GlobalRpoLogica rpo)
        {
            string[] parametros = { "@Folio", "@RPO", "@Print", "@Register", "@Kanban", "@Start", "@Finish", "@Shipped", "@HoraKan" };
            return AccesoDatos.ActualizarPRO("sp_mant_rpo_glob", parametros, rpo.Folio, rpo.RPO, rpo.Print, rpo.Register, rpo.Kanban, rpo.Start, rpo.Finish, rpo.Shipped, rpo.HorKan);
        }
        public static int GuardarTrans(GlobalRpoLogica rpo)
        {
            string[] parametros = { "@Folio", "@Post", "@Scanned", "@Shipped", "@Transfer", "@Truck" };
            return AccesoDatos.ActualizarPRO("sp_mant_rpo_glob_trans", parametros, rpo.Folio, rpo.Fposted, rpo.Fscanned, rpo.Ftransfer ,rpo.Transfer,rpo.Truck );
        }
        public static DataTable MonitorGlobals()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultaSPPRO2("sp_mon_globals_trans");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }
        public static DataTable Consultar(GlobalRpoLogica rpo)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_rpo_glob where rpo = '"+rpo.RPO+ "' and CAST(fecha AS DATE) = CAST(GETDATE() AS DATE)");

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        
        public static bool Verificar(GlobalRpoLogica rpo)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_rpo_glob where rpo = '"+rpo.RPO+"' and CAST(fecha AS DATE) = CAST(GETDATE() AS DATE)";
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

        public static bool VerificarTrans(GlobalRpoLogica rpo)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_rpo_glob where rpo = '" + rpo.RPO + "' and cancelado='0'";
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

        public static bool VerificarNto(GlobalRpoLogica rpo)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_rpo_glob where modelo = '"+rpo.Modelo+"' and cantidad = "+rpo.Shipped+" and cancelado='0' and CAST(fecha AS DATE) = CAST(GETDATE() AS DATE)";
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
        public static DataTable ConsultarTrans(GlobalRpoLogica rpo)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT max(folio) FROM t_rpo_glob where rpo = '" + rpo.RPO + "' and cancelado='0'");

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }


        public static bool VerificarListado()
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_rpo_glob where CAST(fecha AS DATE) = CAST(GETDATE() AS DATE)";
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
    }
}
