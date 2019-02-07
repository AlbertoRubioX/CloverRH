using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class KanbanPlanLogica
    {
        public string Linea { get; set; }
        public string Descrip { get; set; }
        public string Turno1 { get; set; }
        public int CantT1 { get; set; }
        public string Turno2 { get; set; }
        public int CantT2 { get; set; }
        public double Horas { get; set; }
        public string Usuario { get; set; }


        public static int Guardar(KanbanPlanLogica kan)
        {
            string[] parametros = { "@Linea", "@Descrip", "@Cant1t", "@Cant2t", "@Usuario" };
            return AccesoDatos.ActualizarPRO("sp_mant_kanban_plan", parametros, kan.Linea, kan.Descrip, kan.CantT1, kan.CantT2, kan.Usuario);
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_kanban_plan order by line");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }


        public static DataTable Listar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("select linea as LINEA,descrip as NOMBRE,ind_1t,cant_1t as [1ER TURNO],ind_2t,cant_2t as [2DO TURNO],'0' from t_kanban_plan order by linea");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool Verificar(KanbanPlanLogica kan)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_kanban_plan where linea = '"+kan.Linea+"'";
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
        public static bool Eliminar(KanbanPlanLogica kan)
        {
            try
            {
                string sQuery = "DELETE FROM t_kanban_plan WHERE linea = '" + kan.Linea + "'";
                if (AccesoDatos.Borrar(sQuery) != 0)
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
