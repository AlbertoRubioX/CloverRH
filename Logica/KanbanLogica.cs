using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class KanbanLogica
    {
        public long Folio { get; set; }
        public DateTime Fecha { get; set; }
        public string Planta { get; set; }
        public string Source { get; set; }
        public string Hora { get; set; }
        public string Turno { get; set; }
        public static int Guardar(KanbanLogica kan)
        {
            string[] parametros = { "@Folio", "@Planta", "@Source", "@Hora", "@Turno" };
            return AccesoDatos.ActualizarPRO("sp_mant_kanban", parametros, kan.Folio, kan.Planta, kan.Source, kan.Hora, kan.Turno);
        }

        public static DataTable Consultar()
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("SELECT * FROM t_kanban");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }


        public static DataTable Listar(KanbanLogica kan)
        {
            DataTable datos = new DataTable();
            try
            {
                datos = AccesoDatos.ConsultarPRO("select * from t_kanban where cast(fecha as date) = cast('" + kan.Fecha + "' as date)");
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static bool Verificar(KanbanLogica kan)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_kanban where cast(fecha as date) = cast('" + kan.Fecha + "' as date) AND hora = '"+kan.Hora+"'";
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
        public static bool VerificarGlobals(KanbanLogica kan)
        {
            try
            {
                string sQuery;
                sQuery = "SELECT * FROM t_kanban where cast(fecha as date) = cast('" + kan.Fecha + "' as date) AND hora = '" + kan.Hora + "'";
                DataTable datos = AccesoDatos.ConsultarPRO(sQuery);
                if (datos.Rows.Count <= 1) // = 1
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }

        }
        public static DataTable ResumenKanbanTurno(KanbanLogica kan)
        {
            DataTable datos = new DataTable();
            try
            {
                string[] parametros = { "@Turno", "@Fecha" };
                datos = AccesoDatos.ConsultaSPPRO("sp_rep_kanban_resumen", parametros, kan.Turno, kan.Fecha);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return datos;
        }

    }
}
