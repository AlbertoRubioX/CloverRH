using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Datos;

namespace Logica
{
    public class TressActivos
    {
        public string Turno { get; set; }
        public string Codigo { get; set; }
        public DateTime Fecha { get; set; }
        public static DataTable Consultar(TressActivos act)
        {
            DataTable datos = new DataTable();
            try
            {
                //"INNER JOIN NIVEL3 pta ON substring(col.CB_NIVEL3,1,3) = pta.TB_CODIGO "+
                string sSql = "SELECT pta.TB_ELEMENT AS Planta,col.CB_NIVEL2 AS Linea,col.CB_TURNO AS Turno,col.CB_CODIGO as NumEmp,col.PRETTYNAME AS NombreEmp,pt.PU_DESCRIP as Nivel,col.CB_SALARIO as [SueldoDiario]   " +
                "FROM COLABORA col " +
                "INNER JOIN NIVEL3 pta ON col.CB_NIVEL3 = pta.TB_CODIGO " +
                "INNER JOIN PUESTO pt ON col.CB_PUESTO = pt.PU_CODIGO " +
                "WHERE col.CB_NIVEL0 = '1' AND col.CB_NIVEL1 = '001' AND col.CB_ACTIVO = 'S' AND col.CB_NIVEL2 <> 'NI001' " +
                "AND pta.TB_CODIGO NOT IN('004','004-A','004-B','004-C')";
                datos = AccesoDatos.ConsultarTress(sSql);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultarAsis(TressActivos act)
        {
            DataTable datos = new DataTable();
            try
            {
                string sSql = "select co.CB_CODIGO as [No.],co.PRETTYNAME as [Nombre], " +
                "CASE (select dbo.sp_status_emp(GETDATE(), co.CB_CODIGO)) when '1' then '' when '2' then 'Vacaciones' when '3' then 'Incapacidad' when '4' then 'Permiso c/Goce' when '5' then 'Permiso s/Goce' when '6' then 'Falta Justificada' when '7' then 'Suspensión' when '8' then 'Otros' END as [Status], "+
                "CASE (select dbo.SP_CHECADAS(cast(getdate() as date), co.CB_CODIGO, 1)) when '' then 'Falta' else '' END as Tipo, "+
                "(SELECT dbo.SP_CHECADAS(cast(GETDATE() as date), co.CB_CODIGO, 1)) as [1 / Ent],'  :  ' as [2 / Sal],'  :  ' as [2 / Ent],'  :  '[2 / Sal] "+
                "FROM COLABORA co "+
                "WHERE co.CB_ACTIVO = 'S' "+
                "AND ('" + act.Turno + "' = '1' and co.CB_TURNO = '1' or ('" + act.Turno + "' = '2' )) " +
                "AND co.CB_NIVEL1 = '001' " +
                "ORDER BY co.CB_TURNO, co.CB_CODIGO";
                datos = AccesoDatos.ConsultarTress(sSql);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static DataTable ConsultarAsisHrs(TressActivos act)
        {
            DataTable datos = new DataTable();
            try
            {
                DateTime dtFecha = act.Fecha;
                if (DateTime.Now.Hour >= 0 && DateTime.Now.Hour <= 5)
                    dtFecha = dtFecha.AddDays(-1);

                //dtFecha = Convert.ToDateTime("2018-07-13");
                //"INNER JOIN NIVEL3 pta ON substring(col.CB_NIVEL3,1,3) = pta.TB_CODIGO "+

                string sSql = "SELECT co.CB_CODIGO as [No.],co.PRETTYNAME as [NOMBRE], pta.TB_ELEMENT AS PLANTA," +
                "co.CB_TURNO AS TURNO,co.CB_NIVEL2 as LINEA, " +
                "CONVERT(varchar(10),cast('" + dtFecha + "' as date), 101) as FECHA, " +
                "(SELECT dbo.SP_CHECADAS(cast('"+ dtFecha +"' as date), co.CB_CODIGO, 1)) as [HORA ENT], " +
                "(SELECT dbo.SP_CHECADAS(cast('" + dtFecha + "' as date), co.CB_CODIGO, 2)) as [HORA SAL], " +
                "(SELECT dbo.SP_CHECADAS(cast('" + dtFecha + "' as date), co.CB_CODIGO, 3)) as [HORA ENT2], " +
                "(SELECT dbo.SP_CHECADAS(cast('" + dtFecha + "' as date), co.CB_CODIGO, 4)) as [HORA SAL2], " +
                "ac.AU_HORASCK as [HRS TRABAJADAS],ac.AU_EXTRAS AS [HRS EXTRAS],ac.AU_NUM_EXT AS [SIN_AUT], " +
                "ac.AU_TIPO AS [TIPO INCIDENCIA]," +
                "CASE(select dbo.SP_CHECADAS(cast('" + dtFecha + "' as date), co.CB_CODIGO, 1)) when '' then 'Falta' else '' END as ESTATUS, " +
                "CASE(select dbo.sp_status_emp(cast('" + dtFecha + "' as date), co.CB_CODIGO)) when '1' then '' when '2' then 'Vacaciones' when '3' then 'Incapacidad' when '4' then 'Permiso c/Goce' when '5' then 'Permiso s/Goce' when '6' then 'Falta Justificada' when '7' then 'Suspensión' when '8' then 'Otros' END as [DESCRIP_INCIDENCIA] " +
                "FROM COLABORA co LEFT OUTER JOIN AUSENCIA ac ON co.CB_CODIGO = ac.CB_CODIGO AND CAST(ac.AU_FECHA AS DATE) = CAST('"+dtFecha+"' AS DATE) " +
                "INNER JOIN NIVEL3 pta ON co.CB_NIVEL3 = pta.TB_CODIGO " +
                "WHERE co.CB_ACTIVO = 'S' " +
                "AND co.CB_NIVEL1 = '001' " +
                "AND co.CB_NIVEL2 <> 'NI001' " +
                "AND co.CB_NIVEL3 NOT IN('004','004-A','004-B','004-C') " +
                "AND(SELECT dbo.SP_CHECADAS(cast('" + dtFecha + "' as date), co.CB_CODIGO, 1)) > '' " +
                "ORDER BY co.CB_TURNO, co.CB_CODIGO";
                datos = AccesoDatos.ConsultarTress(sSql);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return datos;
        }

        public static string TipoAusencia(TressActivos act)
        {
            string sValue = string.Empty;
            DataTable datos = new DataTable();
            try
            {
                DateTime dtFecha = DateTime.Today;
                //dtFecha = Convert.ToDateTime("2018-04-27");

                string sSql = "SELECT AU_TIPO FROM AUSENCIA WHERE CB_CODIGO = " + act.Codigo + " AND CAST(AU_FECHA AS DATE) = CAST('"+dtFecha+"' AS DATE);";
                datos = AccesoDatos.ConsultarTress(sSql);
                if (datos.Rows.Count > 0)
                    sValue = datos.Rows[0][0].ToString();
                else
                    sValue = "";
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return sValue;
        }
        public static string HorasAusencia(TressActivos act)
        {
            string sValue = string.Empty;
            DataTable datos = new DataTable();
            try
            {
                DateTime dtFecha = DateTime.Today;
                //dtFecha = Convert.ToDateTime("2018-04-27");

                string sSql = "SELECT AU_HORASCK FROM AUSENCIA WHERE CB_CODIGO = " + act.Codigo + " AND CAST(AU_FECHA AS DATE) = CAST('"+dtFecha+"' AS DATE)";
                datos = AccesoDatos.ConsultarTress(sSql);
                if (datos.Rows.Count > 0)
                    sValue = datos.Rows[0][0].ToString();
                else
                    sValue = "";
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return sValue;
        }
    }
}
