using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace Datos
{
    public class AccesoDatos
    {
        public static int Actualizar(string as_procedimiento, string[] nomparametros, params Object[] valparametros)
        {
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSP(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                return MetodosDatos.EjecutaComando(_comando);
            }
            return 0;
        }
        private static object ToDBNull(object value)
        {
            if (null != value)
            {
                Type t = value.GetType();
                if (t.Equals(typeof(DateTime)))
                {
                    if (value.Equals(DateTime.MinValue))
                        return DBNull.Value;
                    else
                        return value;
                }
                else
                    return value;
            }
            return DBNull.Value;
        }
        public static DataTable ConsultaSP(string as_procedimiento, string[] nomparametros, params object[] valparametros)
        {
            DataTable dt = new DataTable();
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSP(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                dt = MetodosDatos.EjecutaComandoSelect(_comando);
            }
            return dt;
        }
        //reciben querys
        public static DataTable Consultar(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelect(_comando);
        }

        public static DataTable ConsultarTress(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoTress();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelect(_comando);
        }

        public static int Borrar(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComando(_comando);
        }
        //agrega consecutivo
        public static int Consec(string as_proceso)
        {
            int _iFolio;
            SqlCommand _comando = MetodosDatos.CrearComandoSP("sp_consec");
            _comando.Parameters.AddWithValue("@Proceso", as_proceso);
            SqlParameter _Consec = new SqlParameter("@Folio", SqlDbType.Int);
            _Consec.Direction = ParameterDirection.Output;
            _comando.Parameters.Add(_Consec);
            MetodosDatos.EjecutaComando(_comando);
            _iFolio = Int32.Parse(_Consec.Value.ToString());
            return _iFolio;
        }
        //wfCodBarrasComFIn.cs
        public static string Turno()
        {
            string _sTurno;
            SqlCommand _comando = MetodosDatos.CrearComandoSP("sp_turno");
            SqlParameter _Tur = new SqlParameter("@Turno", SqlDbType.VarChar);
            _Tur.Direction = ParameterDirection.Output;
            _comando.Parameters.Add(_Tur);
            MetodosDatos.EjecutaComando(_comando);
            _sTurno = _Tur.Value.ToString();
            return _sTurno;
        }

        public static DataTable VerificarUsuario(string as_usuario)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = "SELECT * FROM t_usuario WHERE usuario = '" + as_usuario + "'";
            return MetodosDatos.EjecutaComandoSelect(_comando);
        }
        //sust. por ConsultaUsuario
        public static DataTable TraerUsuario(string as_usuario)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = "SELECT us.usuario,us.nombre,us.planta,pl.nombre,us.modulo,us.area,us.turno FROM t_usuario us INNER JOIN t_plantas pl on us.planta = pl.planta WHERE us.usuario = '" + as_usuario + "' ";
            return MetodosDatos.EjecutaComandoSelect(_comando);
        }
        public static DataTable TraerPermisos(string as_usuario)
        {
            SqlCommand _comando = MetodosDatos.CrearComando();
            _comando.CommandText = "select distinct proceso from t_usuaper where usuario = '" + as_usuario + "' and permiso = '1' ";
            return MetodosDatos.EjecutaComandoSelect(_comando);
        }

        #region regCloverPRo
        //actualizar cpro
        public static int ActualizarPRO(string as_procedimiento, string[] nomparametros, params Object[] valparametros)
        {
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                return MetodosDatos.EjecutaComando(_comando);
            }
            return 0;
        }
        public static DataTable ConsultaSPPRO(string as_procedimiento, string[] nomparametros, params object[] valparametros)
        {
            DataTable dt = new DataTable();
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                dt = MetodosDatos.EjecutaComandoSelect(_comando);
            }
            return dt;
        }
        public static DataTable ConsultaSPPRO2(string as_procedimiento)
        {
            DataTable dt = new DataTable();
            
            SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
        
            dt = MetodosDatos.EjecutaComandoSelect(_comando);
            
            return dt;
        }
        //reciben querys cloverpro
        public static DataTable ConsultarPRO(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoPRO();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelectPRO(_comando);
        }

        public static int EjecutaSPCPRO(string as_procedimiento)
        {
            DataTable dt = new DataTable();
            SqlCommand _comando = MetodosDatos.CrearComandoSPPRO(as_procedimiento);
            return MetodosDatos.EjecutaComando(_comando);
        }

        public static int UpdatePRO(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoPRO();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComando(_comando);
        }

        public static int ConsecPRO(string as_proceso)
        {
            int _iFolio;
            SqlCommand _comando = MetodosDatos.CrearComandoSPPRO("sp_consec");
            _comando.Parameters.AddWithValue("@Proceso", as_proceso);
            SqlParameter _Consec = new SqlParameter("@Folio", SqlDbType.Int);
            _Consec.Direction = ParameterDirection.Output;
            _comando.Parameters.Add(_Consec);
            MetodosDatos.EjecutaComando(_comando);
            _iFolio = Int32.Parse(_Consec.Value.ToString());
            return _iFolio;
        }
        #endregion

        #region regCloverIT
        //actualizar cpro
        public static int ActualizarIT(string as_procedimiento, string[] nomparametros, params Object[] valparametros)
        {
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPIT(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                return MetodosDatos.EjecutaComando(_comando);
            }
            return 0;
        }
        public static DataTable ConsultaSPIT(string as_procedimiento, string[] nomparametros, params object[] valparametros)
        {
            DataTable dt = new DataTable();
            if (nomparametros.Length == valparametros.Length)
            {
                SqlCommand _comando = MetodosDatos.CrearComandoSPIT(as_procedimiento);
                int i = 0;
                foreach (string nomparam in nomparametros)
                    _comando.Parameters.AddWithValue(nomparam, ToDBNull(valparametros[i++]));

                dt = MetodosDatos.EjecutaComandoSelectIT(_comando);
            }
            return dt;
        }
        public static DataTable ConsultaSPIT2(string as_procedimiento)
        {
            DataTable dt = new DataTable();

            SqlCommand _comando = MetodosDatos.CrearComandoSPIT(as_procedimiento);

            dt = MetodosDatos.EjecutaComandoSelectIT(_comando);

            return dt;
        }
        //reciben querys cloverpro
        public static DataTable ConsultarIT(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoIT();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComandoSelectIT(_comando);
        }

        public static int EjecutaSPCIT(string as_procedimiento)
        {
            DataTable dt = new DataTable();
            SqlCommand _comando = MetodosDatos.CrearComandoSPIT(as_procedimiento);
            return MetodosDatos.EjecutaComando(_comando);
        }

        public static int UpdateIT(string as_query)
        {
            SqlCommand _comando = MetodosDatos.CrearComandoIT();
            _comando.CommandText = as_query;
            return MetodosDatos.EjecutaComando(_comando);
        }

        public static int ConsecIT(string as_proceso)
        {
            int _iFolio;
            SqlCommand _comando = MetodosDatos.CrearComandoSPIT("sp_consec");
            _comando.Parameters.AddWithValue("@Proceso", as_proceso);
            SqlParameter _Consec = new SqlParameter("@Folio", SqlDbType.Int);
            _Consec.Direction = ParameterDirection.Output;
            _comando.Parameters.Add(_Consec);
            MetodosDatos.EjecutaComando(_comando);
            _iFolio = Int32.Parse(_Consec.Value.ToString());
            return _iFolio;
        }
        #endregion
    }
}
