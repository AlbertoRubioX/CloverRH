using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace Datos
{
    public class MetodosDatos
    {
        public static SqlCommand CrearComando()
        {
            string cadenaConexion = Conexion.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            _conn.ConnectionString = cadenaConexion;
            SqlCommand _comando = new SqlCommand();
            _comando = _conn.CreateCommand();
            _comando.CommandType = CommandType.Text;
            return _comando;
        }
        public static SqlCommand CrearComandoTress()
        {
            string cadenaConexion = ConexionT.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            _conn.ConnectionString = cadenaConexion;
            SqlCommand _comando = new SqlCommand();
            _comando = _conn.CreateCommand();
            _comando.CommandType = CommandType.Text;
            return _comando;
        }
        public static SqlCommand CrearComandoSP(string as_storeProc)
        {
            string cadenaConexion = Conexion.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            SqlCommand _comando = new SqlCommand(as_storeProc, _conn);
            _comando.CommandType = System.Data.CommandType.StoredProcedure;
            return _comando;
        }

        public static int EjecutaComando(SqlCommand comando)
        {
            try
            {
                comando.Connection.Open();
                return comando.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                string sval = ex.ToString();
                throw;
            }
            finally
            {
                comando.Connection.Dispose();
                comando.Connection.Close();
            }
        }

        public static DataTable EjecutaComandoSelect(SqlCommand comando)
        {
            DataTable _tabla = new DataTable();
            try
            {
                comando.Connection.Open();
                SqlDataAdapter _adaptador = new SqlDataAdapter();
                _adaptador.SelectCommand = comando;
                _adaptador.Fill(_tabla);
            }
            catch (Exception ex)
            { throw ex; }
            finally
            {
                comando.Connection.Close();
            }
            return _tabla;
        }
        #region regCPRO
        public static SqlCommand CrearComandoPRO()
        {
            string cadenaConexion = ConexionCPRO.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            _conn.ConnectionString = cadenaConexion;
            SqlCommand _comando = new SqlCommand();
            _comando = _conn.CreateCommand();
            _comando.CommandType = CommandType.Text;
            return _comando;
        }

        public static SqlCommand CrearComandoSPPRO(string as_storeProc)
        {
            string cadenaConexion = ConexionCPRO.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            SqlCommand _comando = new SqlCommand(as_storeProc, _conn);
            _comando.CommandType = CommandType.StoredProcedure;
            return _comando;
        }
        public static DataTable EjecutaComandoSelectPRO(SqlCommand comando)
        {
            DataTable _tabla = new DataTable();
            try
            {
                comando.Connection.Open();
                SqlDataAdapter _adaptador = new SqlDataAdapter();
                _adaptador.SelectCommand = comando;
                _adaptador.Fill(_tabla);
            }
            catch (Exception ex)
            { throw ex; }
            finally
            {
                comando.Connection.Close();
            }
            return _tabla;
        }
        #endregion

        #region regCIT
        public static SqlCommand CrearComandoIT()
        {
            string cadenaConexion = ConexionCIT.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            _conn.ConnectionString = cadenaConexion;
            SqlCommand _comando = new SqlCommand();
            _comando = _conn.CreateCommand();
            _comando.CommandType = CommandType.Text;
            return _comando;
        }

        public static SqlCommand CrearComandoSPIT(string as_storeProc)
        {
            string cadenaConexion = ConexionCIT.CadenaConexion();
            SqlConnection _conn = new SqlConnection(cadenaConexion);
            SqlCommand _comando = new SqlCommand(as_storeProc, _conn);
            _comando.CommandType = CommandType.StoredProcedure;
            return _comando;
        }
        public static DataTable EjecutaComandoSelectIT(SqlCommand comando)
        {
            DataTable _tabla = new DataTable();
            try
            {
                comando.Connection.Open();
                SqlDataAdapter _adaptador = new SqlDataAdapter();
                _adaptador.SelectCommand = comando;
                _adaptador.Fill(_tabla);
            }
            catch (Exception ex)
            { throw ex; }
            finally
            {
                comando.Connection.Close();
            }
            return _tabla;
        }
        #endregion
    }
}
