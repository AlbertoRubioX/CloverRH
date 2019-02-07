using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Datos;

namespace Logica
{
    public class ModeloHcLogica
    {
        public string Planta { get; set; }
        public string Linea { get; set; }
        public string Modelo { get; set; }
        public double Standard1 { get; set; }
        public double Standard2 { get; set; }
        public double Factor { get; set; }
        public int HeadCount { get; set; }
        public static int Guardar(ModeloHcLogica mod)
        {
            string[] parametros = { "@Planta", "@Linea", "@Modelo", "@HeadCount", "@Std1er", "@Std2do", "@Factor" };
            return AccesoDatos.ActualizarPRO("sp_mant_modelohc", parametros, mod.Planta, mod.Linea, mod.Modelo, mod.HeadCount,mod.Standard1,mod.Standard2,mod.Factor);
        }
    }
}
