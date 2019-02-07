using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CloverRH
{
    public partial class CapturaPop : Form
    {
        public DateTime _dtReturn;
        public CapturaPop()
        {
            InitializeComponent();
        }

        private void btnAceptar_Click(object sender, EventArgs e)
        {
            _dtReturn = dtpFecha.Value;
            Close();
        }
    }
}
