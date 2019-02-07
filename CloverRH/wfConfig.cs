using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Logica;

namespace CloverRH
{
    public partial class wfConfig : Form
    {
        public bool _lbValidaKan;
        public wfConfig()
        {
            InitializeComponent();
        }
        private void CargarDatos()
        {
            try
            {
                DataTable dt = ConfigLogica.Consultar();
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["ind_genact"].ToString() == "1")
                        chbActivos.Checked = true;
                    else
                        chbActivos.Checked = false;
                    txtDirecAct.Text = dt.Rows[0]["direc_act"].ToString();
                    txtFileAct.Text = dt.Rows[0]["nombre_act"].ToString();
                    txtHrAct1.Text = dt.Rows[0]["hr_1t"].ToString();
                    txtHrAct2.Text = dt.Rows[0]["hr_2t"].ToString();
                    if (dt.Rows[0]["cargar_actorbis"].ToString() == "1")
                        chbCargaAct.Checked = true;
                    else
                        chbCargaAct.Checked = false;


                    if (dt.Rows[0]["ind_genasis"].ToString() == "1")
                        chbAsistencia.Checked = true;
                    else
                        chbAsistencia.Checked = false;
                    txtDirecAsis.Text = dt.Rows[0]["direc_asis"].ToString();
                    txtFileAsis.Text = dt.Rows[0]["nombre_asis"].ToString();
                    txtHrAsis1.Text = dt.Rows[0]["hr_1tasis"].ToString();
                    txtHrAsis2.Text = dt.Rows[0]["hr_2tasis"].ToString();
                    if (dt.Rows[0]["cargar_asisorbis"].ToString() == "1")
                        chbCargaAsis.Checked = true;
                    else
                        chbCargaAsis.Checked = false;
                    txtGenMin.Text = dt.Rows[0]["asis_genmin"].ToString();

                    //CONEXION
                    txtServer.Text = dt.Rows[0]["server3"].ToString();
                    cbbTipoSer.SelectedValue = dt.Rows[0]["tipo3"].ToString();
                    txtBd.Text = dt.Rows[0]["based3"].ToString();
                    txtUser.Text = dt.Rows[0]["user3"].ToString();
                    txtClave.Text = dt.Rows[0]["passwd3"].ToString();
                    //orbis
                    txtServerOrb.Text = dt.Rows[0]["server_orb"].ToString();
                    cbbTipoOrb.SelectedValue = dt.Rows[0]["tipo_orb"].ToString();
                    txtBdOrb.Text = dt.Rows[0]["based_orb"].ToString();
                    txtUserOrb.Text = dt.Rows[0]["user_orb"].ToString();
                    txtClaveOrb.Text = dt.Rows[0]["passwd_orb"].ToString();
                    txtPuertoOrb.Text = dt.Rows[0]["puerto_orb"].ToString();
                    //kanban
                    /*
                    if (dt.Rows[0]["ind_kanban"].ToString() == "1")
                        chbKanban.Checked = true;
                    else
                        chbKanban.Checked = false;
                    txtKanDirec.Text = dt.Rows[0]["kanban_path"].ToString();
                    txtKanFile.Text = dt.Rows[0]["kanban_file"].ToString();
                    txtKanStart.Text = dt.Rows[0]["kanban_start"].ToString();
                    txtKanEnd.Text = dt.Rows[0]["kanban_end"].ToString();
                    txtKanMins.Text = dt.Rows[0]["kanban_mins"].ToString();

                    DataTable dtKp = KanbanPlanLogica.Listar();
                    dgwKanb.DataSource = dtKp;
                    ColumnasKanban();*/
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void ColumnasKanban()
        {
            int iRows = dgwKanb.Rows.Count;
            if (iRows == 0)
            {
                DataTable dtNew = new DataTable("kanplan");
                dtNew.Columns.Add("LINEA", typeof(string));//0
                dtNew.Columns.Add("NOMBRE", typeof(string));//1
                dtNew.Columns.Add("ind_1t", typeof(string));//2
                dtNew.Columns.Add("1ER TURNO", typeof(int));//3
                dtNew.Columns.Add("ind_2t", typeof(string));//4
                dtNew.Columns.Add("2DO TURNO", typeof(int));//5
                dtNew.Columns.Add("cambio", typeof(string));//6

                dgwKanb.DataSource = dtNew;
            }

            dgwKanb.Columns[2].Visible = false;
            dgwKanb.Columns[4].Visible = false;
            dgwKanb.Columns[6].Visible = false;

            dgwKanb.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgwKanb.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgwKanb.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgwKanb.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgwKanb.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgwKanb.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgwKanb.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgwKanb.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
        }
        private void wfConfig_Load(object sender, EventArgs e)
        {

            cbbTipoSer.ResetText();
            Dictionary<string, string> TipoS = new Dictionary<string, string>();
            TipoS.Add("MS", "MSSQL");
            TipoS.Add("MY", "MySQL");
            cbbTipoSer.DataSource = new BindingSource(TipoS, null);
            cbbTipoSer.DisplayMember = "Value";
            cbbTipoSer.ValueMember = "Key";
            cbbTipoSer.SelectedIndex = 0;

            cbbTipoOrb.ResetText();
            Dictionary<string, string> TipoO = new Dictionary<string, string>();
            TipoO.Add("MS", "MSSQL");
            TipoO.Add("MY", "MySQL");
            cbbTipoOrb.DataSource = new BindingSource(TipoO, null);
            cbbTipoOrb.DisplayMember = "Value";
            cbbTipoOrb.ValueMember = "Key";
            cbbTipoOrb.SelectedIndex = -1;

            CargarDatos();
            _lbValidaKan = false;
        }

        private void wfConfig_Activated(object sender, EventArgs e)
        {
            chbActivos.Focus();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {

                //KANBAN
                /*
                if (!_lbValidaKan)
                    return;

                foreach (DataGridViewRow row in dgwKanb.Rows)
                {
                    int iRows = dgwKanb.Rows.Count;
                    iRows--;
                    if (row.Index == iRows)
                        continue;

                    if (row.Cells[0].Value == null)
                        continue;

                    string sCambio = string.Empty;
                    if (row.Cells[6].Value == null)
                        sCambio = "1";
                    else
                        sCambio = row.Cells[6].Value.ToString();

                    if (sCambio == "0")
                        continue;

                    KanbanPlanLogica kp = new KanbanPlanLogica();
                    kp.Linea = row.Cells[0].Value.ToString();
                    kp.Descrip = row.Cells[1].Value.ToString();

                    kp.CantT1 = int.Parse(row.Cells[3].Value.ToString());
                    kp.CantT2 = int.Parse(row.Cells[5].Value.ToString());
                    kp.Usuario = "SYSRH";
                    KanbanPlanLogica.Guardar(kp);

                }
                */

                ConfigLogica conf = new ConfigLogica();

                if (chbActivos.Checked)
                    conf.Activos = "1";
                else
                    conf.Activos = "0";
                conf.DirecAct = txtDirecAct.Text.ToString();
                conf.FileAct = txtFileAct.Text.ToString();
                conf.HrAct1 = txtHrAct1.Text.ToString();
                conf.HrAct2 = txtHrAct2.Text.ToString();
                if (chbCargaAct.Checked)
                    conf.CargarAct = "1";
                else
                    conf.CargarAct = "0";

                if (chbAsistencia.Checked)
                    conf.Asistencia = "1";
                else
                    conf.Asistencia = "0";
                conf.DirecAsis = txtDirecAsis.Text.ToString();
                conf.FileAsis = txtFileAsis.Text.ToString();
                conf.HrAsis1 = txtHrAsis1.Text.ToString();
                conf.HrAsis2 = txtHrAsis2.Text.ToString();
                if (chbCargaAsis.Checked)
                    conf.CargarAsis = "1";
                else
                    conf.CargarAsis = "0";
                if (!string.IsNullOrEmpty(txtGenMin.Text))
                    conf.AsisGenMin = int.Parse(txtGenMin.Text.ToString());
                else
                    conf.AsisGenMin = 0;
                //CONEXION
                conf.Server = txtServer.Text.ToString();
                conf.Tipo = cbbTipoSer.SelectedValue.ToString();
                conf.Based = txtBd.Text.ToString();
                conf.User = txtUser.Text.ToString();
                conf.Passwd = txtClave.Text.ToString();
                conf.ServerOrb = txtServerOrb.Text.ToString();
                conf.TipoOrb = cbbTipoOrb.SelectedValue.ToString();
                conf.BasedOrb = txtBdOrb.Text.ToString();
                conf.UserOrb = txtUserOrb.Text.ToString();
                conf.PasswdOrb = txtClaveOrb.Text.ToString();
                conf.PuertoOrb = int.Parse(txtPuertoOrb.Text.ToString());
                //kanban
                //if (chbKanban.Checked)
                //    conf.Kanban = "1";
                //else
                //    conf.Kanban = "0";
                //conf.KanPath = txtKanDirec.Text.ToString();
                //conf.KanFile = txtKanFile.Text.ToString();
                //conf.KanStart = txtKanStart.Text.ToString();
                //conf.KanEnd = txtKanEnd.Text.ToString();
                //int iMins = 0;
                //if (int.TryParse(txtKanMins.Text.ToString(), out iMins))
                //    conf.KanMins = iMins;
                //else
                //    conf.KanMins = 0;

                if (ConfigLogica.Guardar(conf) > 0)
                    CargarDatos();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        private void btnDirAct_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtDirecAct.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnDirAsis_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtDirecAsis.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void chbCargaAsis_CheckedChanged(object sender, EventArgs e)
        {
            txtGenMin.Enabled = chbCargaAsis.Checked;
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void btnKanban_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txtKanDirec.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void chbKanban_CheckedChanged(object sender, EventArgs e)
        {
            txtKanDirec.Enabled = chbKanban.Checked;
            txtKanFile.Enabled = chbKanban.Checked;
            txtKanStart.Enabled = chbKanban.Checked;
            txtKanEnd.Enabled = chbKanban.Checked;
            txtKanMins.Enabled = chbKanban.Checked;
        }

        private void dgwKanb_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 3 || e.ColumnIndex == 5)
            {
                _lbValidaKan = true;
                dgwKanb.Rows[e.RowIndex].Cells[6].Value = "1";
            }
            
        }
    }
}
