using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Logica;
using Datos;
using CsvHelper;

namespace CloverRH
{
    public partial class wfAsistente : Form
    {
        private int _iAxo;
        private int _iMes;
        private string _sUsuario;
        private string _lsPath;
        public wfAsistente()
        {
            InitializeComponent();
        }

        #region regInicio
        private void wfAsistente_Load(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = ConfigCPROLogica.Consultar();
                _lsPath = dt.Rows[0]["bin_directory"].ToString();

                _sUsuario = Environment.UserName.ToString().ToUpper();
                if (_sUsuario == "AGONZ0")
                    btnMensual.Visible = true;

                sttVersion.Text = "1.0.0.43";

                CargarColumnas(0);
                CargarColumnas(1);

                CargarData();
                CargarDetalle();

                timer1.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Favor de Notificar al Administrador" + Environment.NewLine + ex.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Iniciar();
        }
        private void Iniciar()
        {
            /*
            //TRESS
            GeneraAsistencia(DateTime.Today, true);
            //CPRO
            GeneraKanban(true);
            GeneraGlobals();
            GeneraEnvios();
            */

            GeneraInvCiclico();
            GenerarInvCiclicoPickline();
        }
        #endregion

        #region regGrids
        private void CargarData()
        {
            try
            {
                DataTable data = ConfigLogica.ListarReportes();
                dgwData.DataSource = data;
            }
            catch
            {
                throw;
            }
        }
        private void CargarDetalle()
        {
            try
            {
                KardexLogica kar = new KardexLogica();
                kar.Fecha = DateTime.Today;

                DataTable det = KardexLogica.ListarDia(kar);
                dgwDetalle.DataSource = det;
                if (dgwDetalle.Rows.Count > 0)
                    dgwDetalle.Rows[0].Cells[0].Selected = false;
            }
            catch
            {
                throw;
            }
        }
        private int ColumnWith(DataGridView _dtGrid, double _dColWith)
        {

            double dW = _dtGrid.Width - 10;
            double dTam = _dColWith;
            double dPor = dTam / 100;
            dTam = dW * dPor;
            dTam = Math.Truncate(dTam);

            return Convert.ToInt32(dTam);
        }
        private void CargarColumnas(int _aiVal)
        {
            if (_aiVal == 0)
            {
                int iRows = dgwData.Rows.Count;
                if (iRows == 0)
                {
                    DataTable dtNew = new DataTable("config");
                    dtNew.Columns.Add("REPORTE", typeof(string));//0
                    dtNew.Columns.Add("ACTIVO", typeof(string));//1
                    dtNew.Columns.Add("DIRECTORIO", typeof(string));//4
                    dtNew.Columns.Add("HORA 1T", typeof(string));//2
                    dtNew.Columns.Add("HORA 2T", typeof(string));//3

                    dgwData.DataSource = dtNew;
                }

                dgwData.Columns[0].Width = ColumnWith(dgwData, 20);//FILE
                dgwData.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwData.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwData.Columns[1].Width = ColumnWith(dgwData, 10);//ACT
                dgwData.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwData.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwData.Columns[2].Width = ColumnWith(dgwData, 50);//DIR
                dgwData.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwData.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwData.Columns[3].Width = ColumnWith(dgwData, 10);//1T
                dgwData.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwData.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwData.Columns[4].Width = ColumnWith(dgwData, 10);//2T
                dgwData.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwData.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }

            if (_aiVal == 1)
            {
                int iRows = dgwDetalle.Rows.Count;
                if (iRows == 0)
                {
                    DataTable dtNew = new DataTable("config");
                    dtNew.Columns.Add("FECHA", typeof(string));//0
                    dtNew.Columns.Add("ARCHIVO", typeof(string));//1
                    dtNew.Columns.Add("UBICACION", typeof(string));//2
                    dtNew.Columns.Add("HORA GENERADO", typeof(string));//3

                    dgwDetalle.DataSource = dtNew;
                }

                dgwDetalle.Columns[0].Width = ColumnWith(dgwDetalle, 10);//DATE
                dgwDetalle.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwDetalle.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwDetalle.Columns[1].Width = ColumnWith(dgwDetalle, 20);//FILE
                dgwDetalle.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwDetalle.Columns[2].Width = ColumnWith(dgwDetalle, 60);//DIR
                dgwDetalle.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgwDetalle.Columns[3].Width = ColumnWith(dgwDetalle, 10);//TIME
                dgwDetalle.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgwDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            }
        }

        private void dgwDetalle_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgwDetalle.Rows.Count == 0)
                return;

            if (e.RowIndex == -1)
                return;

            if (e.ColumnIndex == 2)
            {
                string sFile = dgwDetalle[e.ColumnIndex, e.RowIndex].Value.ToString();
                System.Diagnostics.Process.Start(sFile);
            }
        }
        private void dgwDetalle_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int iRow = e.RowIndex;
            if ((iRow % 2) == 0)
                e.CellStyle.BackColor = Color.LightGreen;
            else
                e.CellStyle.BackColor = Color.White;

            if (e.ColumnIndex == 2)
            {
                e.CellStyle.SelectionForeColor = Color.Blue;
                DataGridViewCellStyle sty = new DataGridViewCellStyle();
                sty.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Underline);
                sty.ForeColor = Color.Blue;

                e.CellStyle.ApplyStyle(sty);
            }

        }

        private void dgwData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgwData.Rows.Count == 0)
                return;

            if (e.RowIndex == -1)
                return;

            if (e.ColumnIndex == 2)
            {
                string sDir = dgwData[e.ColumnIndex, e.RowIndex].Value.ToString();
                System.Diagnostics.Process.Start(sDir);
            }
        }

        private void dgwData_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                e.CellStyle.SelectionForeColor = Color.Blue;
                DataGridViewCellStyle sty = new DataGridViewCellStyle();
                sty.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Underline);
                sty.ForeColor = Color.Blue;

                e.CellStyle.ApplyStyle(sty);
            }
        }

        #endregion

        #region regCsvFile
        private string TipoIncidencia(string _asCodigo)
        {
            string sTipo = string.Empty;
            

            TressActivos act = new TressActivos();
            act.Codigo = _asCodigo;
            sTipo = TressActivos.TipoAusencia(act);


            return sTipo;

        }

        private string HorasAusencia(string _asCodigo)
        {
            string sHoras = string.Empty;


            TressActivos act = new TressActivos();
            act.Codigo = _asCodigo;
            sHoras = TressActivos.HorasAusencia(act);


            return sHoras;

        }
        private void ExportarTexto(DateTime _dtFecha, string _asTipo, string _asDire, string _asFile, DataTable _aDt)
        {
            int iErrRow = 0;
            try
            {
                Cursor = Cursors.WaitCursor;

                string sDia = GetFecha(_dtFecha,1);
                string sMes = GetFecha(_dtFecha, 2);
                string sMesDesc = GetFecha(_dtFecha, 3);
                string sTurno = TurnoGlobal();
                string sLocalFile = string.Empty;
               
                bool bExists = Directory.Exists(_asDire);
                if (!bExists)
                    Directory.CreateDirectory(_asDire);

                //directorio para guardar copias
                sLocalFile = @"\\mxapp7\Interfaces\Orbis\" + sMesDesc;
                bExists = Directory.Exists(sLocalFile);
                if (!bExists)
                    Directory.CreateDirectory(sLocalFile);

                _asFile += " " + sDia.PadLeft(2, '0') + "-" + sMes.PadLeft(2, '0');

                sLocalFile += "\\" + _asFile + ".csv";//archivo para la copia
                string sFile = _asDire + "\\" + _asFile + ".csv";

                //using (var stream = new StreamWriter(sFile, false, Encoding.UTF8))
             
                
                using (var stream = new StreamWriter(sLocalFile, false, Encoding.UTF8))
                {
                    #region regActivos
                    if (_asTipo == "ACT")
                    {
                        string sRow = string.Format("{0},{1},{2},{3},{4},{5},{6}", "Planta", "Linea", "Turno", "Num_Emp", "Nombre_Empleado", "Nivel", "Sueldo_Diario");
                        stream.WriteLine(sRow);
                        for (int x = 0; x < _aDt.Rows.Count; x++)
                        {
                            string sNombre = _aDt.Rows[x][4].ToString().TrimEnd();
                            string sNivel = _aDt.Rows[x][5].ToString().TrimEnd();
                            sNivel = sNivel.Replace("PPH", "");
                            sNivel = sNivel.Replace("NIVEL", "");
                            sNivel = sNivel.Replace("NIV", "");
                            sNivel = sNivel.Replace("Q", "");
                            sNivel = sNivel.TrimStart();
                            int iPos = sNivel.IndexOf(" ");
                            if (iPos != -1)
                            {
                                sNivel = sNivel.Substring(0, iPos);
                            }
                            sNivel = sNivel.TrimEnd();

                            if (sNivel.IndexOf("PPH") != -1)
                                sNivel = sNivel.Substring(10);
                            if (sNivel.IndexOf("GUIA") != -1)
                                sNivel = "GUIA";
                            if (sNivel.IndexOf("NI") != -1) // NUEVO INGRESO - TRESS
                                sNivel = "II";
                            if (sNivel.IndexOf("III Q") != -1)
                                sNivel = "III";
                            if (sNivel.IndexOf("IV Q") != -1)
                                sNivel = "IV";
                            if (sNivel.IndexOf("1") != -1)
                                sNivel = "I";
                            if (sNivel.IndexOf("2") != -1)
                                sNivel = "II";
                            if (sNivel.IndexOf("3") != -1)
                                sNivel = "III";
                            if (sNivel.IndexOf("4") != -1)
                                sNivel = "IV";
                            if (sNivel.IndexOf("MATERIAL") != -1)
                                sNivel = "MAT";
                            if (sNivel.IndexOf("OPERA") != -1)
                                sNivel = "OPG";

                            sRow = string.Format("{0},{1},{2},{3},{4},{5},{6}",
                                _aDt.Rows[x][0].ToString().TrimEnd(), _aDt.Rows[x][1].ToString().TrimEnd(),
                                _aDt.Rows[x][2].ToString().TrimEnd(), _aDt.Rows[x][3].ToString().TrimEnd(), "\"" + sNombre + "\"",
                                sNivel, _aDt.Rows[x][6].ToString().TrimEnd());
                            stream.WriteLine(sRow);
                        }
                    }
                    #endregion


                    if (_asTipo == "ASIS")
                    {
                        string sRow = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}",
                            "Num_Emp", "Nombre_Empleado", "Planta", "Turno", "Linea", "Fecha_DD_MM_AAAA", "Hora_Entrada1", "Hora_Salida1", "Hora_Entrada2", "Hora_Salida2", "Horas_Trabajadas", "Horas_Extras", "Tipo_Incidencia", "Estatus");
                        stream.WriteLine(sRow);

                        for (int x = 0; x < _aDt.Rows.Count; x++)
                        {
                            iErrRow = x;
                            string sCodigo = _aDt.Rows[x][0].ToString().TrimEnd();
                            string sNombre = _aDt.Rows[x][1].ToString().TrimEnd();
                            string sPta = _aDt.Rows[x][2].ToString().TrimEnd();
                            string sTurn = _aDt.Rows[x][3].ToString().TrimEnd();
                            string sLinea = _aDt.Rows[x][4].ToString().TrimEnd();

                            DateTime dtFecha = DateTime.Today;
                            if (!DateTime.TryParse(_aDt.Rows[x][5].ToString(), out dtFecha))
                                dtFecha = DateTime.Today;
                            string sFecha = string.Format("{0:dd/MM/yyyy}", dtFecha);

                            if (!string.IsNullOrEmpty(sNombre))
                                sNombre = sNombre.Replace(", ", " ");
                            string sEnt = _aDt.Rows[x][6].ToString().TrimEnd();
                            string sSal = _aDt.Rows[x][7].ToString().TrimEnd();
                            string sEnt2 = _aDt.Rows[x][8].ToString().TrimEnd();
                            string sSal2 = _aDt.Rows[x][9].ToString().TrimEnd();
                            string sHoras = _aDt.Rows[x][10].ToString().TrimEnd();
                            string sHorasExt = _aDt.Rows[x][11].ToString().TrimEnd();
                            string sHorSinAut = _aDt.Rows[x][12].ToString().TrimEnd();
                            string sTipo = _aDt.Rows[x][13].ToString().TrimEnd();
                            string sEstatus = _aDt.Rows[x][14].ToString().TrimEnd();
                            
                            //double dHoras = 0; // CARGAR HRS XTRAS SIN AUTORIZAR EN HRS XTRAS
                            //if(double.TryParse(sHoras, out dHoras))
                            //{
                            //    if(dHoras == 0)
                            //    {
                            //        if (double.TryParse(sHorasExt, out dHoras))
                            //        {
                            //            if(dHoras == 0)
                            //            {
                            //                if (double.TryParse(sHorSinAut, out dHoras))
                            //                {
                            //                    if (dHoras > 0)
                            //                        sHorasExt = sHorSinAut;
                            //                }

                            //            }
                            //        }
                            //    }
                            //}

                            //if (!string.IsNullOrEmpty(sSal) && !string.IsNullOrWhiteSpace(sSal))
                            //    sHoras = HorasAusencia(sCodigo);

                            //if (!string.IsNullOrEmpty(sTipo))
                            //    sTipo = TipoIncidencia(sCodigo);

                            if (!string.IsNullOrEmpty(sEnt))
                                sEnt = sEnt.Substring(0, 2) + ":" + sEnt.Substring(2);

                            if (!string.IsNullOrEmpty(sSal))
                            {
                                sSal = sSal.Substring(0, 2) + ":" + sSal.Substring(2);
                                sSal = FormatoHora24(sSal);
                            }

                            if (!string.IsNullOrEmpty(sEnt2))
                            {
                                sEnt2 = sEnt2.Substring(0, 2) + ":" + sEnt2.Substring(2);
                                sEnt2 = FormatoHora24(sEnt2);
                            }

                            if (!string.IsNullOrEmpty(sSal2))
                            {
                                sSal2 = sSal2.Substring(0, 2) + ":" + sSal2.Substring(2);
                                sSal2 = FormatoHora24(sSal2);
                            }
                                
                            
                            sSal = @"" + sSal + @"";
                            sRow = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}",
                                sCodigo, sNombre, sPta, sTurn, sLinea, sFecha, sEnt, sSal, sEnt2, sSal2, sHoras, sHorasExt, sTipo, sEstatus, Environment.NewLine);
                            stream.WriteLine(sRow);
                            //sCodigo, "\'" + sNombre + "\'", sPta, sTurn, sLinea, sFecha, sEnt, sSal, sEnt2, sSal2, sHoras, sHorasExt, sTipo, sEstatus);
                        }
                    }

                    stream.Close();

                    if (File.Exists(sLocalFile))
                    {

                        KardexLogica kar = new KardexLogica();
                        kar.Proceso = _asTipo;
                        kar.Descrip = _asFile;
                        kar.Ubicacion = sFile;

                        if (KardexLogica.Guardar(kar) > 0)
                            CargarDetalle();

                        if (File.Exists(sFile))
                            File.Delete(sFile);

                        File.Copy(sLocalFile, sFile);
                    }
                }

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                string sEx = ex.ToString();
                MessageBox.Show(sEx + Environment.NewLine + iErrRow.ToString(), "ExportarTexto(ASISTENCIA)");
                //throw;
            }

        }
        private string FormatoHora24(string _asHor)
        {
            string sReturn;
            string sHora = _asHor.Substring(0, 2);
            string sMin = _asHor.Substring(3, 2);
            int iHora = int.Parse(sHora);
            if (iHora > 23)
                iHora -= 24;
            sHora = iHora.ToString().PadLeft(2, '0');
            sReturn = sHora + ":" + sMin;
            
            return sReturn;
        }
        private void ExportarFormato(string _asDire, string _asFile, DataTable _aDt)
        {
            Cursor = Cursors.WaitCursor;

            Microsoft.Office.Interop.Excel.Application oXL = null;
            Microsoft.Office.Interop.Excel._Workbook oWB = null;
            Microsoft.Office.Interop.Excel._Worksheet oSheet = null;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                //string sFile = @"\\mxni-fs-01\Temp\wrivera\agonz0\CloverRH\Tress\ACTIVOS_ORBIS.xlsx";

                string sFile = @"\\mxni-fs-01\Temp\wrivera\agonz0\CloverRH\Tress\" + _asFile + ".xlsx";

                //oWB = oXL.Workbooks.Open(sFile);

                //DAR FORMATO CON DIA AL ARCHIVO _asFile */*/*/*/*/
                string sDia = Convert.ToString(DateTime.Today.Day);
                string sMes = Convert.ToString(DateTime.Today.Month);
                _asFile += " " + sDia.PadLeft(2, '0') + "-" + sMes.PadLeft(2, '0');
                oWB = oXL.Workbooks.Add(sFile);

                //oSheet = String.IsNullOrEmpty("Sheet1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Sheet1"];
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;


                int iRow = 1;
                for (int x = 0; x < _aDt.Rows.Count; x++)
                {

                    oSheet.Cells[iRow, 1] = _aDt.Rows[x][0].ToString();
                    oSheet.Cells[iRow, 2] = _aDt.Rows[x][1].ToString();
                    oSheet.Cells[iRow, 3] = _aDt.Rows[x][2].ToString();
                    oSheet.Cells[iRow, 4] = _aDt.Rows[x][3].ToString();
                    oSheet.Cells[iRow, 5] = _aDt.Rows[x][4].ToString();

                    iRow++;
                }

                oXL.DisplayAlerts = false;

                sFile = _asDire + "\\" + _asFile + ".csv";
                //sFile = @"C:\CloverPRO\Formatos\ACTIVOS ORBIS.csv";
                oWB.SaveAs(sFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared);

                DialogResult Result = MessageBox.Show("Se ha exportado la consulta." + Environment.NewLine + "Desea abrir el reporte en excel?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (Result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(sFile);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                Cursor = Cursors.Default;
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close();
                    oXL.Quit();
                    Cursor = Cursors.Default;
                }
            }
            Cursor = Cursors.Default;
        }

        #endregion

        #region regReadExcel
        private DataTable getFromExcelTrans(string _asArchivo)
        {

            DataTable dt = new DataTable("TRANSFER");
            dt.Columns.Add("Scanned", typeof(string));
            dt.Columns.Add("Shipment", typeof(string));
            dt.Columns.Add("Truck", typeof(string));
            dt.Columns.Add("Postdate", typeof(string));
            dt.Columns.Add("Transfer", typeof(string));
            dt.Columns.Add("Item", typeof(string));
            dt.Columns.Add("Pallet", typeof(string));
            dt.Columns.Add("RPO", typeof(string));
            dt.Columns.Add("qty_post", typeof(double));
            dt.Columns.Add("qty_ship", typeof(double));
            int iExCont = 0;
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();

                string sValue = string.Empty;

                int iSheets = xlWorkbook.Sheets.Count;

                xlWorksheet = xlWorkbook.Sheets[iSheets];

                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    iExCont = i;

                    string sTransfer = string.Empty;
                    string sTruck = string.Empty;
                    string sRPO = string.Empty;
                    string sItem = string.Empty;
                    string sPallet = string.Empty;
                    string sFecha = string.Empty;
                    string sFechaShip = string.Empty;
                    string sFechaPost = string.Empty;
                    string sCant = string.Empty;
                    string sCant2 = string.Empty;

                    sValue = string.Empty;
                   
                    if (xlRange.Cells[i, 7].Value2 == null)//--transfer order
                        continue;

                    if (xlRange.Cells[i, 2].Value2 != null)
                        sValue = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());

                    if (sValue == "Scanned DateTime")
                    {
                        sValue = string.Empty;
                        continue;
                    }

                    if (string.IsNullOrEmpty(sValue))
                    {
                        i = rowCount;
                        continue;
                    }

                    if (xlRange.Cells[i, 3].Value2 != null)
                        sTruck = Convert.ToString(xlRange.Cells[i, 3].Value2.ToString());
                    sTransfer = Convert.ToString(xlRange.Cells[i, 6].Value2.ToString());
                    sItem = Convert.ToString(xlRange.Cells[i, 7].Value2.ToString());
                    sPallet = Convert.ToString(xlRange.Cells[i, 8].Value2.ToString());
                    sRPO = Convert.ToString(xlRange.Cells[i, 9].Value2.ToString());
                    sRPO = sRPO.TrimStart().TrimEnd().ToUpper();
                    if (xlRange.Cells[i, 1].Value2 != null)//scanned date
                        sFecha = Convert.ToString(xlRange.Cells[i, 1].Value.ToString());
                    else
                        sFecha = Convert.ToString(DateTime.Today);

                    sCant = Convert.ToString(xlRange.Cells[i, 10].Value2.ToString());
                    if (xlRange.Cells[i, 11].Value2 != null)//POST 10
                        sCant2 = Convert.ToString(xlRange.Cells[i, 11].Value.ToString());
                    else
                        sCant2 = "0";

                    double dCant = 0;
                    if(!double.TryParse(sCant,out dCant))
                        dCant = 0;
                    double dCantF = 0;
                    if (!double.TryParse(sCant2, out dCantF))
                        dCantF = 0;

                    
                    if (xlRange.Cells[i, 3].Value2 != null)//shipment date
                        sFechaShip = Convert.ToString(xlRange.Cells[i, 2].Value.ToString());

                    if (xlRange.Cells[i, 6].Value2 != null)//post20 date
                        sFechaPost = Convert.ToString(xlRange.Cells[i, 5].Value.ToString());
                    
                    dt.Rows.Add(sFecha,sFechaShip,sTruck,sFechaPost,sTransfer,sItem,sPallet,sRPO,dCant,dCantF);

                }

                xlApp.DisplayAlerts = false;
                xlWorkbook.Close();
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ex.ToString();
                MessageBox.Show(ex.ToString() + Environment.NewLine + iExCont.ToString(), "getFromExcelEmp(string _asArchivo)");

            }

            return dt;
        }
        private DataTable getFromExcelEmp(string _asArchivo)
        {

            DataTable dt = new DataTable("KANBAN");
            dt.Columns.Add("line", typeof(string));
            dt.Columns.Add("rpo", typeof(string));
            dt.Columns.Add("creation", typeof(string));
            dt.Columns.Add("item", typeof(string));
            dt.Columns.Add("qty", typeof(double));
            dt.Columns.Add("pick_print", typeof(string));
            dt.Columns.Add("pick_register", typeof(string));
            dt.Columns.Add("kanban", typeof(string));
            dt.Columns.Add("pack_start", typeof(string));
            dt.Columns.Add("qty_finished", typeof(double));
            dt.Columns.Add("qty_shipped", typeof(double));
            dt.Columns.Add("saldo", typeof(double));
            int iExCont = 0;
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();

                string sValue = string.Empty;

                int iSheets = xlWorkbook.Sheets.Count;

                xlWorksheet = xlWorkbook.Sheets[iSheets];

                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 2; i < rowCount; i++)
                {
                    iExCont = i;


                    string sLine = string.Empty;
                    string sRPO = string.Empty;
                    string sItem = string.Empty;
                    string sFecha = string.Empty;
                    string sFecha2 = string.Empty;
                    string sFecha3 = string.Empty;
                    string sFecha4 = string.Empty;
                    string sFecha5 = string.Empty;
                    string sCant = string.Empty;
                    string sCant2 = string.Empty;
                    string sCant3 = string.Empty;

                    sValue = string.Empty;
                    /*
                    if (xlRange.Cells[i, 9].Value2 == null)//--SOLO KANBAN PICK
                        continue;*/

                    //if (xlRange.Cells[i, 7].Value2 == null)//--INCLUYE GOLBALS (PICK PRINT)
                    //    continue;

                    if (xlRange.Cells[i, 1].Value2 != null)
                        sValue = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());


                    if (sValue == "Packing Line")
                    {
                        sValue = string.Empty;
                        continue;
                    }
                    if (sValue.IndexOf("MX1PAC") != -1)
                    {
                        sValue = sValue.Replace("MX1PAC", "MX1APAC");
                    }

                    if (sValue.IndexOf("MX1APAC") == -1)
                    {
                        sValue = string.Empty;
                        continue;
                    }

                    if (string.IsNullOrEmpty(sValue))
                        continue;

                    sLine = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());
                    sRPO = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());
                    sRPO = sRPO.TrimStart().TrimEnd().ToUpper();                    
                    if (xlRange.Cells[i, 3].Value2 != null)
                        sFecha = Convert.ToString(xlRange.Cells[i, 3].Value.ToString());
                    else
                        sFecha = Convert.ToString(DateTime.Today);

                    sItem = Convert.ToString(xlRange.Cells[i, 4].Value2.ToString());
                    sCant = Convert.ToString(xlRange.Cells[i, 5].Value2.ToString());
                    if (xlRange.Cells[i, 11].Value2 != null)//POST 10
                        sCant2 = Convert.ToString(xlRange.Cells[i, 11].Value.ToString());
                    else
                        sCant2 = "0";
                    double dCant = double.Parse(sCant);
                    double dCantF = 0;
                    
                    if (!double.TryParse(sCant2, out dCantF))
                        dCantF = 0;

                    /*if (dCant <= dCantF)
                        continue; v. 1.0.0.9*/

                    double dCantDif = dCant - dCantF;
                    //if (dCantDif < 0)
                    //    dCantDif = 0; - Cycle counting

                    if (xlRange.Cells[i, 7].Value2 != null)//print
                        sFecha2 = Convert.ToString(xlRange.Cells[i, 7].Value.ToString());
                    
                    if (xlRange.Cells[i, 8].Value2 != null)//register
                        sFecha3 = Convert.ToString(xlRange.Cells[i, 8].Value.ToString());
                    
                    if (xlRange.Cells[i, 9].Value2 != null)//kanban
                        sFecha4 = Convert.ToString(xlRange.Cells[i, 9].Value.ToString());
                    
                    if (xlRange.Cells[i, 10].Value2 != null)//start
                        sFecha5 = Convert.ToString(xlRange.Cells[i, 10].Value.ToString());

                    if (xlRange.Cells[i, 12].Value2 != null)//POST 20
                        sCant2 = Convert.ToString(xlRange.Cells[i, 12].Value.ToString());
                    else
                        sCant2 = "0";
                    double dCantShip = 0;
                    if (!double.TryParse(sCant2, out dCantShip))
                        dCantShip = 0;

                    dt.Rows.Add(sLine, sRPO, sFecha, sItem, dCant, sFecha2, sFecha3, sFecha4, sFecha5, dCantF, dCantShip, dCantDif);

                }

                xlApp.DisplayAlerts = false;
                xlWorkbook.Close();
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ex.ToString();
                MessageBox.Show(ex.ToString() + Environment.NewLine + iExCont.ToString(), "getFromExcelEmp(string _asArchivo)");

            }

            return dt;
        }
        #endregion

        #region GenReportes
        private string GetFecha(DateTime _dtFecha,int _aiTipo)
        {
            string sValor = string.Empty;
            //DateTime dtFecha = DateTime.Today;
            if (DateTime.Now.Hour >= 0 && DateTime.Now.Hour <= 5)
                _dtFecha = _dtFecha.AddDays(-1);

            if (_aiTipo == 1)
                sValor = Convert.ToString(_dtFecha.Day);

            if (_aiTipo == 2)
                sValor = Convert.ToString(_dtFecha.Month);

            if (_aiTipo == 3)
            {
                sValor = Convert.ToString(_dtFecha.Month);
                if (sValor == "1")
                    sValor = "ENERO";
                if (sValor == "2")
                    sValor = "FEBRERO";
                if (sValor == "3")
                    sValor = "MARZO";
                if (sValor == "4")
                    sValor = "ABRIL";
                if (sValor == "5")
                    sValor = "MAYO";
                if (sValor == "6")
                    sValor = "JUNIO";
                if (sValor == "7")
                    sValor = "JULIO";
                if (sValor == "8")
                    sValor = "AGOSTO";
                if (sValor == "9")
                    sValor = "SEPTIEMBRE";
                if (sValor == "10")
                    sValor = "OCTUBRE";
                if (sValor == "11")
                    sValor = "NOVIEMBRE";
                if (sValor == "12")
                    sValor = "DICIEMBRE";
            }

            return sValor;
        }
        private void GuardarOrbis(string _asOrigen, DataTable _aDt)
        {

        }
        public static string TurnoGlobal()
        {

            string sTurno = "2";
            DateTime dtFecha = DateTime.Now;
            if (dtFecha.Hour >= 6 && dtFecha.Hour < 16)
            {
                sTurno = "1";
            }

            return sTurno;
        }
        private bool CumpleHora(string _asHr1t, string _asHr2t)
        {
            bool bReturn;
            string sTurno = TurnoGlobal();
            string sHora = string.Empty;
            sHora = Convert.ToString(DateTime.Now.Hour);
            sHora += ":" + Convert.ToString(DateTime.Now.Minute).PadLeft(2, '0');

            if (sTurno == "1")
            {
                if (sHora == _asHr1t)
                    bReturn = true;
                else
                    bReturn = false;
            }
            else
            {
                if (sHora == _asHr2t)
                    bReturn = true;
                else
                    bReturn = false;
            }

            return bReturn;
        }

        #region regKanban&AttachFile
        private bool AttachFileKan(string _asFileAtt,string _asTurno)
        {
            bool bReturn = false;
            if (!File.Exists(_asFileAtt))
                return false;

            
            DateTime dtTime = DateTime.Now;
            int iHora = dtTime.Hour;
            string sHora = Convert.ToString(iHora);

            KanbanLogica kan = new KanbanLogica();
            kan.Fecha = dtTime;
            kan.Turno = _asTurno;
            string sHrReg = sHora.PadLeft(2, '0') + ":00";
            kan.Hora = sHrReg;
            DataTable data = KanbanLogica.ResumenKanbanTurno(kan);
            if(data.Rows.Count>0)
            {
                ExportarFormato(data,_asFileAtt,_asTurno);
            }

            bReturn = true;

            

            return bReturn;
        }
        private void ExportarFormato(DataTable _dt,string _asFile, string _asTurno)
        {
            Cursor = Cursors.WaitCursor;
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            try
            {
                
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(_asFile);
                oSheet = String.IsNullOrEmpty("Sheet1") ? (Excel._Worksheet)oWB.ActiveSheet : (Excel._Worksheet)oWB.Worksheets["Sheet1"];

                oSheet.Cells[4, 3] = string.Empty;
                for (int i = 7; i <= 17; i++)
                {
                    oSheet.Cells[i, 2] = string.Empty;
                    oSheet.Cells[i, 3] = string.Empty;
                    oSheet.Cells[i, 4] = string.Empty;
                    oSheet.Cells[i, 5] = string.Empty;
                    oSheet.Cells[i, 6] = string.Empty;
                }

                int iRow = 7;
                oSheet.Cells[4, 3] = Convert.ToString(DateTime.Now);
                for (int x = 0; x < _dt.Rows.Count; x++)
                {
                    string sHora = _dt.Rows[x][0].ToString();
                    string sClaseg = _dt.Rows[x][1].ToString();
                    string sClasey = _dt.Rows[x][2].ToString();
                    string sClaser = _dt.Rows[x][3].ToString();
                    string sClasen = _dt.Rows[x][4].ToString();

                    oSheet.Cells[iRow, 2] = sHora;
                    oSheet.Cells[iRow, 3] = sClaseg;
                    oSheet.Cells[iRow, 4] = sClasey;
                    oSheet.Cells[iRow, 5] = sClaser;
                    oSheet.Cells[iRow, 6] = sClasen;

                    iRow++;
                }

                oXL.DisplayAlerts = false;
                
                oWB.SaveAs(_asFile, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlShared);

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                Cursor = Cursors.Default;
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, Type.Missing, Type.Missing);
                    oXL.Quit();
                    Cursor = Cursors.Default;
                }
            }
            Cursor = Cursors.Default;
        }
        
        private void GeneraKanban(bool _abValida)
        {
            int iErrRow = 0;
            try
            {

                DataTable dt = ConfigCPROLogica.Consultar();
                string sKanban = dt.Rows[0]["ind_kanban"].ToString();
                if (string.IsNullOrEmpty(sKanban) || sKanban == "0")
                    return;

                string sKanDir = dt.Rows[0]["kanban_direc"].ToString();
                string sKanFile = dt.Rows[0]["kanban_file"].ToString();
                string sKanStart = dt.Rows[0]["kanban_start"].ToString();
                string sKanEnd = dt.Rows[0]["kanban_end"].ToString();
                int iMins = int.Parse(dt.Rows[0]["kanban_minutes"].ToString());

                int iAttaFileHr1 = 16; 
                int iAttaFileHr2 = 1; 
                string sKanAttFile = "KANBAN_RESUMEN";
                string sTurno = "2";

                DateTime dtTime = DateTime.Now;
                int iHora = dtTime.Hour;
                if (iHora >= 0 && iHora < 15)
                    sTurno = "1";

                string sHrStart = sKanStart.Substring(0, 1);
                int iHrStart = int.Parse(sHrStart);
                int iHrEnd = int.Parse(sKanEnd.Substring(0, 1));

                if (iHora > 23)
                    iHora -= 23;

                if (iHora < iHrStart && iHora > iHrEnd)
                    return;

                //VALIDA EL KARDEX DE KANBAN POR HORA
                KanbanLogica kan = new KanbanLogica();
                if (iHora >= 0 && iHora < 6)
                    dtTime = dtTime.AddDays(-1);

                kan.Fecha = dtTime;
                string sHrReg = Convert.ToString(iHora).PadLeft(2, '0') + ":00";
                kan.Hora = sHrReg;
                if(_abValida)
                {
                    if (KanbanLogica.Verificar(kan))
                        return;
                }
                

                bool bAttach = false;
                if (iAttaFileHr1 == iHora || iAttaFileHr2 == iHora)
                    bAttach = true;

                //SELECT NEXT FILE
                string sHrFile = string.Empty;
                iHora += 2;//CT zone

                if (iHora < 12)
                    sHrFile = iHora.ToString() + "am";
                else
                {
                    if (iHora > 23)
                    {
                        if (iHora == 24)
                            iHora = 12;
                        else
                            iHora -= 24;
                        sHrFile = iHora.ToString() + "am";
                    }
                    else
                    {
                        if (iHora > 12)
                            iHora -= 12;
                        sHrFile = iHora.ToString() + "pm";
                    }
                }

                sKanFile += "_" + sHrFile + ".xlsx";
                string sFile = sKanDir + @"\" + sKanFile;
                //VERIFY FILE EXIST
                long lFolio = 0;
                DataTable dtExc = new DataTable();
               if (!File.Exists(sFile))
                    return;
                else
                {
                    if (!FileLocked(sFile))
                        dtExc = getFromExcelEmp(sFile);

                    File.Delete(sFile);
                }

                lFolio = AccesoDatos.ConsecPRO("KANBAN");

                if (dtExc.Rows.Count > 0)
                {
                    //GET DATA FROM FILE
                    for (int x = 0; x < dtExc.Rows.Count; x++)
                    {
                        iErrRow = x;
                        string sLine = dtExc.Rows[x][0].ToString();
                        string sRpo = dtExc.Rows[x][1].ToString();

                        DateTime dtFecha = DateTime.Parse(dtExc.Rows[x][2].ToString());

                        string sItem = dtExc.Rows[x][3].ToString();
                        double dCant = double.Parse(dtExc.Rows[x][4].ToString());

                        DateTime dtPrint = DateTime.Today;
                        if (!string.IsNullOrEmpty(dtExc.Rows[x][5].ToString()))
                        {
                            if (!DateTime.TryParse(dtExc.Rows[x][5].ToString(), out dtPrint))
                                dtPrint = DateTime.Today;
                        }
                       
                        double dCantF = double.Parse(dtExc.Rows[x][9].ToString());
                        double dCantS = double.Parse(dtExc.Rows[x][10].ToString());
                        double dSaldo = double.Parse(dtExc.Rows[x][11].ToString());

                        KanbanDetLogica kdet = new KanbanDetLogica();
                        kdet.Folio = lFolio;
                        kdet.Line = sLine;
                        kdet.RPO = sRpo;
                        kdet.Creation = dtFecha;
                        kdet.Item = sItem;
                        kdet.Quantity = dCant;
                        kdet.Print = dtPrint;
                        if (DateTime.TryParse(dtExc.Rows[x][6].ToString(), out dtFecha))//1.0.0.12
                            kdet.Register = dtFecha;
                        if (DateTime.TryParse(dtExc.Rows[x][7].ToString(), out dtFecha))//1.0.0.9
                            kdet.Kanban = dtFecha;
                        if (DateTime.TryParse(dtExc.Rows[x][8].ToString(), out dtFecha))
                            kdet.Start = dtFecha;
                        kdet.QtyFinish = dCantF;
                        kdet.QtyShipped = dCantS;
                        kdet.Saldo = dSaldo;
                        kdet.Hora = sHrReg;

                        KanbanDetLogica.Guardar(kdet);
                    }
                }

                //save t_kanban
                kan.Planta = "EMPN";
                kan.Source = sFile;
                kan.Hora = sHrReg;
                kan.Turno = sTurno;
                kan.Folio = lFolio;

                KanbanLogica.Guardar(kan);

                /*ATTACH*/
                if(bAttach)
                {
                    string sFileAtt = sKanDir + @"\" + sKanAttFile +"_" + sTurno + "T.xlsx";
                    bAttach = AttachFileKan(sFileAtt,sTurno);
                }
            }
            catch (Exception ex)
            {
                string sEx = ex.ToString();
                MessageBox.Show(sEx + Environment.NewLine + iErrRow.ToString(), "GeneraKanban(bool _abValida)");
                //throw;
            }
        }
        #endregion

        #region regGlobals

        protected virtual bool FileLocked(string _asFile)
        {
            try
            {
                using (Stream stream = new FileStream(_asFile, FileMode.Open))
                    stream.Close();
                
            }
            catch (IOException)
            { 
                return true;
            }
            
            return false;
        }
        private void GeneraEnvios()
        {
            int iErrRow = 0;
            try
            {

                DataTable dt = ConfigCPROLogica.Consultar();
                string sKanban = dt.Rows[0]["ind_globals"].ToString();
                if (string.IsNullOrEmpty(sKanban) || sKanban == "0")
                    return;

                string sKanDir = dt.Rows[0]["kanban_direc"].ToString();
                string sKanFile = dt.Rows[0]["kanban_file"].ToString();
                string sTransFile = dt.Rows[0]["Kanban_file"].ToString(); //transfer_file
                string sKanStart = dt.Rows[0]["global_start"].ToString();
                string sKanEnd = dt.Rows[0]["global_end"].ToString();
                int iMins = int.Parse(dt.Rows[0]["global_min"].ToString());
                sKanFile = "TransferOrder";
                string sTurno = "2";

                DateTime dtTime = DateTime.Now;
                int iHora = dtTime.Hour;
                int iMin = dtTime.Minute;
                if (iHora >= 6 && iHora <= 16)
                    sTurno = "1";

                string sHrStart = sKanStart.Substring(0, 2);
                int iHrStart = int.Parse(sHrStart);
                int iHrEnd = int.Parse(sKanEnd.Substring(0, 2));

                if (iHora > 23)
                    iHora -= 23;

                if (iHora < iHrStart && iHora > iHrEnd)
                    return;

                //VALIDA EL KARDEX DE KANBAN POR HORA
                KanbanLogica kan = new KanbanLogica();
                if (iHora >= 0 && iHora < 6)
                    dtTime = dtTime.AddDays(-1);

                kan.Fecha = dtTime;
                string sHrReg = Convert.ToString(iHora).PadLeft(2, '0') + ":00";
                kan.Hora = sHrReg;

                
                //SELECT NEXT FILE
                string sHrFile = string.Empty;
                string sHora = iHora.ToString();
                string sTipo = "am";

                iHora += 2;//CT zone

                if (iHora < 12)
                    sTipo = "am";
                else
                {
                    if (iHora > 23)
                    {
                        if (iHora == 24)
                            iHora = 12;
                        else
                            iHora -= 24;
                        sTipo = "am";
                    }
                    else
                    {
                        if (iHora > 12)
                            iHora -= 12;
                        sTipo = "pm";
                    }
                }

                sHrFile = iHora.ToString() + sTipo;
                
                sTransFile = sKanFile + " " + sHrFile + ".xlsx";
                string sFile = sKanDir + @"\" + sTransFile;
                //VERIFY FILE EXIST
                long lFolio = 0;
                DataTable dtExc = new DataTable();
                if (!File.Exists(sFile))
                {
                    sHrFile = iHora.ToString() + "30" + sTipo;
                    sTransFile = sKanFile + " " + sHrFile + ".xlsx";
                    sFile = sKanDir + @"\" + sTransFile;
                    if (!File.Exists(sFile))
                        sFile = string.Empty;
                }

                //Listar globals pendientes de transfer order data [TO y Truck]
                DataTable dtGlob = new DataTable();
                if (!string.IsNullOrEmpty(sFile))
                {
                    dtGlob = GlobalRpoLogica.MonitorGlobals();
                    if (dtGlob.Rows.Count > 0)
                    {           
                        if (!FileLocked(sFile))
                        {
                            dtExc = getFromExcelTrans(sFile);
                            File.Delete(sFile);
                        }
                    }
                    else
                    {
                        if (iMin >= 55)
                            File.Delete(sFile);
                    }
                }

                if (dtExc.Rows.Count > 0)
                {
                    DataTable dtTO = new DataTable("ToTruck");
                    dtTO.Columns.Add("to", typeof(string));//0
                    dtTO.Columns.Add("truck", typeof(string));//1
                    dtTO.Columns.Add("f_close", typeof(string));

                    for (int i = 0; i < dtExc.Rows.Count; i++)
                    {
                        string sTO = dtExc.Rows[i][4].ToString();
                        string sTruck2 = dtExc.Rows[i][2].ToString();
                        string sFclose = dtExc.Rows[i][1].ToString();
                        if (string.IsNullOrEmpty(sTruck2) || string.IsNullOrEmpty(sTO))
                            continue;

                        dtTO.Rows.Add(sTO, sTruck2,sFclose);
                    }

                    for (int r = 0; r < dtGlob.Rows.Count; r++) // globals pendientes de captura TO
                    {
                        long lFoliog = long.Parse(dtGlob.Rows[r][0].ToString());
                        string sRPO = dtGlob.Rows[r][1].ToString();
                        string sModelo = dtGlob.Rows[r][3].ToString();
                        double dCant = double.Parse(dtGlob.Rows[r][5].ToString());
                        string sEnvSal = dtGlob.Rows[r][11].ToString();
                        //registros de TransferOrder.xlsx

                        for (int x = 0; x < dtExc.Rows.Count; x++)
                        {
                            iErrRow = x;

                            string sRpo = dtExc.Rows[x][7].ToString();
                            if (sRpo != sRPO)
                                continue;

                            string sTruck = dtExc.Rows[x][2].ToString();
                            string sTransfer = dtExc.Rows[x][4].ToString();
                            string sItem = dtExc.Rows[x][5].ToString();
                            string sFtransfer = dtExc.Rows[x][1].ToString();
                            // string sPallet = dtExc.Rows[x][6].ToString();

                            double dCantF = double.Parse(dtExc.Rows[x][8].ToString());
                            double dCantS = double.Parse(dtExc.Rows[x][9].ToString());

                            if (dCantS > dCant)
                                continue;

                            if (string.IsNullOrEmpty(sTruck))
                            {
                                for (int i = 0; i < dtTO.Rows.Count; i++)
                                {
                                    string sTO = dtTO.Rows[i][0].ToString();
                                    string sTruck2 = dtTO.Rows[i][1].ToString();
                                    string sFclose2 = dtTO.Rows[i][2].ToString();
                                    if (sTransfer == sTO && (!string.IsNullOrEmpty(sTruck2)))
                                    {
                                        sTruck = sTruck2;
                                        sFtransfer = sFclose2;
                                        break;
                                    }
                                }
                            }

                            GlobalRpoLogica glob = new GlobalRpoLogica();
                            glob.Folio = lFoliog;

                            DateTime dtFecha = DateTime.Today;
                            if (DateTime.TryParse(dtExc.Rows[x][0].ToString(), out dtFecha))//1.0.0.12
                                glob.Fscanned = dtFecha;
                                
                            if (string.IsNullOrEmpty(sEnvSal))
                            {
                                if (DateTime.TryParse(sFtransfer, out dtFecha))
                                {
                                    /*
                                    sFechaf = dtFecha.ToString();
                                    if (sFechaf.Trim() == "10:00:00 PM")
                                        glob.Ftransfer = glob.Fscanned;
                                    else
                                    {
                                        
                                        glob.Ftransfer = dtFecha;
                                    }
                                    */
                                    glob.Ftransfer = dtFecha;
                                }
                            }
                            
                            if (DateTime.TryParse(dtExc.Rows[x][3].ToString(), out dtFecha))//1.0.0.9
                                glob.Fposted = dtFecha;

                            glob.Truck = sTruck;
                            glob.Transfer = sTransfer;

                            GlobalRpoLogica.GuardarTrans(glob);
                            x = dtExc.Rows.Count;
                        }

                    }   
                }
            }
            catch (Exception ex)
            {
                string sEx = ex.ToString();
                MessageBox.Show(sEx + Environment.NewLine + iErrRow.ToString(), "GeneraGlobals()");
                //throw;
            }
        }
        private void GeneraGlobals()
        {
            int iErrRow = 0;
            try
            {

                DataTable dt = ConfigCPROLogica.Consultar();
                string sKanban = dt.Rows[0]["ind_globals"].ToString();
                if (string.IsNullOrEmpty(sKanban) || sKanban == "0")
                    return;

                string sKanDir = dt.Rows[0]["kanban_direc"].ToString();
                string sKanFile = dt.Rows[0]["kanban_file"].ToString();
                string sKanStart = dt.Rows[0]["global_start"].ToString();
                string sKanEnd = dt.Rows[0]["global_end"].ToString();
                int iMins = int.Parse(dt.Rows[0]["global_min"].ToString());

                string sTurno = "2";

                DateTime dtTime = DateTime.Now;
                int iHora = dtTime.Hour;
                int iMin = dtTime.Minute;
                if (iHora >= 6 && iHora <= 16)
                    sTurno = "1";

                string sHrStart = sKanStart.Substring(0, 2);
                int iHrStart = int.Parse(sHrStart);
                int iHrEnd = int.Parse(sKanEnd.Substring(0, 2));

                if (iHora > 23)
                    iHora -= 23;

                if (iHora < iHrStart && iHora > iHrEnd)
                    return;

                //VALIDA EL KARDEX DE KANBAN POR HORA
                KanbanLogica kan = new KanbanLogica();
                if (iHora >= 0 && iHora < 6)
                    dtTime = dtTime.AddDays(-1);

                kan.Fecha = dtTime;
                string sHrReg = Convert.ToString(iHora).PadLeft(2, '0') + ":00";
                kan.Hora = sHrReg;

                if (!KanbanLogica.VerificarGlobals(kan))
                    return;
                
                //SELECT NEXT FILE
                string sHrFile = string.Empty;
                string sMnFile = "30";
                string sHora = iHora.ToString();
                sHora += ":" + sMnFile; //hora salida prod & entrada env

                iHora += 2;//CT zone
                
                if (iHora < 12)
                    sHrFile = iHora.ToString() + sMnFile + "am";
                else
                {
                    if (iHora > 23)
                    {
                        if (iHora == 24)
                            iHora = 12;
                        else
                            iHora -= 24;
                        sHrFile = iHora.ToString() + sMnFile + "am";
                    }
                    else
                    {
                        if (iHora > 12)
                            iHora -= 12;
                        sHrFile = iHora.ToString() + sMnFile + "pm";
                    }
                }
                
                sKanFile += "_" + sHrFile + ".xlsx";
                string sFile = sKanDir + @"\" + sKanFile;
                //VERIFY FILE EXIST
                long lFolio = 0;
                DataTable dtExc = new DataTable();
                if (!File.Exists(sFile))
                    return;
                else
                {
                    if(GlobalRpoLogica.VerificarListado())
                    {
                        if(iMin >= 30)
                        {
                            if(!FileLocked(sFile))
                            {
                                dtExc = getFromExcelEmp(sFile);
                                File.Delete(sFile);
                            }
                            
                        }
                    }
                    else
                    {
                        if(iMin >= 55)
                            File.Delete(sFile);
                    }
                }
                
                if (dtExc.Rows.Count > 0)
                {
                    //GET DATA FROM FILE
                    for (int x = 0; x < dtExc.Rows.Count; x++)
                    {
                        iErrRow = x;
                        string sLine = dtExc.Rows[x][0].ToString();
                        string sRpo = dtExc.Rows[x][1].ToString();

                        GlobalRpoLogica rpo = new GlobalRpoLogica();
                        rpo.RPO = sRpo;
                        if (!GlobalRpoLogica.Verificar(rpo))
                            continue;

                        DataTable data = GlobalRpoLogica.Consultar(rpo);
                        lFolio = long.Parse(data.Rows[0]["folio"].ToString());
                        DateTime dtFecha = DateTime.Today;
                        DateTime dtFecha2 = DateTime.Today;
                        if (DateTime.TryParse(dtExc.Rows[x][2].ToString(), out dtFecha))
                            dtFecha2 = dtFecha;

                        DateTime dtPrint = DateTime.Today;
                        if (!string.IsNullOrEmpty(dtExc.Rows[x][5].ToString()))
                        {
                            if (DateTime.TryParse(dtExc.Rows[x][5].ToString(), out dtPrint))//1.0.0.12
                                rpo.Print = dtPrint;
                        }
                        
                        double dCantF = double.Parse(dtExc.Rows[x][9].ToString());
                        double dCantS = double.Parse(dtExc.Rows[x][10].ToString());
                        
                        rpo.Folio = lFolio;
                        rpo.Print = dtPrint;
                        if (DateTime.TryParse(dtExc.Rows[x][6].ToString(), out dtFecha))//1.0.0.12
                            rpo.Register = dtFecha;
                        if (DateTime.TryParse(dtExc.Rows[x][7].ToString(), out dtFecha))//1.0.0.9
                            rpo.Kanban = dtFecha;
                        if (DateTime.TryParse(dtExc.Rows[x][8].ToString(), out dtFecha))
                            rpo.Start = dtFecha;
                        rpo.Finish = dCantF;
                        rpo.Shipped = dCantS;
                        rpo.HorKan = sHora;

                        GlobalRpoLogica.Guardar(rpo);
                    }
                }
            }
            catch (Exception ex)
            {
                string sEx = ex.ToString();
                MessageBox.Show(sEx + Environment.NewLine + iErrRow.ToString(), "GeneraGlobals()");
                //throw;
            }
        }
        #endregion
        private void GeneraActivos(bool _abValida)
        {
            DataTable dt = ConfigLogica.Consultar();
            string sActivos = dt.Rows[0]["ind_genact"].ToString();
            string sDirec = dt.Rows[0]["direc_act"].ToString();
            string sFile = dt.Rows[0]["nombre_act"].ToString();
            string sHr1t = dt.Rows[0]["hr_1t"].ToString();
            string sHr2t = dt.Rows[0]["hr_2t"].ToString();
            string sCargar = dt.Rows[0]["cargar_actorbis"].ToString();

            //if (_abValida && !CumpleHora(sHr1t, sHr2t))
            //    return;

            //generar datos
            TressActivos act = new TressActivos();
            dt = TressActivos.Consultar(act);
            if (dt.Rows.Count > 0)
            {
                if (sCargar == "1")
                    GuardarOrbis("ACT", dt);
                else
                {
                    ExportarTexto(DateTime.Today, "ACT", sDirec, sFile, dt);
                    //IF( EXPORTAR ) > SAVER DATE IN DWDATA
                }
            }
        }
        private void GeneraAsistencia(DateTime _dtFecha, bool _abValida)
        {
            DataTable dt = ConfigLogica.Consultar();
            string sAsistencia = dt.Rows[0]["ind_genasis"].ToString();
            string sActivos = dt.Rows[0]["ind_genasis"].ToString();
            string sDirec = dt.Rows[0]["direc_asis"].ToString();
            string sFile = dt.Rows[0]["nombre_asis"].ToString();
            string sHr1t = dt.Rows[0]["hr_1tasis"].ToString();
            string sHr2t = dt.Rows[0]["hr_2tasis"].ToString();
            string sCargar = dt.Rows[0]["cargar_asisorbis"].ToString();
            int iGenMin = int.Parse(dt.Rows[0]["asis_genmin"].ToString());

            if (_abValida && !CumpleHora(sHr1t, sHr2t))
            {
                if(sCargar == "1" && iGenMin > 0)
                {
                    if (!ScheduleMin("ASIS",iGenMin))
                        return;
                }
                else
                    return;
            }

            GeneraActivos(true);
            string sTurno = TurnoGlobal();

            TressActivos act = new TressActivos();
            act.Turno = sTurno;
            act.Fecha = _dtFecha;
            dt = TressActivos.ConsultarAsisHrs(act);
            if (dt.Rows.Count > 0)
                ExportarTexto(_dtFecha,"ASIS", sDirec, sFile, dt);
            
        }
        private bool ScheduleMin(string _asProceso, int iMin)
        {
            bool bReturn = false;
            DateTime dtFecha = DateTime.Now;

            KardexLogica kar = new KardexLogica();
            kar.Proceso = _asProceso;
            kar.Fecha = dtFecha;

            DateTime dtIni = KardexLogica.ConsultaDiaGen(kar);
            DateTime dtFin = DateTime.Now;
            TimeSpan time = dtFin - dtIni;
            int iCant = (int)time.TotalMinutes;

            if (iCant >= iMin)
            {
                if (DateTime.Now.Hour >= 3 && DateTime.Now.Hour <= 5)
                    bReturn = false;
                else
                    bReturn = true;
            }
            else
            {
                if (DateTime.Now.Hour == 7 || DateTime.Now.Hour == 16 )
                {
                    if (DateTime.Now.Minute >= 18 && DateTime.Now.Minute <= 22)
                    {
                        kar.Hora = DateTime.Now.Hour.ToString();
                        if (KardexLogica.ValidaDiaHoraGen(kar))
                            bReturn = true;
                    }   
                }
            }

            return bReturn;
                
        }
        #endregion

        #region regBotones
        private void btnConfig_Click(object sender, EventArgs e)
        {
            wfConfig wfConf = new wfConfig();
            wfConf.ShowDialog();

            CargarData();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            
            //if (dgwData.Rows.Count == 0)
            //    return;

            //int iRow = dgwData.CurrentCell.RowIndex;
            //if (dgwData.CurrentRow.Index == -1)
            //    return;

            //if (iRow == 0)
            //    GeneraActivos(false);
            //else
            //{
            //    if (iRow == 1)
            //    {
            //        CapturaPop wfPop = new CapturaPop();
            //        wfPop.ShowDialog();
            //        DateTime dtFecha = wfPop._dtReturn;
            //        GeneraAsistencia(dtFecha,false);
            //    }
            //}
            
        }


        #endregion

        private void btGlobal_Click(object sender, EventArgs e)
        {
            //GeneraKanban(false);
            //GeneraGlobals();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneraEnvios();
        }

        #region regActividadesIT
        private void btnMensual_Click(object sender, EventArgs e)
        {
            _iAxo = DateTime.Today.Year;
            _iMes = DateTime.Today.Month;
            if (_iMes == 1)
            {
                _iMes = 12;
                _iAxo--;
            }
            else
                _iMes--;
            
            /*default*/
            _iAxo = 2018;
            _iMes = 12;

            string sAxo = _iAxo.ToString();
            string sMes = string.Empty;
            //if (iMes == 11)
            //    sMes = "NOVIEMBRE";

            string monthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(_iMes);

            string sDirec = @"C:\Users\agonz0\Documentos  Personales\Actividades\" + sAxo + @"\" + monthName;
            OpenFileDialog fileOpen = new OpenFileDialog();
            string sPath = sDirec;
            /*
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                sPath = folderBrowserDialog1.SelectedPath;
            }
            */
            string sMesFile = monthName + " " + sAxo + ".xlsx";
            string sFile = sPath + @"\" + sMesFile;

            ActuserLogicaIT act = new ActuserLogicaIT();
            act.Usuario = _sUsuario;
            act.Axo = _iAxo;
            act.Mes = _iMes;
            if (!ActuserLogicaIT.Verificar(act))
            {
                if (!File.Exists(sFile))
                    return;

                DataTable dtExc = new DataTable("semanal");
                dtExc.Columns.Add("semana", typeof(string));
                dtExc.Columns.Add("dia", typeof(string));
                dtExc.Columns.Add("catego", typeof(string));
                dtExc.Columns.Add("depto", typeof(string));
                dtExc.Columns.Add("mins", typeof(string));
                dtExc.Columns.Add("clave", typeof(string));
                dtExc.Columns.Add("descrip", typeof(string));
                dtExc.Columns.Add("solicita", typeof(string));

                for (int i = 1; i <= 5; i++)
                {
                    string sWeekFile = monthName + " W0" + i.ToString() + ".xlsx";
                    string sFileW = sPath + @"\" + sWeekFile;
                    if (!File.Exists(sFileW))
                        continue;


                    dtExc = getFromExcelWeek(sFileW, dtExc);
                }

                for (int x = 0; x < dtExc.Rows.Count; x++)
                {
                    //setExcelMonth(sFile,dtExc);
                    string sSem = dtExc.Rows[x][0].ToString();
                    string sDia = dtExc.Rows[x][1].ToString();
                    string sCatego = dtExc.Rows[x][2].ToString();
                    string sDepto = dtExc.Rows[x][3].ToString();
                    long lMins = long.Parse(dtExc.Rows[x][4].ToString());
                    string sSistema = dtExc.Rows[x][5].ToString();
                    string sDescrip = dtExc.Rows[x][6].ToString();
                    string sSolicita = dtExc.Rows[x][7].ToString();

                    string sFecha = _iMes.ToString() + "/" + sDia + "/" + sAxo;
                    DateTime dtFecha = DateTime.Parse(sFecha);

                    act.Fecha = dtFecha;
                    act.Consec = 0;
                    act.Cant = 1;
                    act.Sistema = sSistema;
                    act.Actividad = sDescrip;
                    act.Solicita = sSolicita;
                    act.Minutos = lMins;
                    act.Catego = sCatego;
                    act.Depto = sDepto;
                    act.Semana = int.Parse(sSem);

                    if (ActuserLogicaIT.Guardar(act) > 0)
                        continue;
                }
            }
            for (int i = 1; i <= 5; i++)
            {
                string sWeekFile = monthName + " W0" + i.ToString() + ".xlsx";
                string sFileW = sPath + @"\" + sWeekFile;
                if (!File.Exists(sFileW))
                    continue;

                setExcelWeek(sFileW, i);
            }

            setExcelMonth(sFile);

            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))// && p.StartTime.AddSeconds(+60) < DateTime.Now)
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }
        private void setExcelWeek(string _asArchivo,int _aiWeek)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();
                string sTipo = string.Empty;
                string sSemana = string.Empty;
                int iSheets = xlWorkbook.Sheets.Count;                

                xlWorksheet = xlWorkbook.Sheets[6]; //Reporte semanal por proyecto
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = 50; //xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                    
                ActuserLogicaIT act = new ActuserLogicaIT();
                act.Usuario = _sUsuario;
                act.Axo = _iAxo;
                act.Mes = _iMes;
                act.Semana = _aiWeek;
                DataTable _data = ActuserLogicaIT.ActividadProyectoCat(act);

                string sSisAnt = string.Empty;
                int iExc = 3;
                string[] sCategos = { "0", "0","A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "M" };

                for (int x = 0; x < _data.Rows.Count; x++) //reporte semanal x dia
                {
                    string sSistema = _data.Rows[x][0].ToString();
                    if (string.IsNullOrEmpty(sSistema) || sSistema == "0")
                        continue;

                    if (string.IsNullOrEmpty(sSisAnt))
                        sSisAnt = sSistema;

                    if (sSistema != sSisAnt)
                        iExc++;

                    string sDescrip = _data.Rows[x][1].ToString();
                    string sCat = _data.Rows[x][2].ToString();
                    int iCant = int.Parse(_data.Rows[x][3].ToString());
                    int iCol = Array.IndexOf(sCategos, sCat);

                    xlRange.Cells[iExc, 1].Value2 = sDescrip;
                    xlRange.Cells[iExc, iCol].Value2 = iCant.ToString();

                    sSisAnt = sSistema;
                }
                if (iExc >= 3)
                    iExc -= 2;

                xlRange.Cells[20,2].Value2 = iExc.ToString();

                xlApp.DisplayAlerts = false;
                xlWorkbook.Save();
                xlWorkbook.Close(true);
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ex.ToString();
                MessageBox.Show(ex.ToString() + Environment.NewLine + "setExcelMonth(" + _asArchivo + ")");

            }
        }

        private void setExcelMonth(string _asArchivo)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();
                string sTipo = string.Empty;
                string sSemana = string.Empty;
                int iSheets = xlWorkbook.Sheets.Count;
                for (int s = 1; s <= 4; s++) //mensual
                {
                    if (s < 2)
                        continue;
                    
                    xlWorksheet = xlWorkbook.Sheets[s]; //Caluclos Sheet
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = 50; //xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;
                    if(s == 2)
                    {
                        ActuserLogicaIT act = new ActuserLogicaIT();
                        act.Usuario = _sUsuario;
                        act.Axo = _iAxo;
                        act.Mes = _iMes;
                        DataTable _data = ActuserLogicaIT.ActividadMesCat(act);

                        string sSisAnt = string.Empty;
                        int iExc = 3;
                        string[] sCategos = { "0","0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J" };

                        for (int x = 0; x < _data.Rows.Count; x++) //reporte semanal x dia
                        {
                            string sSistema = _data.Rows[x][0].ToString();
                            if (string.IsNullOrEmpty(sSistema) || sSistema == "0")
                                continue;

                            if (string.IsNullOrEmpty(sSisAnt))
                                sSisAnt = sSistema;

                            if (sSistema != sSisAnt)
                                iExc++;

                            string sAct = _data.Rows[x][1].ToString();
                            string sCat = _data.Rows[x][2].ToString();
                            int iCant = int.Parse(_data.Rows[x][3].ToString());
                            int iCol = Array.IndexOf(sCategos, sCat);

                            xlRange.Cells[iExc, 1].Value2 = sAct;
                            xlRange.Cells[iExc, iCol].Value2 = iCant.ToString();

                            sSisAnt = sSistema;
                        }
                        if (iExc >= 3)
                            iExc -= 2;

                        xlRange.Cells[23, 1].Value2 = iExc.ToString();
                    }
                    if (s == 4) //-->> Productividad
                    {
                        rowCount = 10;
                        for (int i = 3; i < rowCount; i++)
                        {
                            if (xlRange.Cells[i, 2].Value2 == null)//--categoria / depto
                            {
                                sTipo = string.Empty;
                                i = rowCount;
                                continue;
                            }

                            if (xlRange.Cells[i, 2].Value2 != null)
                                sTipo = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());

                            sTipo = sTipo.TrimEnd();
                            if (sTipo.IndexOf("Sem") == -1)
                                continue;

                            string sSem = sTipo.Substring(sTipo.Length - 1);
                            int iSem = 0;
                            if (!int.TryParse(sSem, out iSem))
                                iSem = 0;

                            if (iSem == 0)
                                continue;

                            ActuserLogicaIT act = new ActuserLogicaIT();
                            act.Usuario = _sUsuario;
                            act.Axo = _iAxo;
                            act.Mes = _iMes;
                            act.Semana = iSem;

                            DataTable dt = ActuserLogicaIT.ActividadSemanal(act);
                            if(dt.Rows.Count > 0)
                            {
                                int iCant = Int32.Parse(dt.Rows[0][0].ToString());
                                long lMin = long.Parse(dt.Rows[0][1].ToString());
                                xlRange.Cells[i,3].Value2 = iCant.ToString();
                                xlRange.Cells[i, 4].Value2 = lMin.ToString();
                            }
                            
                        }
                    }
                    if( s == 3) //-->> CALCULOS
                    {
                        //categorias
                        for (int i = 2; i < rowCount; i++)
                        {
                            if (xlRange.Cells[i, 1].Value2 == null)//--categoria / depto
                            {
                                sTipo = string.Empty;
                                //i = rowCount;
                                continue;
                            }

                            if (i == 2)
                            {
                                if (xlRange.Cells[i, 1].Value2 != null)
                                    sTipo = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());
                                if (xlRange.Cells[i, 2].Value2 != null)
                                    sSemana = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());

                                if (sTipo == "Categorias")
                                    continue;
                            }
                            else
                            {
                                if (i > 18)
                                {
                                    if (xlRange.Cells[i, 1].Value2 != null)
                                        sTipo = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());
                                }
                            }

                            int iDiaAdd = 2;
                            int iDiasW = 6;

                            if (sTipo == "Categorias")
                            {
                                string sCategoria = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());                                                                                                           //conting by day

                                ActuserLogicaIT act = new ActuserLogicaIT();
                                act.Catego = sCategoria;
                                act.Usuario = _sUsuario;
                                act.Axo = _iAxo;
                                act.Mes = _iMes;
                                DataTable _data = ActuserLogicaIT.ActividadMensualCat(act);

                                for (int x = 0; x < _data.Rows.Count; x++) //reporte semanal x dia
                                {
                                    int iWeek = int.Parse(_data.Rows[x][0].ToString());
                                    int iDia = int.Parse(_data.Rows[x][1].ToString());
                                   
                                    int iCant = int.Parse(_data.Rows[x][2].ToString());
                                    
                                    iWeek--;
                                    int iDiaW = iWeek * iDiasW; // 2 * 6 = 12
                                    int iDiaCol = iDiaAdd + iDiaW + iDia; // 14 + 4 

                                    int iCantAnt = 0;
                                   
                                    iCant += iCantAnt;
                                    xlRange.Cells[i, iDiaCol].Value2 = iCant.ToString();
                                }
                            }
                            else
                            { // Departamentos
                                if (xlRange.Cells[i, 2].Value2 == null)//--categoria / depto
                                {
                                    continue;
                                }

                                string sDepto = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());// A - B - C - D - E - F - G - H - I - J //

                                ActuserLogicaIT act = new ActuserLogicaIT();
                                act.Depto = sDepto;
                                act.Usuario = "agonz0";
                                act.Axo = 2018;
                                act.Mes = 10;
                                DataTable _data = ActuserLogicaIT.ActividadMensualDepto(act);

                                for (int x = 0; x < _data.Rows.Count; x++) //reporte semanal x dia
                                {
                                    int iWeek = int.Parse(_data.Rows[x][0].ToString());
                                    int iDia = int.Parse(_data.Rows[x][1].ToString());
                                    int iCant = int.Parse(_data.Rows[x][2].ToString());

                                    iWeek--;
                                    int iDiaW = iWeek * iDiasW;
                                    int iDiaCol = iDiaAdd + iDiaW + iDia;
                                    xlRange.Cells[i, iDiaCol].Value2 = iCant.ToString();
                                }
                            }
                        }
                    }
                }

                xlApp.DisplayAlerts = false;
                xlWorkbook.Save();
                xlWorkbook.Close(true);
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ex.ToString();
                MessageBox.Show(ex.ToString() + Environment.NewLine + "setExcelMonth(" + _asArchivo + ")");

            }
        }

        private DataTable getFromExcelWeek(string _asArchivo,DataTable _dt)
        {
            int iExCont = 0;
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo,true);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();

                string sSemana = _asArchivo.Substring(_asArchivo.Length - 6,1);
                
                string sValue = string.Empty;
                
                int iSheets = xlWorkbook.Sheets.Count;
                for(int s = 1; s <= 5; s++)
                {
                    xlWorksheet = xlWorkbook.Sheets[s];
                    
                    string sDia = xlWorksheet.Name.Substring(3, 2);

                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    for (int i = 7; i <= rowCount; i++)
                    {
                        iExCont = i;

                        string sCatego = string.Empty;
                        string sDepto = string.Empty;
                        string sMins = string.Empty;
                        string sServicio = string.Empty;
                        string sSolicita = string.Empty;
                        string sProyecto = string.Empty;
                        
                        sValue = string.Empty;

                        if (xlRange.Cells[i, 2].Value2 == null)//-- cantidad --//
                            continue;

                        if (xlRange.Cells[i, 5].Value2 == null)//-- minutos --//
                            continue;
                        
                        if (xlRange.Cells[i, 6].Value2 == null)//-- categoria --//
                            continue;

                        if (xlRange.Cells[i, 3].Value2 != null)
                            sValue = Convert.ToString(xlRange.Cells[i, 3].Value2.ToString());

                        if (sValue == "Totales")
                        {
                            i = rowCount;
                            continue;
                        }
                        if (xlRange.Cells[i, 3].Value2 != null)
                            sServicio = Convert.ToString(xlRange.Cells[i, 3].Value2.ToString());
                        if (xlRange.Cells[i, 4].Value2 != null)
                            sSolicita = Convert.ToString(xlRange.Cells[i, 4].Value2.ToString());
                        if (xlRange.Cells[i, 5].Value2 != null)
                            sMins = Convert.ToString(xlRange.Cells[i, 5].Value2.ToString());
                        if (xlRange.Cells[i, 6].Value2 != null)
                            sCatego = Convert.ToString(xlRange.Cells[i, 6].Value2.ToString());
                        if (xlRange.Cells[i, 7].Value2 != null)
                            sDepto = Convert.ToString(xlRange.Cells[i, 7].Value2.ToString());
                        if (xlRange.Cells[i, 8].Value2 != null)
                            sProyecto = Convert.ToString(xlRange.Cells[i, 8].Value2.ToString());

                        _dt.Rows.Add(sSemana,sDia,sCatego,sDepto,sMins,sProyecto,sServicio,sSolicita);

                    }
                }
                
                xlApp.DisplayAlerts = false;
                xlWorkbook.Close(false);
                xlApp.DisplayAlerts = true;
                xlApp.Quit();
                xlApp = null;

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ex.ToString();
                MessageBox.Show(ex.ToString() + Environment.NewLine + iExCont.ToString(), "getFromExcelWeek("+_asArchivo+")");

            }

            return _dt;
        }
        #endregion
        #region regHeadCount
        private void btHeadc_Click(object sender, EventArgs e)
        {
            string sArchivo = @"C:\Users\agonz0\Documentos  Personales\CloverPro\SIMULADOR DE HC.xlsx";
            getFromExcelHC(sArchivo);
        }
        private void getFromExcelHC(string _asArchivo)
        {
            int iExCont = 0;
            try
            {
                Cursor.Current = Cursors.WaitCursor;

               

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo, true);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();

                string sValue = string.Empty;

                int iSheets = xlWorkbook.Sheets.Count;
                
                xlWorksheet = xlWorkbook.Sheets["Lines COLOR"];

                //                string sDia = xlWorksheet.Name.Substring(3, 2);

                //Excel.Range xlRange = xlWorksheet.UsedRange;
                xlWorksheet.Select();
                Excel.Range xlRange = xlWorksheet.get_Range("A1", "M600");
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                string sLinea = string.Empty;
                for (int i = 4; i <= rowCount; i++)
                {
                    iExCont = i;
                    string sModelo = string.Empty;
                    double dStd1 = 0;
                    double dStd2 = 0;
                    double dFactor = 0;
                    int iHc = 0;

                    sValue = string.Empty;

                    if (xlRange.Cells[i, 2].Value2 == null)//-- linea/modelo --//
                        continue;

                    if (xlRange.Cells[i, 2].Value2 == null && xlRange.Cells[i, 3].Value2 == null)
                    { 
                        i = rowCount;
                        continue;
                    }

                    
                    sValue = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());//linea
                    if(sValue.IndexOf("MX1A") != -1)
                    {
                        sLinea = sValue;
                        continue;
                    }
                    if (xlRange.Cells[i, 2].Value2 != null)
                        sModelo = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());

                    if (!double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out dStd1))
                        dStd1 = 0;

                    if (!double.TryParse(xlRange.Cells[i, 4].Value2.ToString(), out dStd2))
                        dStd2 = 0;

                    if (!double.TryParse(xlRange.Cells[i, 5].Value2.ToString(), out dFactor))
                        dFactor = 0;

                    if (xlRange.Cells[i, 10].Value2 != null)//-- head count --//
                    {
                        if (!int.TryParse(xlRange.Cells[i, 10].Value2.ToString(), out iHc))
                            iHc = 0;
                    }
                   
                    int iCont = 0;
                    string[] sLinex = new string[5];
                    string sLin = sLinea;
                    while (!string.IsNullOrEmpty(sLin))
                    {
                        int iId = sLin.IndexOf("-");
                        if ((iId > 0))
                        {
                            string sLine = sLin.Substring(iId + 1);

                            sLinex[iCont] = sLine.Trim();
                            sLin = sLin.Substring(0, iId - 1).Trim();
                            
                            iCont++;
                        }
                        else
                        {
                            sLinex[iCont] = sLin.Trim();
                            sLin = string.Empty;
                        }
                    }
                    
                    for (int x = 0; x <= iCont; x++)
                    {
                        ModeloHcLogica mod = new ModeloHcLogica();
                        mod.Planta = "COL";
                        mod.Linea = sLinex[x].ToString();
                        mod.Modelo = sModelo;
                        mod.HeadCount = iHc;
                        mod.Standard1 = dStd1;
                        mod.Standard2 = dStd2;
                        mod.Factor = dFactor;
                        ModeloHcLogica.Guardar(mod);
                    }
                    
                }
                
                xlApp.DisplayAlerts = false;
                xlWorkbook.Close(false);
                xlApp.DisplayAlerts = true;
                xlApp.Quit();
                xlApp = null;

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                ex.ToString();
                MessageBox.Show(ex.ToString() + Environment.NewLine + iExCont.ToString(), "getFromExcelWeek(" + _asArchivo + ")");

            }
        }
        #endregion


        private void GeneraInvCiclico()
        {
            BinContLogica bincont = new BinContLogica();
            Cursor = Cursors.WaitCursor;
            try
            {
                string sPta = "EMPN";
                string sLine = string.Empty;
                switch (sPta)
                {
                    case "EMPN":
                        sLine = "TPACKA";
                        break;
                    case "COL":
                        sLine = "CTNR";
                        break;
                    case "MON":
                        sLine = "TNR";
                        break;
                    case "INKM":
                        sLine = "INKM";
                        break;
                    case "INKP":
                        sLine = "INKP";
                        break;
                    case "FUS":
                        sLine = "FUS";
                        break;
                }


                DateTime dtTime = DateTime.Now;
                int iHora = dtTime.Hour;
                string sHoraG = Convert.ToString(iHora);
                string sHrReg = sHoraG.PadLeft(2, '0') + ":00";
                bincont.hora = sHrReg;

                string sHora = getFileTime();
                string sArchivo = "Warehouse Bin Contents " + sHora;// "Contador1";
                if (sLine != "TPACKA")
                    sArchivo += sLine.ToLower();

                sArchivo = _lsPath + @"\" + sArchivo + ".csv";
                if (!File.Exists(sArchivo))
                {
                    Cursor = Cursors.Default;
                    return;
                }

                if (BinContLogica.VerificarRegistros(bincont))
                {
                    return;
                }

                DataTable dt = LoadFile(sArchivo);
                               

                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    //eliminar BIN_CODE <> TPACKA
                    string sBin = dt.Rows[x][0].ToString();
                    if (!sBin.StartsWith(sLine))
                    {
                        continue;
                    }
                    string sBinLine = sBin.Substring(sBin.Length - 2);
                    int iP = 0;
                    if (!int.TryParse(sBinLine.Substring(0, 1), out iP))
                        sBinLine = sBinLine.Substring(1);

                    
                    if (sBinLine == "0")
                    {
                        continue;
                    }
                   

                    //eliminar sin saldo
                    int iCant = 0;
                    if (!int.TryParse(dt.Rows[x][5].ToString(), out iCant))//QUANTITY --> BEFORE //Available qty to take 18
                        iCant = 0;

                    if (iCant > 0)
                    {
                        bincont.bincode = dt.Rows[x][0].ToString();
                        bincont.item = dt.Rows[x][1].ToString();
                        bincont.descrip = dt.Rows[x][2].ToString();
                        bincont.um= dt.Rows[x][4].ToString();
                        bincont.cantidad =double.Parse(dt.Rows[x][5].ToString());
                        bincont.planta = sPta;
                        BinContLogica.guardar(bincont);                        
                    }
                }

               

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Cursor = Cursors.Arrow;
            }
            Cursor = Cursors.Arrow;



        }

        private void GenerarInvCiclicoPickline()
        {
            BinContLogica bincont = new BinContLogica();

            try
            {
                string sPta = "EMPN";
                string sLine = string.Empty;
                string sLineF = string.Empty;
                switch (sPta)
                {
                    case "EMPN":
                        sLine = "TPACKA";
                        sLineF = "MX1APAC";
                        break;
                    case "COL":
                        sLine = "CTNR";
                        sLineF = "MX1ACTNR";
                        break;
                    case "MON":
                        sLine = "TNR";
                        sLineF = "MX1ATNR";
                        break;
                    case "INKM":
                        sLine = "INKM";
                        break;
                    case "INKP":
                        sLine = "INKP";
                        break;
                    case "FUS":
                        sLine = "FUS";
                        sLineF = "MX1AFUSER";
                        break;
                }

                string sHora = getFileTime();
                string sArchivo = "Registered pickline " + sHora;
                if (sLine != "TPACKA")
                    sArchivo += sLine.ToLower();

                sArchivo = _lsPath + @"\" + sArchivo + ".csv";
                if (!File.Exists(sArchivo))
                {
                    Cursor = Cursors.Default;
                    return;
                }
                bincont.hora = sHora;                
                if (!BinContLogica.VerificarRegistros(bincont))
                {
                    return;
                }
                DataTable dtBinCont = BinContLogica.obtenerBinCont(bincont);
                DataTable dt = LoadFile(sArchivo);

                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    string sRouting = dt.Rows[x][2].ToString();//9
                    if (string.IsNullOrEmpty(sRouting))
                    {
                        dt.Rows[x].Delete();
                        x--;
                        continue;
                    }

                    if (!sRouting.StartsWith(sLine))
                    {
                        dt.Rows[x].Delete();
                        x--;
                        continue;
                    }

                    //sRouting = sRouting.Substring(7);
                    int iP = 0;
                    string sRouLine = sRouting.Substring(sRouting.Length - 2);
                    if (!int.TryParse(sRouLine.Substring(0, 1), out iP))
                        sRouLine = sRouLine.Substring(1);
                    sRouLine = sRouLine.PadLeft(2, '0');

                }
                
                if (dt.Rows.Count > 0)
                {
                    compararInvCiclico(dtBinCont,dt);
                }
                
                Cursor = Cursors.Arrow;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public void compararInvCiclico(DataTable dt_BinCont, DataTable dt_Pickline)
        {
            BinContLogica bincont = new BinContLogica();
            string bincodepick = "";
            string itempick = "";
            double SUMCant = 0;

            for (int x = 0; x < dt_BinCont.Rows.Count; x++)
            {
                SUMCant = 0;
                bincont.folio = dt_BinCont.Rows[x][0].ToString();
                bincont.fecha = DateTime.Parse(dt_BinCont.Rows[x][1].ToString());
                bincont.hora = dt_BinCont.Rows[x][2].ToString();
                bincont.planta = dt_BinCont.Rows[x][3].ToString();
                bincont.bincode = dt_BinCont.Rows[x][4].ToString();
                bincont.item= dt_BinCont.Rows[x][5].ToString();
                bincont.descrip= dt_BinCont.Rows[x][6].ToString();
                bincont.um= dt_BinCont.Rows[x][7].ToString();
                bincont.cantidad=double.Parse(dt_BinCont.Rows[x][8].ToString());

                for (int cont_pick=0; cont_pick < dt_Pickline.Rows.Count; cont_pick++)
                {
                    bincodepick = dt_Pickline.Rows[cont_pick][2].ToString();
                    itempick = dt_Pickline.Rows[cont_pick][3].ToString();                   

                    if (bincont.bincode.Equals(bincodepick) && bincont.item.Equals(itempick))
                    {
                         SUMCant+=double.Parse(dt_Pickline.Rows[cont_pick][4].ToString());
                    }

                }
               
                    BinContLogica.ActualizarBinCont(bincont, SUMCant, bincont.cantidad - SUMCant);

                

            }


        }

        private string getFileTime()
        {
            DateTime dtTime = DateTime.Now;
            int iHora = dtTime.Hour;
            if (iHora > 23)
                iHora -= 23;

            if (iHora >= 0 && iHora < 6)
                dtTime = dtTime.AddDays(-1);
            string sHrFile = string.Empty;
            if (iHora < 12)
                sHrFile = iHora.ToString() + "am";
            else
            {
                if (iHora > 23)
                {
                    if (iHora == 24)
                        iHora = 12;
                    else
                        iHora -= 24;
                    sHrFile = iHora.ToString() + "am";
                }
                else
                {
                    if (iHora > 12)
                        iHora -= 12;
                    sHrFile = iHora.ToString() + "pm";
                }
            }
            return sHrFile;
        }

        private DataTable LoadFile(string _asFile)
        {
            int iErr = 0;
            DataTable dt = new DataTable();
            try
            {
                using (StreamReader sr = new StreamReader(_asFile))
                {
                    string[] headers = sr.ReadLine().Split(',');
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }

                    while (!sr.EndOfStream)
                    {

                        List<string> result = SplitCSV(sr.ReadLine());
                        //string[] rows = sr.ReadLine().Split(',');
                        if (result.Count > 0)
                        {
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < headers.Length; i++)
                            {
                                //dr[i] = rows[i];
                                dr[i] = result[i];
                                iErr = i;
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string sErr = iErr.ToString() + " " + e.ToString();
                MessageBox.Show(sErr, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return dt;
        }

        public List<string> SplitCSV(string line)
        {
            List<string> result = new List<string>();

            if (string.IsNullOrEmpty(line))
            {
                //throw new ArgumentException();
                return result;
            }




            int index = 0;
            int start = 0;
            bool inQuote = false;
            StringBuilder val = new StringBuilder();

            // parse line
            foreach (char c in line)
            {
                switch (c)
                {
                    case '"':
                        inQuote = !inQuote;
                        break;

                    case ',':
                        if (!inQuote)
                        {
                            result.Add(line.Substring(start, index - start)
                                .Replace("\"", ""));

                            start = index + 1;
                        }

                        break;
                }

                index++;
            }

            if (start < index)
            {
                result.Add(line.Substring(start, index - start).Replace("\"", ""));
            }

            return result;
        }


    }

}
