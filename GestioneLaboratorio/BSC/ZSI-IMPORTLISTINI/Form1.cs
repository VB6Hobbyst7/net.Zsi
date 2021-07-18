using ExcelImportExportLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls.Data;
using ToDoNotificheBSC;

namespace ZSI_IMPORTLISTINI
{



    public partial class Form1 : Form
    {

        DataSet d;

        Logger log;

        public class XLManage
        {

            public static DataSet ImportExcelXLS(string FileName, bool hasHeaders)
            {
                string HDR = hasHeaders ? "Yes" : "No";
                string strConn;
                if (FileName.Substring(FileName.LastIndexOf('.')).ToLower() == ".xlsx")
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                else
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";

                DataSet output = new DataSet();

                using (OleDbConnection conn = new OleDbConnection(strConn))
                {
                    conn.Open();

                    System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(
                        OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    foreach (DataRow schemaRow in schemaTable.Rows)
                    {
                        string sheet = schemaRow["TABLE_NAME"].ToString();

                        //if (!sheet.EndsWith("_") && !sheet.Contains("Print"))
                        if (sheet.ToLower().Contains(Properties.Settings.Default.FoglioListino))
                        {
                            try
                            {
                                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                                cmd.CommandType = CommandType.Text;

                                System.Data.DataTable outputTable = new System.Data.DataTable(sheet);
                                output.Tables.Add(outputTable);
                                new OleDbDataAdapter(cmd).Fill(outputTable);
                            }
                            catch (OleDbException ex)
                            {
                                if ((ex.ErrorCode != -2147217865) )
                                {
                                    throw new Exception(ex.Message + string.Format("Sheet:{0}.File:F{1}", sheet, FileName), ex);
                                }
                            }
                        }
                    }
                }
                return output;
            }

        }


        Dictionary<string, string> colFields = new Dictionary<string, string>();


        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void btnCheckFile_Click(object sender, EventArgs e)
        {

        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog.FileName = "";

            openFileDialog.DefaultExt = ".xls|.xlsx";
            openFileDialog.Title = "Cerca file Excel";
            openFileDialog.Multiselect = false;

            openFileDialog.ShowDialog();

            cmbFoglioDiLavoro.Items.Clear();

            if (openFileDialog.FileName != "")
            {
                textBoxInputFilePath.Text = openFileDialog.FileName;
                //Uri u = new Uri(openFileDialog.FileName);
                //webBrowser1.Url = u;




                d = XLManage.ImportExcelXLS(openFileDialog.FileName, true);

                if (d.Tables.Count == 0)
                {
                    MessageBox.Show(string.Format("Il programma cerca un foglio di lavoro del file excel selezionato che contenga nel nome la parola '{0}' \r\n verificare il nome del foglio di lavoro contenente i listini del file selezionato", Properties.Settings.Default.FoglioListino), "Apertura foglio di lavoro", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }


                foreach (DataTable t in d.Tables)
                {
                    cmbFoglioDiLavoro.Items.Add(t.TableName);
                }

                if (cmbFoglioDiLavoro.Items.Count > 0)
                    selezionaFoglioDiLavoro(cmbFoglioDiLavoro.Items[0].ToString());



                btnPubblica.Enabled = true;

            }

        }

        void selezionaFoglioDiLavoro(string f)
        {
            dGExcel.DataSource = null;

            dGExcel.DataSource = d.Tables[f];
            dGExcel.EnableHeadersVisualStyles = false;
            dGExcel.Columns[2].Frozen = true;

            // nascondo le colonne inutili quelle senza lettara alfabeto
            for (int c = 0; c < dGExcel.Columns.Count; c++)
            {
                if (dGExcel.Columns[c].Name.Length > 1)
                {
                    dGExcel.Columns[c].Visible = false;
                }
            }



        }



        private void Form1_Load(object sender, EventArgs e)
        {
            log = new Logger();
            log.Setup();
            log.LogSomething("Start applicazione");



            foreach (string p in Properties.Settings.Default.ColumnFields)
            {
                string[] a = p.Split('|');

                if (a[1] != "")
                    colFields.Add(a[0], a[1]);

            }
        }

        bool esistearticolo(string codarticolo, string col, double valorelistino)
        {
            bool bEsiste = false;

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                try
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand(string.Format(Properties.Settings.Default.MetodoVerifica, codarticolo, col, valorelistino), cn))
                    {
                        bEsiste = (cmd.ExecuteScalar().ToString() == "1");
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return bEsiste;


        }

        private void btnPubblica_Click(object sender, EventArgs e)
        {
            string codart = "";
            double valorelistino = 0;
            double VAL_COL_A = 0;
            double VAL_COL_B = 0;
            double VAL_COL_C = 0;
            double VAL_COL_D = 0;
            double VAL_COL_E = 0;
            double VAL_COL_F = 0;
            double VAL_COL_G = 0;
            double VAL_COL_H = 0;
            double VAL_COL_I = 0;
            double VAL_COL_J = 0;
            double VAL_COL_K = 0;
            double VAL_COL_L = 0;
            double VAL_COL_M = 0;
            double VAL_COL_N = 0;
            double VAL_COL_O = 0;
            double VAL_COL_P = 0;
            double VAL_COL_Q = 0;
            double VAL_COL_R = 0;
            double VAL_COL_S = 0;
            double VAL_COL_T = 0;

            foreach (DataGridViewRow r in dGExcel.Rows)
            {
                Application.DoEvents();


                if ((r.Cells[colFields["MACROARTICOLO"].ToString()].Value.ToString() != "") && (r.HeaderCell.Value.ToString() == "Da Aggiornare"))
                {

                    codart = r.Cells[colFields["MACROARTICOLO"].ToString()].Value.ToString();
                    lblProgress.Text = string.Format("Analisi riga {0}, macroarticolo {1}", r.Index, codart);

                    log.LogSomething(string.Format("Analisi riga {0}, macroarticolo {1}", r.Index, codart));

                    double.TryParse(r.Cells[colFields["VAL_COL_D"].ToString()].Value.ToString(), out valorelistino);

                    if (valorelistino == 0)
                    { double.TryParse(r.Cells[colFields["VAL_COL_L"].ToString()].Value.ToString(), out valorelistino); }


                    double.TryParse(r.Cells[colFields["VAL_COL_A"].ToString()].Value.ToString(), out VAL_COL_A);
                    double.TryParse(r.Cells[colFields["VAL_COL_B"].ToString()].Value.ToString(), out VAL_COL_B);
                    double.TryParse(r.Cells[colFields["VAL_COL_C"].ToString()].Value.ToString(), out VAL_COL_C);
                    double.TryParse(r.Cells[colFields["VAL_COL_D"].ToString()].Value.ToString(), out VAL_COL_D);
                    double.TryParse(r.Cells[colFields["VAL_COL_E"].ToString()].Value.ToString(), out VAL_COL_E);
                    double.TryParse(r.Cells[colFields["VAL_COL_F"].ToString()].Value.ToString(), out VAL_COL_F);
                    double.TryParse(r.Cells[colFields["VAL_COL_G"].ToString()].Value.ToString(), out VAL_COL_G);
                    double.TryParse(r.Cells[colFields["VAL_COL_H"].ToString()].Value.ToString(), out VAL_COL_H);
                    double.TryParse(r.Cells[colFields["VAL_COL_I"].ToString()].Value.ToString(), out VAL_COL_I);
                    //                    double.TryParse(r.Cells[colFields["VAL_COL_J"].ToString()].Value.ToString(), out VAL_COL_J);
                    //                    double.TryParse(r.Cells[colFields["VAL_COL_K"].ToString()].Value.ToString(), out VAL_COL_K);
                    double.TryParse(r.Cells[colFields["VAL_COL_L"].ToString()].Value.ToString(), out VAL_COL_L);
                    double.TryParse(r.Cells[colFields["VAL_COL_M"].ToString()].Value.ToString(), out VAL_COL_M);
                    double.TryParse(r.Cells[colFields["VAL_COL_N"].ToString()].Value.ToString(), out VAL_COL_N);
                    double.TryParse(r.Cells[colFields["VAL_COL_O"].ToString()].Value.ToString(), out VAL_COL_O);
                    double.TryParse(r.Cells[colFields["VAL_COL_P"].ToString()].Value.ToString(), out VAL_COL_P);
                    double.TryParse(r.Cells[colFields["VAL_COL_Q"].ToString()].Value.ToString(), out VAL_COL_Q);
                    double.TryParse(r.Cells[colFields["VAL_COL_R"].ToString()].Value.ToString(), out VAL_COL_R);
                    double.TryParse(r.Cells[colFields["VAL_COL_S"].ToString()].Value.ToString(), out VAL_COL_S);
                    double.TryParse(r.Cells[colFields["VAL_COL_T"].ToString()].Value.ToString(), out VAL_COL_T);


                    if (valorelistino == 0)
                    {
                        r.HeaderCell.Value = "Listino a 0";
                        log.LogSomething(string.Format("Analisi riga {0}, macroarticolo {1} - LISTINO = 0!!!", r.Index, codart));
                    }
                    else
                    {
                        using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                        {
                            try
                            {
                                cn.Open();

                                using (SqlCommand cmd = new SqlCommand(string.Format(Properties.Settings.Default.MetodoAggiornaCMD), cn))
                                {
                                    cmd.CommandType = CommandType.StoredProcedure;
                                    cmd.Parameters.AddWithValue("@MACROARTICOLO", codart);
                                    cmd.Parameters.AddWithValue("@VAL_COL_A", Math.Round(VAL_COL_A, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_B", Math.Round(VAL_COL_B, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_C", Math.Round(VAL_COL_C, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_D", Math.Round(VAL_COL_D, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_E", Math.Round(VAL_COL_E, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_F", Math.Round(VAL_COL_F, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_G", Math.Round(VAL_COL_G, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_H", Math.Round(VAL_COL_H, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_I", Math.Round(VAL_COL_I, 3, MidpointRounding.AwayFromZero));
                                    //cmd.Parameters.AddWithValue("@VAL_COL_J", Math.Round(VAL_COL_J, 3, MidpointRounding.AwayFromZero));
                                    //cmd.Parameters.AddWithValue("@VAL_COL_K", Math.Round(VAL_COL_K, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_L", Math.Round(VAL_COL_L, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_M", Math.Round(VAL_COL_M, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_N", Math.Round(VAL_COL_N, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_O", Math.Round(VAL_COL_O, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_P", Math.Round(VAL_COL_P, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_Q", Math.Round(VAL_COL_Q, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_R", Math.Round(VAL_COL_R, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_S", Math.Round(VAL_COL_S, 3, MidpointRounding.AwayFromZero));
                                    cmd.Parameters.AddWithValue("@VAL_COL_T", Math.Round(VAL_COL_T, 3, MidpointRounding.AwayFromZero));


                                    cmd.ExecuteNonQuery();
                                    r.HeaderCell.Value = "OK";
                                    r.HeaderCell.Style.BackColor = Color.Lime;

                                    log.LogSomething(string.Format("AGGIORNATO riga {0}, macroarticolo {1}", r.Index, codart));
                                }
                            }
                            catch (Exception ex)
                            {
                                r.HeaderCell.Value = ex.Message;
                                r.HeaderCell.Style.BackColor = Color.Red;

                                log.LogSomething(string.Format("ERRORE riga {0}, macroarticolo {1} - {2}", r.Index, codart, ex.Message));
                            }
                            
                        }
                    }
                    //textBoxOutput.Text += string.Format("\r\n{0} {1} {2} {3}", r.Index, k, colFields[k].ToString(), r.Cells[colFields[k].ToString()].Value.ToString());

                }
            }
            btnPubblica.Enabled = false;
            MessageBox.Show("Procedura conclusa!");
        }


        private void btnVerifica_Click(object sender, EventArgs e)
        {
            int nrDaAggiornare = 0;
            int nrListini0 = 0;
            int nrRighe = 0;

            dGExcel.RowHeadersWidth = 130;

            string codart = "";
            string col = "";
            double valorelistino = 0;

            progressBar1.Visible = true;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = dGExcel.Rows.Count;
            progressBar1.Step = 1;
            progressBar1.Value = 0;

            foreach (DataGridViewRow r in dGExcel.Rows)
            {
                Application.DoEvents();
                //foreach (string k in colFields.Keys)
                //{

                lblProgress.Text = string.Format("Analisi riga {0}", r.Index);

                progressBar1.Increment(1);
                if (r.Cells[colFields["MACROARTICOLO"].ToString()].Value.ToString() != "")
                {
                    nrRighe++;

                    codart = r.Cells[colFields["MACROARTICOLO"].ToString()].Value.ToString();
                    col = "VAL_COL_D";
                    double.TryParse(r.Cells[colFields[col].ToString()].Value.ToString(), out valorelistino);

                    if (valorelistino == 0)
                    {
                        col = "VAL_COL_L";
                        double.TryParse(r.Cells[colFields[col].ToString()].Value.ToString(), out valorelistino);
                    }


                    if (valorelistino == 0)
                    {
                        r.HeaderCell.Value = "Listino a 0";
                        r.HeaderCell.Style.ForeColor = Color.Red;
                        nrListini0++;
                    }
                    else
                    {

                        if (esistearticolo(codart, col, valorelistino) == false)
                        {
                            r.HeaderCell.Value = "Da Aggiornare";
                            r.HeaderCell.Style.ForeColor = Color.Orange;
                            nrDaAggiornare++;
                        }
                        else
                        {
                            r.HeaderCell.Value = "";
                            //r.HeaderCell.Style.BackColor = Color.Black;
                        }
                    }
                    //textBoxOutput.Text += string.Format("\r\n{0} {1} {2} {3}", r.Index, k, colFields[k].ToString(), r.Cells[colFields[k].ToString()].Value.ToString());

                }
                else
                {
                    r.Height = 0;
                }
                //}            

            }
            progressBar1.Visible = false;
            lblProgress.Text = "";
            textBoxOutput.Text = string.Format("Righe: {0}\r\nRighe da aggiornare: {1}\r\nRighe listini 0: {2}", nrRighe, nrDaAggiornare, nrListini0);

        }

        private void btnStorico_Click(object sender, EventArgs e)
        {
            if (radGridViewPP.Visible == false)
            {
                radGridViewPP.DataSource = caricastorico();
                radGridViewPP.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.DisplayedDataCells);
                radGridViewPP.Visible = true;
                radGridViewPP.BringToFront();
            }
            else
            {
                radGridViewPP.Visible = false;
            }

        }

        DataTable caricastorico()
        {
            DataTable dt = new DataTable();

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                try
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand(string.Format(Properties.Settings.Default.MetodoStoricoCMD), cn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return dt;


        }

        private void cmbFoglioDiLavoro_SelectedIndexChanged(object sender, EventArgs e)
        {
            selezionaFoglioDiLavoro(cmbFoglioDiLavoro.SelectedItem.ToString());
        }

        private void btnCaricaPrezzi_Click(object sender, EventArgs e)
        {
            radGridViewPP.DataSource = CaricaProvvigioni();

            radGridViewPP.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.DisplayedDataCells);

            clsRadGridSettings.GetColumnsSettings(radGridViewPP, Application.StartupPath, "GestionePrezzi");

        }

        DataTable CaricaProvvigioni()
        {

            DataTable dt = new DataTable();

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                try
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand(string.Format(Properties.Settings.Default.MetodoPrezziparticolari), cn))
                    {
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                    }
                }
                catch (Exception ex)
                {

                }
            }

            return dt;
        }

        private void radGridViewPP_CellEndEdit(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {

                if (chkEuro.Checked)
                {
                    radGridViewPP.Rows[e.RowIndex].Cells["PREZZO_MAGGEURO"].Value = radGridViewPP.Rows[e.RowIndex].Cells["PREZZO_MAGG"].Value;
                }


                radGridViewPP.Rows[e.RowIndex].Cells[0].Value = true;

            }
        }

        private void btnAggiornaPrezzi_Click(object sender, EventArgs e)
        {

            clsRadGridSettings.SaveColumnsSettings(radGridViewPP, Application.StartupPath, "GestionePrezzi");

            if (radGridViewPP.ChildRows.Count > 0)
            {
                progressBar2.Visible = true;
                progressBar2.Minimum = 0;
                progressBar2.Maximum = radGridViewPP.ChildRows.Count;
                progressBar2.Step = 1;

                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    try
                    {
                        cn.Open();


                        for (int i = 0; i < radGridViewPP.ChildRows.Count; i++)
                        {
                            if (radGridViewPP.ChildRows[i].Cells[0].Value != null)
                            {
                                if (radGridViewPP.ChildRows[i].Cells[0].Value.ToString() == "True")
                                {
                                    string xSQL = radGridViewPP.ChildRows[i].Cells["SQLUPDATE"].Value.ToString();
                                    string xSQL_TESTATE = radGridViewPP.ChildRows[i].Cells["SQLUPDATE_TESTATE"].Value.ToString();

                                    xSQL = xSQL.Replace("#PREZZO_MAGG#", radGridViewPP.ChildRows[i].Cells["PREZZO_MAGG"].Value.ToString().Replace(",", "."));
                                    xSQL = xSQL.Replace("#PREZZO_MAGGEURO#", radGridViewPP.ChildRows[i].Cells["PREZZO_MAGGEURO"].Value.ToString().Replace(",", "."));

                                    //this.Text = xSQL;

                                    using (SqlCommand cmd = new SqlCommand(string.Format(xSQL), cn))
                                    {
                                        cmd.ExecuteNonQuery();
                                        //progressBar2.Increment(1);
                                        radGridViewPP.ChildRows[i].Cells[0].Style.CustomizeFill = true;
                                        radGridViewPP.ChildRows[i].Cells[0].Style.BackColor = Color.Lime;
                                        radGridViewPP.ChildRows[i].Cells[0].Style.DrawFill = true;
                                        radGridViewPP.ChildRows[i].Cells[0].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                    }


                                    using (SqlCommand cmd = new SqlCommand(string.Format(xSQL_TESTATE), cn))
                                    {
                                        if (radGridViewPP.ChildRows[i].Cells["T"].Value.ToString() == "T")
                                        {
                                            cmd.Parameters.AddWithValue("@0", radGridViewPP.ChildRows[i].Cells["INIZIOVALIDITA"].Value);
                                            cmd.Parameters.AddWithValue("@1", radGridViewPP.ChildRows[i].Cells["FINEVALIDITA"].Value);
                                            cmd.ExecuteNonQuery();

                                            radGridViewPP.ChildRows[i].Cells["T"].Style.CustomizeFill = true;
                                            radGridViewPP.ChildRows[i].Cells["T"].Style.BackColor = Color.Lime;
                                            radGridViewPP.ChildRows[i].Cells["T"].Style.DrawFill = true;
                                            radGridViewPP.ChildRows[i].Cells["T"].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                        }
                                    }
                                }
                            }

                            progressBar2.Increment(1);

                        }

                        if (radGridViewPP.FilterDescriptors.Count == 1)
                        {
                            radGridViewPP.FilterDescriptors.Remove("Modificato");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

                progressBar2.Visible = false;
            }

        }

        private void btnVerificaPrezzi_Click(object sender, EventArgs e)
        {
            this.radGridViewPP.MasterTemplate.EnableFiltering = true;

            FilterDescriptor filter = new FilterDescriptor();
            filter.PropertyName = "Modificato";
            filter.Operator = FilterOperator.IsEqualTo;
            filter.Value = true;
            filter.IsFilterEditor = true;

            if (radGridViewPP.FilterDescriptors.Count == 1)
            {
                radGridViewPP.FilterDescriptors.Remove(filter.PropertyName);
            }
            else
            {
                radGridViewPP.FilterDescriptors.Add(filter);
            }

        }
    }
}
