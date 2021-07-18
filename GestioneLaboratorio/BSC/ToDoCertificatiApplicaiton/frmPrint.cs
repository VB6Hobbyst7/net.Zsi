using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;


namespace crPrint
{
    
    public partial class frmPrint : Form
    {

        public frmPrint()
        {
            InitializeComponent();
        }

        public SqlConnection cn = new SqlConnection();
        public string user = "";
        public string pwd = "";
        public string reportFile = "";
        public string macchina = "";
        public bool anteprima;
        public int nCopie = 1;
        public string stampante = "";
        public int nrGiorniOrizzonte = 0;
        public double NumeroBolla = 0;
        public int AnnoBolla = 0;
        public double IdTesta = 0;
        public double IdRiga = 0;
        public string CODCONTO = "";

        


        //private void frmPrint_Load(object sender, EventArgs e)
        //{
        //    try
        //    {



        //        CrystalDecisions.Shared.ConnectionInfo connectionInfo = new
              
        //        CrystalDecisions.Shared.ConnectionInfo();
        //        //connectionInfo.IntegratedSecurity = true;
        //        connectionInfo.DatabaseName = cn.Database;// sq RPCSShipping.Program.Database;
        //        connectionInfo.ServerName = cn.DataSource; //ser  RPCSShipping.Program.Server;
        //        connectionInfo.UserID = user; //d RPCSShipping.Program.UserID;
        //        connectionInfo.Password = pwd; // RPCSShipping.Program.Password;

        //        MessageBox.Show(Application.StartupPath + "\\MFTOTALE.rpt");

        //        CrystalDecisions.CrystalReports.Engine.ReportDocument reportDocument1 = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        //        reportDocument1.Load(Application.StartupPath + "\\MFTOTALE.rpt");


        //        CrystalDecisions.CrystalReports.Engine.Tables tables = reportDocument1.Database.Tables;
        //        foreach (CrystalDecisions.CrystalReports.Engine.Table table in tables)
        //        {
        //            CrystalDecisions.Shared.TableLogOnInfo tableLogonInfo = table.LogOnInfo;
        //            tableLogonInfo.ConnectionInfo = connectionInfo;
        //            table.ApplyLogOnInfo(tableLogonInfo);
        //        }

        //        //reportDocument1.SetParameterValue("@DITTA", "''");

        //        crystalReportViewer1.ReportSource = reportDocument1;
        //        //crystalReportViewer1.PrintReport();
        //        //reportDocument1.PrintOptions.PrinterName = "";
        //        //reportDocument1.PrintToPrinter(1, true, 1, 1);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.StackTrace.ToString());
        //    }
        //}



    
        private void frmPrint_Load(object sender, EventArgs e)
        {
            string strMessage = "";

            if (reportFile == "")
            {
                MessageBox.Show("Non è stato imostato un report da stampare");
                return;
            }
            else
            {
                if (System.IO.File.Exists(reportFile) == false)
                {
                    MessageBox.Show("Il report impostato: " + reportFile + " non esiste");
                    return;
                }
            
            }

            
            try
            {
		        ParameterFieldDefinitions crPF_Defs;
		        ParameterFieldDefinition crPF_Def;
		        ParameterDiscreteValue crP_DiscreteVal = new ParameterDiscreteValue();
		        ParameterValues crP_Values = new ParameterValues();



                CrystalDecisions.Shared.ConnectionInfo connectionInfo = new
              
                CrystalDecisions.Shared.ConnectionInfo();
                //connectionInfo.IntegratedSecurity = true;
                connectionInfo.DatabaseName = cn.Database;// sq RPCSShipping.Program.Database;
                connectionInfo.ServerName = cn.DataSource; //ser  RPCSShipping.Program.Server;
                connectionInfo.UserID = user; //d RPCSShipping.Program.UserID;
                connectionInfo.Password = pwd; // RPCSShipping.Program.Password;


                CrystalDecisions.CrystalReports.Engine.ReportDocument reportDocument1 = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                reportDocument1.Load(reportFile);


                CrystalDecisions.CrystalReports.Engine.Tables tables = reportDocument1.Database.Tables;
                foreach (CrystalDecisions.CrystalReports.Engine.Table table in tables)
                {
                    CrystalDecisions.Shared.TableLogOnInfo tableLogonInfo = table.LogOnInfo;
                    tableLogonInfo.ConnectionInfo = connectionInfo;
                    table.ApplyLogOnInfo(tableLogonInfo);
                }

                //reportDocument1.SetParameterValue("@DITTA", "''");

                // stampa
                try
                {

                    //passaggio parametro
                    if (reportDocument1.ParameterFields.Count > 0)
                    {

                        crPF_Defs = reportDocument1.DataDefinition.ParameterFields;

                        if ((reportFile.Contains("CalSpedizioni.rpt") == true))
                        {
                            
                            //reportDocument1.Refresh();
                            crPF_Def = crPF_Defs["CODCONTO"];
                            crP_Values = crPF_Def.CurrentValues;
                            // attribuisco il valore del parametro
                            crP_DiscreteVal.Value = CODCONTO;
                            crP_Values.Add(crP_DiscreteVal);

                            crPF_Def.ApplyCurrentValues(crP_Values);

                            crPF_Def = crPF_Defs["Pm-?CODCONTO"];
                            crP_Values = crPF_Def.CurrentValues;
                            // attribuisco il valore del parametro
                            crP_DiscreteVal.Value = CODCONTO;
                            crP_Values.Add(crP_DiscreteVal);

                            crPF_Def.ApplyCurrentValues(crP_Values);
                        }

                    }

                    //cr.PrintToPrinter(nrcopie, collate(t/f),fromPage (0=all),toPage (0=all));
                    //cr.PrintToPrinter(nCopie, false, 0, 0);
                    strMessage = strMessage + string.Format("{0: MM/dd/yyyy HH:mm:ss ms tt}", DateTime.Now) + "\n";
                    strMessage = strMessage + "Stampa Effettuata \n";
                }
                catch (Exception ex)
                {
                    strMessage = strMessage + ex.Message;
                    strMessage = strMessage + "\n *** Stampa NON Effettuata  **** \n";

                }

                crystalReportViewer1.ReportSource = reportDocument1;
                //crystalReportViewer1.PrintReport();

                if (stampante != "")
                {
                    try
                    {
                        reportDocument1.PrintOptions.PrinterName = stampante;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Errore nell'impostazione della stampante: " + stampante + "\n\r" + ex.Message);
                    }
                }

                if (anteprima == false)
                {
                    reportDocument1.PrintToPrinter(nCopie, true, 0, 0);
                    reportDocument1.Dispose();
                    MessageBox.Show("Stampa effettuata sulla stampante: " + stampante + ".");
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace.ToString());
            }
        }
    }




}
