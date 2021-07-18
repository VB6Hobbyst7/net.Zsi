using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.Office.Interop.Excel;

namespace ToDoNotificheBSC
{
    class ExportExcel
    {
        public static string excelFile = "";


        public static void creaModuloM12ITA(System.DateTime dt, string spedizioniere, string xlsmodel = "", string filename = "", int indexsheet = 0, int startrow = 1)
        {

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                if (System.IO.File.Exists(xlsmodel))
                {
                    // copia in locale
                    System.IO.File.Copy(xlsmodel, filename, true);
                    xlsmodel = filename;

                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(xlsmodel));
                }
                else
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());
                }

                if (indexsheet == 0)
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                }
                else
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets[indexsheet];
                }

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = dt;
                //oSheet.Cells[1, 2] = "Last Name";
                //oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[4, 2] = spedizioniere;

                // Run the macro, "First_Macro"
                oXL.Run("CaricaM12");
                
                ////Format A1:D1 as bold, vertical alignment = center.
                //oSheet.get_Range("A1", "D1").Font.Bold = true;
                //oSheet.get_Range("A1", "D1").VerticalAlignment =
                //    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                //// Create an array to multiple values at once.
                //string[,] saNames = new string[5, 2];

                //saNames[0, 0] = "John";
                //saNames[0, 1] = "Smith";
                //saNames[1, 0] = "Tom";

                //saNames[4, 1] = "Johnson";

                ////Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "B6").Value2 = saNames;

                ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //oRng = oSheet.get_Range("C2", "C6");
                //oRng.Formula = "=A2 & \" \" & B2";

                ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                //oRng = oSheet.get_Range("D2", "D6");
                //oRng.Formula = "=RAND()*100000";
                //oRng.NumberFormat = "$0.00";

                ////AutoFit columns A:D.
                //oRng = oSheet.get_Range("A1", "D1");
                //oRng.EntireColumn.AutoFit();


                oXL.UserControl = false;
                oWB.Save();

                oXL.Visible = true;


                //oWB.Close();
                //oXL.Quit();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }

        }

        public static void creaModuloM10EST(System.DateTime dt, string xlsmodel = "", string filename = "", int indexsheet = 0, int startrow = 1)
        {

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                if (System.IO.File.Exists(xlsmodel))
                {
                    // copia in locale
                    System.IO.File.Copy(xlsmodel, filename, true);
                    xlsmodel = filename;

                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(xlsmodel));
                }
                else
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());
                }

                if (indexsheet == 0)
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                }
                else
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets[indexsheet];
                }

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = dt;
                //oSheet.Cells[1, 2] = "Last Name";
                //oSheet.Cells[1, 3] = "Full Name";
                //oSheet.Cells[4, 2] = spedizioniere;

                // Run the macro, "First_Macro"
                oXL.Run("CaricaM10");

                ////Format A1:D1 as bold, vertical alignment = center.
                //oSheet.get_Range("A1", "D1").Font.Bold = true;
                //oSheet.get_Range("A1", "D1").VerticalAlignment =
                //    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                //// Create an array to multiple values at once.
                //string[,] saNames = new string[5, 2];

                //saNames[0, 0] = "John";
                //saNames[0, 1] = "Smith";
                //saNames[1, 0] = "Tom";

                //saNames[4, 1] = "Johnson";

                ////Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "B6").Value2 = saNames;

                ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //oRng = oSheet.get_Range("C2", "C6");
                //oRng.Formula = "=A2 & \" \" & B2";

                ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                //oRng = oSheet.get_Range("D2", "D6");
                //oRng.Formula = "=RAND()*100000";
                //oRng.NumberFormat = "$0.00";

                ////AutoFit columns A:D.
                //oRng = oSheet.get_Range("A1", "D1");
                //oRng.EntireColumn.AutoFit();


                oXL.UserControl = false;
                oWB.Save();

                oXL.Visible = true;


                //oWB.Close();
                //oXL.Quit();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }

        }

        public static void creaModuloM04(System.DateTime dt, string xlsmodel = "", string filename = "", int indexsheet = 0, int startrow = 1)
        {

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                if (System.IO.File.Exists(xlsmodel))
                {
                    // copia in locale
                    System.IO.File.Copy(xlsmodel, filename, true);
                    xlsmodel = filename;

                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(xlsmodel));
                }
                else
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());
                }

                if (indexsheet == 0)
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                }
                else
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets[indexsheet];
                }

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = dt;

                if (xlsmodel.Contains("M04"))
                {
                    oXL.Run("CaricaM04");
                }

                if (xlsmodel.Contains("M10"))
                {
                    oXL.Run("CaricaM10");
                }


                oXL.UserControl = false;
                oWB.Save();

                oXL.Visible = true;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }

        }

        public static void creaModuloCOA(System.DateTime dt, string xlsmodel = "", string filename = "", int indexsheet = 0, int startrow = 1)
        {

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                if (System.IO.File.Exists(xlsmodel))
                {
                    // copia in locale
                    System.IO.File.Copy(xlsmodel, filename, true);
                    xlsmodel = filename;

                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(xlsmodel));
                }
                else
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());
                }

                if (indexsheet == 0)
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                }
                else
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets[indexsheet];
                }

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = dt;

                if (xlsmodel.Contains("COA"))
                {
                    oXL.Run("CaricaPosizionamentiCOA");
                }

                if (xlsmodel.Contains("M10"))
                {
                    oXL.Run("CaricaM10");
                }


                oXL.UserControl = false;
                oWB.Save();

                oXL.Visible = true;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }

        }


        public static void creaModulo(System.Data.DataTable dt, string xlsmodel = "", string filename = "", int indexsheet = 0, int startrow = 1)
        {

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                if (System.IO.File.Exists(xlsmodel))
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(xlsmodel));
                }
                else
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add());
                }

                if (indexsheet == 0)
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                }
                else
                {
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets[indexsheet];
                }

                ////Add table headers going cell by cell.
                //oSheet.Cells[1, 1] = "First Name";
                //oSheet.Cells[1, 2] = "Last Name";
                //oSheet.Cells[1, 3] = "Full Name";
                //oSheet.Cells[1, 4] = "Salary";

                ////Format A1:D1 as bold, vertical alignment = center.
                //oSheet.get_Range("A1", "D1").Font.Bold = true;
                //oSheet.get_Range("A1", "D1").VerticalAlignment =
                //    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                //// Create an array to multiple values at once.
                //string[,] saNames = new string[5, 2];

                //saNames[0, 0] = "John";
                //saNames[0, 1] = "Smith";
                //saNames[1, 0] = "Tom";

                //saNames[4, 1] = "Johnson";

                ////Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "B6").Value2 = saNames;

                ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //oRng = oSheet.get_Range("C2", "C6");
                //oRng.Formula = "=A2 & \" \" & B2";

                ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                //oRng = oSheet.get_Range("D2", "D6");
                //oRng.Formula = "=RAND()*100000";
                //oRng.NumberFormat = "$0.00";

                ////AutoFit columns A:D.
                //oRng = oSheet.get_Range("A1", "D1");
                //oRng.EntireColumn.AutoFit();


                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        oSheet.Cells[startrow + j, i+1] = dt.Rows[j][i].ToString();
                    }
                    //oSheet.Cells[1, 2] = "Last Name";
                    //oSheet.Cells[1, 3] = "Full Name";
                    //oSheet.Cells[1, 4] = "Salary";



                }

                if (System.IO.File.Exists(filename))
                    System.IO.File.Delete(filename);

                //oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
                oXL.Quit();
                
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }

        }



        public static void exportToPDF(string infile, string filename)
        {

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                if (System.IO.File.Exists(filename))
                    System.IO.File.Delete(filename);

                //Get a new workbook.
                if (System.IO.File.Exists(infile))
                {
                    oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Open(infile));
                    //oXL.Visible = false;
                    oXL.UserControl = false;
                    oWB.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filename);

                    oWB.Close(XlSaveAction.xlDoNotSaveChanges);
                }



                oXL.Quit();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return;
            }

        }




    }
}
