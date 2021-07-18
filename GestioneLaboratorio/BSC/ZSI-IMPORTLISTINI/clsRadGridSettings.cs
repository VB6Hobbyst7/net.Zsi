
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Telerik.WinControls.UI;

namespace ToDoNotificheBSC
{


    public static class clsRadGridSettings
    {

        public static Logger log;

        public static bool SaveColumnsSettings(RadGridView rgv, string s, string name)
        {
            bool bOK = true;

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            string sf = Path.Combine(s, System.Environment.UserName + string.Format("_{0}.xml", name));


            rgv.SaveLayout(sf);

            if (File.Exists(sf))
            {
                try
                {
                    if (Directory.Exists(Path.Combine(ZSI_IMPORTLISTINI.Properties.Settings.Default.pathServerApp, "GridLayout")))
                    {
                        //string df = Path.Combine(Path.Combine(ZSI_IMPORTLISTINI.Properties.Settings.Default.pathServerApp, "GridLayout"), System.Environment.UserName + string.Format("_{0}.xml", name));
                        //File.Copy(sf, df, true);
                    }
                    else
                    {
                        Directory.CreateDirectory(Path.Combine(ZSI_IMPORTLISTINI.Properties.Settings.Default.pathServerApp, "GridLayout"));

                    }



                    string df = Path.Combine(Path.Combine(ZSI_IMPORTLISTINI.Properties.Settings.Default.pathServerApp, "GridLayout"), System.Environment.UserName + string.Format("_{0}.xml", name));
                    File.Copy(sf, df, true);
                }
                catch (Exception ex)
                {
                    bOK = false;
                }

            }

            return bOK;
        }

        public static void GetColumnsSettings(RadGridView rgv, string s, string name, ComboBox cmbI)
        {

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            string[] fi = Directory.GetFiles(s);

            cmbI.DataSource = fi;


            try
            {
                for (int i = 0; i < cmbI.Items.Count; i++)
                {
                    if (cmbI.Items[i].ToString().Contains(System.Environment.UserName + string.Format("_{0}.xml", name)))
                    {
                        cmbI.SelectedIndex = i;
                        string sf = Path.Combine(s, System.Environment.UserName + string.Format("_{0}.xml", name));

                        check_columns(s, name);


                        rgv.LoadLayout(cmbI.Items[i].ToString());
                        break;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delle righe ordine {0}", ex.Message));
            }

        }

        public static void GetColumnsSettings(RadGridView rgv, string s, string name)
        {

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            string[] fi = Directory.GetFiles(s);

            try
            {
                string sf = Path.Combine(s, System.Environment.UserName + string.Format("_{0}.xml", name));

                check_columns(s, name);

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delle righe ordine {0}", ex.Message));
            }

        }

        public static void check_columns(string s, string name)
        {
            try
            {
                string sXML = Path.Combine(s, System.Environment.UserName + string.Format("_{0}.xml", name));
                string tXML = Path.Combine(s, string.Format("template_{0}.xml", name));

                try
                {
                    if (Directory.Exists(Path.Combine(ZSI_IMPORTLISTINI.Properties.Settings.Default.pathServerApp, "GridLayout")))
                    {
                        string sf = Path.Combine(Path.Combine(ZSI_IMPORTLISTINI.Properties.Settings.Default.pathServerApp, "GridLayout"), System.Environment.UserName + string.Format("_{0}.xml", name));
                        File.Copy(sf, tXML, true);
                    }
                }
                catch (Exception ex)
                {

                }

                if (!File.Exists(sXML))
                {
                    if (File.Exists(tXML))
                    {
                        File.Copy(tXML, sXML, true);
                    }
                }
                else
                {
                    XmlDocument xmlDOC = new XmlDocument();
                    xmlDOC.Load(sXML);

                    if (File.Exists(tXML))
                    { 
                        XmlDocument tXmlDOC = new XmlDocument();
                        tXmlDOC.Load(tXML);

                        foreach (XmlElement te in tXmlDOC.GetElementsByTagName("Columns"))
                        {

                            if (te.ParentNode.Name == "MasterTemplate")
                            {
                                // impostazioni colonne
                                foreach (XmlNode tXMLNode in te.ChildNodes)
                                {
                                    string nSearch = tXMLNode.Attributes["Name"].InnerText;

                                    if (!xmlDOC.SelectNodes("RadGridView/MasterTemplate/Columns")[0].InnerXml.ToString().Contains(string.Format("Name=\"{0}\"", nSearch)))
                                    {

                                        //Import the last book node from doc2 into the original document.
                                        XmlNode newNode = xmlDOC.ImportNode(tXMLNode, true);
                                        xmlDOC.SelectNodes("RadGridView/MasterTemplate/Columns")[0].AppendChild(newNode);

                                    }

                                }

                            }

                        }
                    }


                    xmlDOC.Save(sXML);
                }
            }
            catch (Exception ex)
            {

                log.LogSomething(string.Format("ERRORE IN CARICAMENTO IMPOSTAZIONI COLONNE \r\n{0}", ex.Message));

            }

        }

        


    }
}
