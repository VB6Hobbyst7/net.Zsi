using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Xml.Linq;

namespace ToDoNotificheBSC
{
    

    public static class clsLogin
    {
        public static string CurrentUser;
        public static string CurrentPWD;
        public static string MetodoConnectionStringUSER;

        public static string loginXMLFile = Path.Combine(Application.StartupPath, string.Format("loginXMLFile.xml"));

        public static bool loadCredenziali()
        {

            bool bOK = false;

            if (File.Exists(loginXMLFile))
            {

                try
                {
                    XmlDocument d = new XmlDocument();
                    d.Load(loginXMLFile);

                    XmlNode nCurrentUser = d.SelectSingleNode("root/CurrentUser");
                    XmlNode nCurrentPWD = d.SelectSingleNode("root/CurrentPWD");
                    XmlNode nMetodoConnectionStringUSER = d.SelectSingleNode("root/MetodoConnectionStringUSER");


                    using (SqlConnection cnUser = new SqlConnection(string.Format(nMetodoConnectionStringUSER.InnerText, nCurrentUser.InnerText, nCurrentPWD.InnerText)))
                    {
                        try
                        {
                            cnUser.Open();

                            //Properties.Settings.Default.MetodoConnectionStringUSER = string.Format(nMetodoConnectionStringUSER.Value, nCurrentUser.Value, nCurrentPWD.Value);
                            //Properties.Settings.Default.CurrentUser = nCurrentUser.Value;
                            //Properties.Settings.Default.Save();

                            CurrentUser = nCurrentUser.InnerText;
                            CurrentPWD = nCurrentPWD.InnerText;
                            MetodoConnectionStringUSER = string.Format(nMetodoConnectionStringUSER.InnerText, nCurrentUser.InnerText, nCurrentPWD.InnerText);

                            Global.Ditta = cnUser.Database;
                            Global.UtenteMetodo = CurrentUser;
                            Global.PwdMetodo = CurrentPWD;


                            bOK = true;


                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Accesso non riuscito!\r\n" + ex.Message);

                        }

                        return bOK;
                    }

                }
                catch (Exception ex)
                {
                    return bOK;
                }


            }

            return bOK;

        }

        public static bool saveCredenziali(string user, string pwd, string conn)
        {

            bool bOK = false;

            if (File.Exists(loginXMLFile))
            {
                File.Delete(loginXMLFile);
            }
            XDocument d;

            try
            {

                using (SqlConnection cnUser = new SqlConnection(string.Format(conn, user, pwd)))
                {
                    try
                    {
                        cnUser.Open();

                        //Properties.Settings.Default.MetodoConnectionStringUSER = string.Format(conn, user, pwd);
                        //Properties.Settings.Default.CurrentUser = user;
                        //Properties.Settings.Default.Save();
                        CurrentUser = user;
                        CurrentPWD = pwd;
                        MetodoConnectionStringUSER = string.Format(conn, user, pwd);

                        d = new XDocument();
                        XElement eRoot = new XElement("root");
                        XElement eCurrentUser = new XElement("CurrentUser", user);
                        XElement eCurrentPWD = new XElement("CurrentPWD", pwd);
                        XElement eMetodoConnectionStringUSER = new XElement("MetodoConnectionStringUSER", conn);
                        eRoot.Add(eCurrentUser);
                        eRoot.Add(eCurrentPWD);
                        eRoot.Add(eMetodoConnectionStringUSER);

                        d.Add(eRoot);
                        //d.Add(eCurrentUser);
                        //d.Add(eCurrentPWD);
                        //d.Add(eMetodoConnectionStringUSER);


                        d.Save(loginXMLFile);

                        Global.Ditta = cnUser.Database;
                        Global.UtenteMetodo = user;
                        Global.PwdMetodo = pwd;

                        bOK = true;


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Accesso non riuscito!");

                    }
                }

                return bOK;

            }
            catch (Exception ex)
            {
                return bOK;
            }
            

        }



    }
}
