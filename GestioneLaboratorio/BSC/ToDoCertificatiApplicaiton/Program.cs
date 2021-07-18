using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Web;
using System.Runtime.InteropServices;
using System.Text;

namespace ToDoNotificheBSC
{
    static class Program
    {

        public static string queryStringDescription = "";
        public const string UrlProtocol = "KNOSAPI";

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetLongPathName(
            [MarshalAs(UnmanagedType.LPTStr)]
        string path,
            [MarshalAs(UnmanagedType.LPTStr)]
        StringBuilder longPath,
            int longPathLength
            );
        
        public static string checkURL(string _sURL)
        {
            string a = _sURL;
            string b = a.Substring(0, a.IndexOf("baseurl="));

            a = a.Substring(a.IndexOf("baseurl=") + 8);

            string c = "";

            if (a.IndexOf(':') > 0)
            {
                if (a.IndexOf('&') > 0)
                {
                    a = a.Substring(0, a.IndexOf('&'));
                    c = a.Substring(a.IndexOf('&'));

                }

                string x = "baseurl=" + HttpUtility.UrlEncode(a);
                //label1.Text = a;
                _sURL = b + '&' + x + c;
            }

            return _sURL;
        
        }

        public static void RegisterUrlProtocol()
        {
            string subkeyValue = "\"" + Application.ExecutablePath + "\" %1";

            RegistryKey rKey = Registry.ClassesRoot.OpenSubKey(UrlProtocol, true);

            if (rKey == null)
            {
                //MessageBox.Show(string.Format("Creazione protocollo KNOSAPI per l'estensione: {0}?", Application.ExecutablePath), "Impostazione Protocollo KNOSAPI", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                rKey = Registry.ClassesRoot.CreateSubKey(UrlProtocol);
                rKey.SetValue("", "URL: KnoSAPI Protocol");
                rKey.SetValue("URL Protocol", "");

                rKey = rKey.CreateSubKey(@"shell\open\command");
                rKey.SetValue("", subkeyValue);
            }
            if (rKey != null)
            {
                
                RegistryKey rSubKey;

                try
                {
                    rSubKey = rKey.OpenSubKey(@"shell\open\command", true);
                    //MessageBox.Show(rSubKey.GetValue("").ToString());

                    if (rSubKey.GetValue("").ToString() != subkeyValue)
                    {
                        MessageBox.Show(string.Format("Il protocollo KNOSAPI è già registrato. Aggiornarlo per l'estensione: {0}?", Application.ExecutablePath), "Impostazione Protocollo KNOSAPI", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        rSubKey.SetValue("", subkeyValue);
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                    rSubKey = rKey.CreateSubKey(@"shell\open\command");
                    rSubKey.SetValue("", subkeyValue);
                
                }

                

                rKey.Close();
            }
        }

        //Check for URL parameters
        private static bool CheckForProtocolMessage()
        {
            string[] arguments = Environment.GetCommandLineArgs();

            if (arguments.Length > 1)
            {
                string _queryString = arguments[1].ToString();

                _queryString = checkURL(_queryString); 

                // Format = "Owf:OpenForm?id=111"

                string[] args = _queryString.Split(':');

                if (args[0].Trim().ToUpper() == "KNOSAPI" && args.Length > 1)
                {
                    // Means this is a URL protocol
                    string[] actionDetail = args[1].Split('?');

                    queryStringDescription = "argomenti: \r\n";

                    //if (actionDetail.Length > 1)
                    {

                        queryStringDescription += actionDetail[0].Trim().ToUpper();
                        switch (actionDetail[0].Trim().ToUpper())
                        {





                            case "OPENTECEUROLABCERTIFICATIPDL":

                                string[] qsDetails = actionDetail[1].Split('&');
                                

                                for (int j = 0; j < qsDetails.Length; j++)
                                {
                                    string[] details = qsDetails[j].Split('=');

                                    for (int i = 0; i < details.Length; i++)
                                    {
                                        switch (details[0].ToUpper())
                                        {
                                            case "BASEURL":
                                                
                                                SignFiles.startXML_baseurl = HttpUtility.UrlDecode(details[1]);

                                                if (SignFiles.startXML_baseurl == "http")
                                                {
                                                    MessageBox.Show(args[2]);
                                                
                                                }


                                                queryStringDescription += details[i].Trim() + "\r\n";
                                                break;

                                            case "ID":

                                                int id = 0;
                                                int.TryParse(details[1], out id);
                                                SignFiles.startXML_idobject = id;
                                                break;

                                            case "IDCERT":

                                                int idCert = 0;
                                                int.TryParse(details[1], out idCert);
                                                SignFiles.startXML_idobject_certificato = idCert;
                                                break;

                                        }

                                    }
                                   
                                }

                                break;

                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            //SignFiles.tipofirma = -1;

            //// Register the client protocol for invoking the client
            //// from a web page link
            //try
            //{
            //    RegisterUrlProtocol();
            //}
            //catch (Exception ex)
            //{
            //}


            //if (args.Length == 0)
            //{
            //    MessageBox.Show("Non è stato passato alcun argomento al programma di firma");
            //}
            //else 
            //{
            //    SignFiles.startXML = "";

            //    for (int i = 0; i < args.Count(); i++)
            //    {
            //        string argParse = args[i];

            //        if (argParse.Contains("~"))
            //        {
            //            StringBuilder longPath = new StringBuilder(255);
            //            GetLongPathName(@argParse, longPath, longPath.Capacity);
            //            argParse = longPath.ToString();
            //        }

            //        if (argParse.Contains(".knos-fr"))
            //        {
            //            SignFiles.startXML = args[i];
            //            break;
            //        }
            //    }

            //    if (SignFiles.startXML == "")
            //    {
            //        // Check protocol message
            //        CheckForProtocolMessage();            

            //    }

            //    /*

            //        if (args[i].Substring(0, 2) == "/D")
            //        {
            //            if (File.Exists(args[i].Substring(2)))
            //            {
            //                SignFiles.fileRTF = args[i].Substring(2);
            //                File.Delete(SignFiles.tempOriginalPDF);
            //            }

            //        }

            //        if (args[i].Substring(0, 2) == "/S")
            //        {
            //            if (File.Exists(args[i].Substring(2)))
            //            {
            //                SignFiles.filePNG = args[i].Substring(2);
            //            }

            //        }

            //        if (args[i].Substring(0, 5) == "/TEST")
            //        {
            //            SignFiles.testSignPDF = true;
            //        }

            //        if (args[i].Substring(0, 10) == "/TIPOFIRMA")
            //        {
            //            try
            //            {
            //                SignFiles.tipofirma = int.Parse(args[i].Substring(10));
            //            }
            //            catch (Exception ex)
            //            {
            //            }
            //        }

            //    }

            //    */


            //}


            if (args.Length == 0)
            {
                //MessageBox.Show("Non è stato passato alcun argomento al programma di firma");

                try
                {

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    //Application.Run(new frmToDoNotificheBSC());
                    Application.Run(new frmMagazzino());
                    //Application.Run(new frmIngressoSpedizionieri());
                    //SingleInstance.SingleApplication.Run(new frmToDoNotificheBSC());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                SignFiles.startXML = "";

                for (int i = 0; i < args.Count(); i++)
                {
                    string argParse = args[i];

                    if (argParse.Contains("/MDI"))
                    {
                        try
                        {

                            Application.EnableVisualStyles();
                            Application.SetCompatibleTextRenderingDefault(false);
                            //Application.Run(new frmToDoNotificheBSC());
                            Application.Run(new ZSIManager());
                            //Application.Run(new frmIngressoSpedizionieri());
                            //SingleInstance.SingleApplication.Run(new frmToDoNotificheBSC());
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                   
                }
                


            }


            
        }
    }
}
