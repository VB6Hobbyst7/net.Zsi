using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;


namespace MetodoApp
{
    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        /// 



        public static string queryStringDescription = "";
        public const string UrlProtocol = "ATMETODO";
        public static bool bNotifica = false;

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetLongPathName(
            [MarshalAs(UnmanagedType.LPTStr)]
        string path,
            [MarshalAs(UnmanagedType.LPTStr)]
        StringBuilder longPath,
            int longPathLength
            );

        public static void RegisterUrlProtocol()
        {
            string subkeyValue = "\"" + Application.ExecutablePath + "\" %1";

            RegistryKey rKey = Registry.ClassesRoot.OpenSubKey(UrlProtocol, true);

            if (rKey == null)
            {
                //MessageBox.Show(string.Format("Creazione protocollo ATMETODO per l'estensione: {0}?", Application.ExecutablePath), "Impostazione Protocollo ATMETODO", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                rKey = Registry.ClassesRoot.CreateSubKey(UrlProtocol);
                rKey.SetValue("", "URL: ATMETODO Protocol");
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
                        MessageBox.Show(string.Format("Il protocollo ATMETODO è già registrato. Aggiornarlo per l'estensione: {0}?", Application.ExecutablePath), "Impostazione Protocollo ATMETODO", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
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

                if (args[0].Trim().ToUpper() == "ATMETODO" && args.Length > 1)
                {
                    // Means this is a URL protocol
                    string[] actionDetail = args[1].Split('?');

                    queryStringDescription = "argomenti: \r\n";

                    string[] qsDetails;

                    //if (actionDetail.Length > 1)
                    {

                        queryStringDescription += actionDetail[0].Trim().ToUpper();
                        switch (actionDetail[0].Trim().ToUpper().Replace(@"/", ""))
                        {
                            case "OPENMETODO":

                                qsDetails = actionDetail[1].Split('&');


                                for (int j = 0; j < qsDetails.Length; j++)
                                {
                                    string[] details = qsDetails[j].Split('=');

                                    for (int i = 0; i < details.Length; i++)
                                    {
                                        switch (details[0].ToUpper())
                                        {
                                            case "UTENTEDB":

                                                MetodoApp.Form1.UtenteDB = details[1];

                                                //SignFiles.startXML_baseurl = HttpUtility.UrlDecode(details[1]);

                                                //if (SignFiles.startXML_baseurl == "http")
                                                //{
                                                //    MessageBox.Show(args[2]);

                                                //}


                                                queryStringDescription += details[i].Trim() + "\r\n";
                                                break;

                                            case "DITTA":

                                                MetodoApp.Form1.Ditta = details[1];
                                                break;

                                            case "ACTION":

                                                MetodoApp.Form1.Action = details[1];
                                                break;

                                            case "KEY":

                                                MetodoApp.Form1.Key = details[1].Replace(@"'", "").Replace(@"#", "=");
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



        public static string checkURL(string _sURL)
        {
            string a = _sURL;

            if (a.IndexOf("baseurl=") < 0)
            {
                return a;
            }

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


        [STAThread]
        static void Main(string[] args)
        {

            // Register the client protocol for invoking the client
            // from a web page link
            try
            {
                RegisterUrlProtocol();
            }
            catch (Exception ex)
            {
            }


            if (args.Length == 0)
            {
                MessageBox.Show("Non è stato passato alcun parametro a @MetodoApp");
            }
            else
            {

                for (int i = 0; i < args.Count(); i++)
                {
                    string argParse = args[i];

                    if (argParse.Contains("~"))
                    {
                        StringBuilder longPath = new StringBuilder(255);
                        GetLongPathName(@argParse, longPath, longPath.Capacity);
                        argParse = longPath.ToString();
                    }

                    if (argParse.Contains(".knos-fr"))
                    {
                        //SignFiles.startXML = args[i];
                        break;
                    }
                }

                //if (SignFiles.startXML == "")
                //{
                    // Check protocol message
                    CheckForProtocolMessage();

                //}
            }



            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
