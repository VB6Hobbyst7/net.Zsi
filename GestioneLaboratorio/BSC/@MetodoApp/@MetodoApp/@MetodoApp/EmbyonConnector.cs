using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetodoApp
{

    public class EmbyonConnector
    {

        public static dynamic imetodotarget;
        public static string ditta;
        public static string utente;

        public static bool InizializzaAtMetodo()
        {
            bool bOK = false;

            try
            {
                MetodoInterop.startlog();

                bOK = true;
            }
            catch (Exception ex)
            {

                MetodoInterop.LogError(ex.Message);
            
            }


            return bOK;

        }

        public static bool CheckEmbyonClient()
        {
            bool bOK = false;

            MetodoInterop.LogSomething("Check Embyon Client");
            try
            {
                if (InizializzaAtMetodo() == true)
                {
                    MetodoInterop.LogSomething(string.Format("Check Embyon Client {0}, {1}", ditta, utente));
        
                    imetodotarget = MetodoInterop.GetObjecFromRot(ditta, utente);

                    bOK = true;
                }
            }
            catch (Exception ex)
            {
                MetodoInterop.LogError(ex.Message);
            }

            return bOK;
        }


        public static bool ExecMetodoAction(string menuaction, string strkey)
        {
            bool bOK = false;

            string url = string.Format("metodo://MENU/{0}/{1}", menuaction, strkey);

            if (CheckEmbyonClient())
            {
                try
                {
                    MetodoInterop.LogSomething(string.Format("Action {0}, {1}", menuaction, strkey));
                    imetodotarget.NavigateTo(url);

                    if (imetodotarget == null)
                    {
                        //MessageBox.Show("Metodo non attivo!");
                        MetodoInterop.LogSomething(string.Format("Metodo non attivo! {0}, {1}", menuaction, strkey));
                    }
                    else
                    {
                        //controllo che la form sia aperta
                        MetodoInterop.LogSomething(string.Format("Action {0}", url));
                        CollectionWrapper2.MetodoHelper.WaitForHelpContextID(imetodotarget, 3000);

                        bOK = true;
                    }
                }
                catch (Exception ex)
                {
                    MetodoInterop.LogError(ex.Message);
                }
            }
            else
            { 
            
            }
            
            return bOK;

        }

        public static bool ExecMetodoAction(string menuactionstrkey)
        {
            bool bOK = false;

            string url = string.Format("metodo://MENU/{0}", menuactionstrkey);

            if (CheckEmbyonClient())
            {
                try
                {
                    MetodoInterop.LogSomething(string.Format(url));
                    imetodotarget.NavigateTo(url);

                    if (imetodotarget == null)
                    {
                        //MessageBox.Show("Metodo non attivo!");
                        MetodoInterop.LogSomething(string.Format("Metodo non attivo! {0}", url));
                    }
                    else
                    {
                        //controllo che la form sia aperta
                        MetodoInterop.LogSomething(string.Format("Action {0}", url));
                        CollectionWrapper2.MetodoHelper.WaitForHelpContextID(imetodotarget, 3000);

                        bOK = true;
                    }
                }
                catch (Exception ex)
                {
                    MetodoInterop.LogError(ex.Message);
                }
            }
            else
            {

            }

            return bOK;

        }
    }
}
