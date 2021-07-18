using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace MetodoApp
{
    internal class MetodoInterop
    {
        public static LogApplication.Logger log;

        public static void startlog()
        {
            log = new LogApplication.Logger();
            log.Setup();
            log.LogSomething("Start Log @metodo");
        }



        public static void LogSomething(string what)
        {
            log.LogSomething(string.Format("@metodo {0} - {1}", System.DateTime.Now.ToLongTimeString(), what));
        }

        public static void LogError(string what)
        {
            log.LogSomething(string.Format("@metodo ERRORE {0} - {1}", System.DateTime.Now.ToLongTimeString(), what));
        }

        public static object GetObjecFromRot(string ditta, string username)
        {
            string monikerkey = string.Concat("!METODO|", username, "|", ditta);
            log.LogSomething(monikerkey);

            object internalmxbrowser = null;
            if (ditta == string.Empty & username == string.Empty)
            {
                monikerkey = monikerkey.Substring(0, monikerkey.Length - 1);

                log.LogSomething(monikerkey);
            }

            try
            {
                IRunningObjectTable runningObjectTable = default(IRunningObjectTable);
                IEnumMoniker monikerEnumerator = null;
                IMoniker[] monikers = new IMoniker[2];
                runningObjectTable = NativeMethods.GetRunningObjectTable(0);
                runningObjectTable.EnumRunning(out monikerEnumerator);
                monikerEnumerator.Reset();

                IntPtr numFetched = new IntPtr();
                while ((monikerEnumerator.Next(1, monikers, numFetched) == 0))
                {

                    IBindCtx ctx = default(IBindCtx);
                    ctx = NativeMethods.CreateBindCtx(0);
                    string runningObjectName = "";
                    monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                    runningObjectName = runningObjectName.ToUpper();
                    log.LogSomething(runningObjectName);

                    if (runningObjectName.ToUpper().StartsWith(monikerkey, StringComparison.CurrentCultureIgnoreCase))
                    {
                        runningObjectTable.GetObject(monikers[0], out internalmxbrowser);
                        break;
                    }
                    //Console.WriteLine(runningObjectName);
                }
            }
            catch (Exception ex)
            {
                internalmxbrowser = null;
            }

            return internalmxbrowser;
        }
    }
}
