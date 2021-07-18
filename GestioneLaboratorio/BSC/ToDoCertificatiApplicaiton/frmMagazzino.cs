using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;

using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using System.IO;

using System.Runtime.InteropServices;

using Knos;
using Knos.API.NET;
using Knos.API.COM;
using System.Net;

using System.Reflection;
using System.Diagnostics;


//using Outlook = Microsoft.Office.Interop.Outlook;
using System.Configuration;
using Telerik.WinControls.UI;
using Telerik.WinControls;
using Telerik.WinControls.Data;
using Telerik.WinControls.UI.Export;
using Telerik.WinControls.Export;

using ItaCalendar;





namespace ToDoNotificheBSC
{

    public partial class frmMagazzino : Form
    {
        Logger log;

        public static int CurrentIDStatusPDL = 0;
        public static string CurrentStatusNamePDL = "";
        public static string CurrentPDFPDLUrl = "";

        string CurrentUser = "";

        public static string CurrentTecnico = "";
        public static string CurrentResponsabileTecnico = "";
        public static string CurrentCapoCommessa = "";


        // variabili riassuntive stato certificati
        public static int nrCertificatiTot = 0;
        public static int nrCertificatiUtente = 0;
        public static int nrCertificatiUtente1F = 0;
        public static int nrCertificatiUtente2F = 0;
        public static int nrCertificatiUtente1FDaFirmare = 0;
        public static int nrCertificatiUtente2FDaFirmare = 0; 
        public static int nrCertificati1F = 0;
        public static int nrCertificati2F = 0;

        public static string strFilePDFPDL = "";

        // @Metodo
        //Richiama documento ordine utilizzando l'indirizzo "metodo://MENU/GestioneDoc_1/@Progressivo=657"
        string actionRip = "GestioneDoc_1";
        string actionRipKey = "@Progressivo={0}";

        enum  TipoUtente
        {
            Magazzino = 3,
            BackOffice = 4
        };


        public bool notifyPopUp = true;

        int idtesta = 0;
        int idriga = 0;
        int nrpezziimballo = 0;
        string articolo = "";

        bool posizionato = false;

        string s = Path.Combine(Application.StartupPath, "GridLayout");

        string sqlConnectionString = "";

        bool bolShowTotali = true;
        int hRows = 40;


        RadContextMenu contextMenu;

        public class KnoSWrapper
        {

            Logger logDMS;

            IKnosObject knosObject;
            IKnosObject knosObjectCertificato;
            IKnosObjectMaker knosObjectMaker;
            IKnosObject knosObjectCliente;

            int cIdSubject = 0;
            string cUserName = "";

            public string DefaultSite;
            public string CurrentUser;
            public string PWD;

            public bool Inizializza(string _defaultSite = "")
            {
                logDMS = new Logger();

                logDMS.Setup();

                bool retvalue = false;

                if (_defaultSite == "")
                {
                    _defaultSite = KnosInstance.DefaultSite;
                }
                else
                {

                }

                KnosInstance.Initialization();
                IKnosResult ikr = KnosInstance.Open(_defaultSite);

                if (ikr.NoErrors == true)
                {
                    try
                    {

                        // cIdSubject = KnosInstance.Client.GetCurrentIdSubject();
                        ikr = KnosInstance.Client.CheckCurrentUser(out cIdSubject, out cUserName);

                        //if (ikr.HasErrors == false)
                        if (cIdSubject > 0)
                        {
                            DefaultSite = _defaultSite;
                            CurrentUser = cUserName;
                            retvalue = true;
                        }
                        else
                        {
                            CurrentUser = Properties.Settings.Default.KnoS_User;
                            PWD = Properties.Settings.Default.KnoS_PWD;

                            if (KnosInstance.Client.Login(CurrentUser, PWD, out cIdSubject).NoErrors)
                            {
                                retvalue = true;

                            }
                            else
                            {
                                retvalue = false;
                                //MessageBox.Show("Utente non loggato da Internet Explorer");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("Impossibile aprire KnoS all'indirizzo{0}", _defaultSite));

                        retvalue = false;
                    }
                }


                return retvalue;


            }



            //    public List<Allegato> GetAllegatiPubblicazione(int _idObject)
            //    {
            //        bool retvalue = false;
            //        string fileName = "";
            //        string fileUrl = "";
            //        string fileLocalPath = "";
            //        int fileIdDoc = 0;
            //        string fileDescr = "";

            //        knosObject = KnosInstance.Client.CreateKnosObject();

            //        List<Allegato> _allegati = new List<Allegato>();

            //        try
            //        {
            //            IKnosResult ikr = knosObject.GetObjectLinks(_idObject);

            //            if (ikr.HasErrors == false)
            //            {

            //                for (int i = 0; i < knosObject.LinkList.ItemCount; i++)
            //                {
            //                    fileDescr = knosObject.LinkList.GetItem(i).LinkDescr;
            //                    fileUrl = knosObject.LinkList.GetItem(i).Url;
            //                    fileName = knosObject.LinkList.GetItem(i).Url;

            //                    Allegato a = new Allegato(fileDescr, fileDescr, fileUrl.Replace("file://", ""));

            //                    _allegati.Add(a);
            //                }

            //            }

            //        }

            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(string.Format("Errore \r\n{0}", ex.Message));

            //        }

            //        return _allegati;
            //    }


            public bool UploadFileCertificato(int _idObject,
            int _idDoc,
            string _filePath,
            string _fileDescr,
            string _fileName,
            int _actionWF,
            string _attrNameDate,
            string versione
            )
        {
            bool retvalue = false;
            IKnosUploadInfo kui;

            // upload file
            IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


            Cursor c;

            c = Cursors.WaitCursor;

            if (_filePath != "")
            {

                // cancello il documento attualmente presente
                logDMS.LogSomething(string.Format("Cancello il file presente (IdObject - IdDoc) : {0} - {1}", _idObject, _idDoc));
                IKnosResult ikr = knosObjectMaker.DeleteDoc(_idObject, _idDoc);
                if (ikr.HasErrors == false)
                {
                    retvalue = true;
                }
                else
                {
                    if (ikr.GetError(0).Number == 20014)
                    {
                        retvalue = true;
                    }
                    else
                    {
                        c = Cursors.Default;

                        MessageBox.Show(ikr.GetError(0).Description);
                        return retvalue;
                    }
                }



                IKnosUploadItem ui = KnosInstance.Client.CreateKnosUploadItem();

                ui.FileDescr = _fileDescr;
                ui.FileName = _fileName;

                if (_filePath.EndsWith("\\"))
                {
                    ui.FilePath = Path.Combine(_filePath, _fileName);
                }
                else
                {
                    ui.FilePath = _filePath;
                }

                int iversione = 1;
                int.TryParse(versione, out iversione);
                if (iversione > 1)
                {
                    ui.Version = iversione;
                }
                knosObjectMaker.AddUploadItem(ui);

                logDMS.LogSomething(string.Format("Upload nuovo file : {0}", ui.FilePath));

                ikr = knosObjectMaker.UploadFiles(_idObject, out kui);

                if (ikr.HasErrors == false)
                {
                    retvalue = true;
                }
                else
                {
                    c = Cursors.Default;
                    MessageBox.Show(ikr.GetError(0).Description);

                    for (int iError = 0; iError < ikr.ErrorCount; iError++)
                    {
                        logDMS.LogSomething(ikr.GetError(iError).Description);
                    }

                    for (int iError = 0; iError < ikr.WarningCount; iError++)
                    {
                        logDMS.LogSomething(ikr.GetWarning(iError).Description);
                    }
                }

                // se qualcosa è andato storto esco
                if (retvalue == false)
                {
                    return retvalue;
                }

            }


            if (_actionWF > 0)
            {
                retvalue = true;

                // altrimenti faccio la transizione di stato della pubblicazione certificato

                IKnosActionNotify knosActionWF = KnosInstance.Client.CreateKnosActionNotify();
                knosActionWF.IdAction = _actionWF;
                IKnosResult knosResult = knosObjectMaker.ChangeStatusByAction(_idObject, knosActionWF, out retvalue);

                // se qualcosa è andato storto esco
                if (knosResult.HasWarningsErrors)
                {
                    return retvalue;
                }

                // aggiornamento attributo cambio stato 
                knosObjectMaker.SetAttrValue(_attrNameDate, System.DateTime.Now, EnumKnosDataType.DateTimeType);

                knosResult = knosObjectMaker.UpdateObject(_idObject);

                if (knosResult.HasWarningsErrors)
                {
                    c = Cursors.Default;

                    MessageBox.Show(knosResult.ToString());
                }
                else
                {
                    retvalue = true;
                }
            }

            return retvalue;


        }


            //    public bool DeleteFiles(int _idObject,
            //        int _idDoc
            //        )
            //    {
            //        bool retvalue = false;
            //        IKnosUploadInfo kui;

            //        // delete file
            //        IKnosObject knosObject = KnosInstance.Client.CreateKnosObject();
            //        IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


            //        Cursor c;

            //        c = Cursors.WaitCursor;

            //        IKnosResult ikr = knosObject.GetObjectDocuments(_idObject);

            //        if (ikr.HasErrors == false)
            //        {
            //            if(_idDoc >= 0)
            //            {
            //                // cancello il documento attualmente presente
            //                ikr = knosObjectMaker.DeleteDoc(_idObject, _idDoc);
            //                if (ikr.HasErrors == false)
            //                {
            //                    retvalue = true;
            //                }
            //                else
            //                {
            //                    if (ikr.GetError(0).Number == 20014)
            //                    {
            //                        retvalue = true;
            //                    }
            //                    else
            //                    {
            //                        c = Cursors.Default;

            //                        MessageBox.Show(ikr.GetError(0).Description);
            //                        return retvalue;
            //                    }
            //                }

            //            }
            //            else
            //            {
            //                for (int i = 0; i < knosObject.DocumentList.ItemCount; i++)
            //                {
            //                    int iddoc = 0;
            //                    int.TryParse(knosObject.DocumentList.GetItem(i).IdDoc.ToString(), out iddoc);

            //                    // cancello il documento attualmente presente
            //                    ikr = knosObjectMaker.DeleteDoc(_idObject, iddoc);
            //                    if (ikr.HasErrors == false)
            //                    {
            //                        retvalue = true;
            //                    }
            //                    else
            //                    {
            //                        if (ikr.GetError(0).Number == 20014)
            //                        {
            //                            retvalue = true;
            //                        }
            //                        else
            //                        {
            //                            c = Cursors.Default;

            //                            MessageBox.Show(ikr.GetError(0).Description);
            //                            return retvalue;
            //                        }
            //                    }                            
            //                }

            //            }
            //        }

            //        return retvalue;


            //    }

            //    public bool EseguiAzione(int _idObject,
            //            int _actionWF,
            //            string _attrNameDate
            //    )
            //    {
            //        bool retvalue = false;
            //        IKnosUploadInfo kui;

            //        // upload file
            //        IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


            //        Cursor c;

            //        c = Cursors.WaitCursor;

            //        if (_actionWF > 0)
            //        {
            //            retvalue = true;

            //            // altrimenti faccio la transizione di stato della pubblicazione certificato

            //            IKnosActionNotify knosActionWF = KnosInstance.Client.CreateKnosActionNotify();
            //            knosActionWF.IdAction = _actionWF;
            //            IKnosResult knosResult = knosObjectMaker.ChangeStatusByAction(_idObject, knosActionWF, out retvalue);

            //            // se qualcosa è andato storto esco
            //            if (knosResult.HasWarningsErrors)
            //            {
            //                return retvalue;
            //            }

            //            if (_attrNameDate != "")
            //            {

            //                // aggiornamento attributo cambio stato 
            //                knosObjectMaker.SetAttrValue(_attrNameDate, System.DateTime.Now, EnumKnosDataType.DateTimeType);

            //                knosResult = knosObjectMaker.UpdateObject(_idObject);

            //                if (knosResult.HasWarningsErrors)
            //                {
            //                    c = Cursors.Default;

            //                    MessageBox.Show(knosResult.ToString());
            //                }
            //                else
            //                {
            //                    retvalue = true;
            //                }
            //            }
            //        }

            //        return retvalue;

            //    }



            //    public bool EseguiAzioneWS(int idObject,
            //            int _actionWF,
            //            string _attrNameDate)
            //    {


            //        bool bOK = false;
            //        IKnosResult result;


            //        //Inizializza(KnosInstance.Client.KnosBaseUrl);

            //        IKnosRequest request = KnosInstance.Client.CreateKnosRequest();



            //        request.SetParameter("IdObject", idObject.ToString());
            //        request.SetParameter("IdAction", _actionWF.ToString());
            //        request.SetParameter("IgnoreError", "1");
            //        request.SetParameter("SkipAction", "0");
            //        request.SetParameter("SkipNotify", "2");
            //        request.SetParameter("CheckObjectUnlock", "0");
            //        IKnosResponse response;

            //        result = KnosInstance.Client.ParseResponse(string.Format("{0}/knos/system/webservices/object_changestatusbyaction.asp", KnosInstance.Client.KnosBaseUrl), ref request, out response);

            //        if (result.NoErrors == true)
            //        {
            //            bOK = true;// Pubblicazione bloccata, si può elaborare
            //        }
            //        //else
            //        //{
            //        //    ;// Pubblicazione non bloccata, si deve saltare 
            //        //}

            //        return bOK;

            //    }

            //    public bool EliminaAllegato(int _idObject,
            //                        int _idDoc,
            //                        string _attrNameDate
            //                )
            //    {
            //        bool retvalue = false;
            //        IKnosUploadInfo kui;

            //        // upload file
            //        IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


            //        Cursor c;

            //        c = Cursors.WaitCursor;

            //        if ((_idObject > 0) && (_idDoc>0))
            //        {
            //            retvalue = true;

            //            // altrimenti faccio la transizione di stato della pubblicazione certificato

            //            IKnosResult knosResult = knosObjectMaker.DeleteDoc(_idObject, _idDoc);

            //            // se qualcosa è andato storto esco
            //            if (knosResult.HasWarningsErrors)
            //            {
            //                return retvalue;
            //            }

            //        }

            //        return retvalue;

            //    }




            //    public bool downloadDoc(int _idCertificato, int _idDoc = 1, string filePath = "")
            //    {
            //        IKnosResult ikr;
            //        bool bOK = false;

            //        IKnosObject knosObject = KnosInstance.Client.CreateKnosObject();

            //        ikr = knosObject.GetObjectDocuments(_idCertificato);
            //        if (ikr.HasErrors == false)
            //        {
            //            //download local del file
            //            for (int i = 0; i < knosObject.DocumentList.ItemCount; i++)
            //            {
            //                if (_idDoc == knosObject.DocumentList.GetItem(i).IdDoc)
            //                {
            //                    ikr = knosObject.DocumentList.GetItem(i).DownloadFile(Path.GetTempPath(), filePath);
            //                    break;
            //                }
            //            }
            //        }

            //        if (ikr.HasErrors == false)
            //        {
            //            return true;
            //        }
            //        else
            //        {
            //            MessageBox.Show(string.Format("{0}\\{1} \r\n {2}", Path.GetTempPath(), filePath, ikr.GetError(0).Description), "Errore in Download allegato");
            //            return false;
            //        }
            //    }


            //    public bool downloadDoc(int _idCertificato, int _idDoc = 1, string filePath = "", string filename = "")
            //    {
            //        IKnosResult ikr;
            //        bool bOK = false;

            //        IKnosObject knosObject = KnosInstance.Client.CreateKnosObject();

            //        ikr = knosObject.GetObjectDocuments(_idCertificato);
            //        if (ikr.HasErrors == false)
            //        {
            //            //download local del file
            //            for (int i = 0; i < knosObject.DocumentList.ItemCount; i++)
            //            {
            //                if (_idDoc == knosObject.DocumentList.GetItem(i).IdDoc)
            //                {
            //                    ikr = knosObject.DocumentList.GetItem(i).DownloadFile(filePath, filename);
            //                    break;
            //                }
            //            }
            //        }

            //        if (ikr.HasErrors == false)
            //        {
            //            return true;
            //        }
            //        else
            //        {
            //            MessageBox.Show(string.Format("{0}\\{1} \r\n {2}", filePath, filename, ikr.GetError(0).Description), "Errore in Download allegato");
            //            return false;
            //        }
            //    }




            public string GetEmailSubjectByName(string _name)
            {
                string outEmail = "";

                IKnosSubjectMaker knosSubjectMaker = KnosInstance.Client.CreateKnosSubjectMaker();

                IKnosResult kr = knosSubjectMaker.GetSubject(0, _name);

                if (kr.HasErrors == false)
                {
                    if (knosSubjectMaker.IdSubject == 0)
                    {
                        outEmail = "utente non trovato";
                    }
                    else
                    {
                        outEmail = knosSubjectMaker.Email;
                    }

                }

                return outEmail;



            }


            public int GetIdSubjectByName(string _name)
            {
                int idSubject = 0;

                IKnosSubjectMaker knosSubjectMaker = KnosInstance.Client.CreateKnosSubjectMaker();

                IKnosResult kr = knosSubjectMaker.GetSubject(0, _name);

                if (kr.HasErrors == false)
                {
                    if (knosSubjectMaker.IdSubject == 0)
                    {
                        idSubject = 0;
                    }
                    else
                    {
                        idSubject = knosSubjectMaker.IdSubject;
                    }

                }

                return idSubject;



            }


            //    public bool GetSignImage(int _idObject, ListView lvFirme, string _signer )
            //    {
            //        bool retvalue = false;

            //        foreach (ListViewItem li in lvFirme.Items)
            //        { 
            //            if (li.Text == _signer)
            //            {
            //                SignFiles.filePNG = (li.SubItems[1].Text);
            //                retvalue = true;
            //                break;
            //            }

            //        }


            //        //lvFirme.Clear();
            //        //lvFirme.Columns.Clear();
            //        //lvFirme.Columns.Add("Utente");
            //        //lvFirme.Columns.Add("PathFileFirma");


            //        //knosObjectCertificato = KnosInstance.Client.CreateKnosObject();

            //        //IKnosResult ikr = knosObjectCertificato.GetObjectLinks(_idObject);

            //        //if (ikr.HasErrors == false)
            //        //{

            //        //    for (int i = 0; i < knosObjectCertificato.LinkList.ItemCount; i++)
            //        //    {
            //        //        lvFirme.Items.Add(knosObjectCertificato.LinkList.GetItem(i).LinkDescr);


            //        //        if (knosObjectCertificato.LinkList.GetItem(i).Url.StartsWith("file:"))
            //        //        { 
            //        //            lvFirme.Items[i].SubItems.Add(knosObjectCertificato.LinkList.GetItem(i).Url.ToString().Replace(@"file://", ""));

            //        //        }


            //        //        if (knosObjectCertificato.LinkList.GetItem(i).LinkDescr == _signer)
            //        //        {
            //        //            SignFiles.filePNG = (knosObjectCertificato.LinkList.GetItem(i).Url.ToString().Replace(@"file://", ""));
            //        //            retvalue = true;
            //        //        }

            //        //    }


            //        //}
            //        //else
            //        //{
            //        //    MessageBox.Show(ikr.GetError(0).Description);
            //        //}

            //        return retvalue;


            //    }

                

            }

            //string nomeRTF = "";
            //string nomePDF = "";
            //bool opened = false;
            KnoSWrapper kw = new KnoSWrapper();

            //Microsoft.Office.Interop.Word.Document wd ;
            //Microsoft.Office.Interop.Word.Application wa;

            public static long GetFileSizeOnDisk(string file)
        {
            FileInfo info = new FileInfo(file);
            uint dummy, sectorsPerCluster, bytesPerSector;
            int result = GetDiskFreeSpaceW(info.Directory.Root.FullName, out sectorsPerCluster, out bytesPerSector, out dummy, out dummy);
            if (result == 0) throw new Win32Exception();
            uint clusterSize = sectorsPerCluster * bytesPerSector;
            uint hosize;
            uint losize = GetCompressedFileSizeW(file, out hosize);
            long size;
            size = (long)hosize << 32 | losize;
            return ((size + clusterSize - 1) / clusterSize) * clusterSize;
        }


        [DllImport("kernel32.dll")]
        static extern uint GetCompressedFileSizeW([In, MarshalAs(UnmanagedType.LPWStr)] string lpFileName,
           [Out, MarshalAs(UnmanagedType.U4)] out uint lpFileSizeHigh);

        [DllImport("kernel32.dll", SetLastError = true, PreserveSig = true)]
        static extern int GetDiskFreeSpaceW([In, MarshalAs(UnmanagedType.LPWStr)] string lpRootPathName,
           out uint lpSectorsPerCluster, out uint lpBytesPerSector, out uint lpNumberOfFreeClusters,
           out uint lpTotalNumberOfClusters);

       
 
        public frmMagazzino()
        {
            InitializeComponent();
        }




        private void frmToDoNotificheBSC_Load(object sender, EventArgs e)
        {
            sqlConnectionString = Properties.Settings.Default.MetodoConnectionString;


            //if (!Properties.Settings.Default.MetodoConnectionStringUSER.Contains("{0}"))
            //{
            //    sqlConnectionString = Properties.Settings.Default.MetodoConnectionStringUSER;
            //}

            //if (CurrentUser != "")
            //{
            //    toolStripStatusLabelCurrentUser.Text = "Benvenuto " + Properties.Settings.Default.CurrentUser;

            //}


            if (clsLogin.loadCredenziali())
            {
                CurrentUser = clsLogin.CurrentUser;
            }
            else
            {
                frmLogin f = new frmLogin();
                f.StartPosition = FormStartPosition.CenterParent;
                f.ShowDialog();
            }

            CurrentUser = clsLogin.CurrentUser;
            sqlConnectionString = clsLogin.MetodoConnectionStringUSER;

            toolStripStatusLabelCurrentUser.Text = "Benvenuto " + CurrentUser;


            if (Properties.Settings.Default.RicercaAutomaticaSchede)
            {
                this.Visible = false;
            }

            log = new Logger();
            log.Setup();
            log.LogSomething("Start servizio");

            this.Text = string.Format("ZSI - MAGAZZINO ({0}) - Utente connesso: {1}", Application.ProductVersion, CurrentUser);

            try
            {

                //opened = false;

                bool.TryParse(Properties.Settings.Default.sendMailPopUp, out notifyPopUp);

                //Knos
                //if (refreshKnosLogin() == false)
                //return;
                refreshKnosLogin();

            }
            catch (Exception ex)
            {
                //return;
                log.LogSomething(ex.Message);
            }

            checkBoxPopUpMail.Checked = notifyPopUp;

            // caricamento automatico schede da epy
            dTP_BOLLEDA.Value = dTP_COADA.Value = new System.DateTime((DateTime.Today.Year - 1), 1, 1);
            dTP_BOLLEA.Value = dTP_COAA.Value = System.DateTime.Today.AddMonths(12);

            dTP_POSIZDA.Value = dTP_COADA.Value = new System.DateTime((DateTime.Today.Year - 1), 1, 1);
            dTP_POSIZA.Value = dTP_COAA.Value = System.DateTime.Today.AddMonths(12);

            cmbIMBALLI.SelectedIndex = Properties.Settings.Default.MAGTipiImballo;
            cmbSTATOM05.SelectedIndex = Properties.Settings.Default.MAGTStatoM05;
            cmbTIPIORDINI.SelectedIndex = Properties.Settings.Default.MAGTipoOrdini;




            //splitContainer2.SplitterDistance = splitContainer2.Width - 400;

            getImpostazioniGriglia();

            for (int i = 0; i< Properties.Settings.Default.sendMailMAGBackOffice.Count; i++)
            {
                checkedListBoxMailTo.Items.Add(Properties.Settings.Default.sendMailMAGBackOffice[i]);
            }

            chkFP.Checked = false;

            Application.DoEvents();

            



            this.WindowState = FormWindowState.Maximized;


            
            //contextMenu = new RadContextMenu();
            //RadMenuItem menuDispComp = new RadMenuItem("Disponibilità Completa (X)");
            //menuDispComp.ForeColor = Color.DarkGreen;
            //menuDispComp.Click += new EventHandler(menuDispComp_Click);
            //RadMenuItem menuDispParz = new RadMenuItem("Disponibilità Parziale");
            //menuDispParz.Click += new EventHandler(menuDispParz_Click);
            //contextMenu.Items.Add(menuDispComp);
            //contextMenu.Items.Add(menuDispParz);

        }



        public static void CombineMultiplePDFs(string[] fileNames, string outFile)
            {
            //int pageOffset = 0;
            //int f = 0;
            //iTextSharp.text.Document document = null;
            //PdfCopy writer = null;
            //PdfReader reader = null;
            //while (f < fileNames.Length)
            //{
            //    // we create a reader for a certain document
            //    reader = new PdfReader(fileNames[f]);
            //    reader.ConsolidateNamedDestinations();
            //    // we retrieve the total number of pages
            //    int n = reader.NumberOfPages;
            //    pageOffset += n;
            //    if (f == 0)
            //    {
            //        // step 1: creation of a document-object
            //        document = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));
            //        // step 2: we create a writer that listens to the document
            //        writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
            //        // step 3: we open the document
            //        document.Open();
            //    }
            //    // step 4: we add content
            //    for (int i = 0; i < n; )
            //    {
            //        ++i;
            //        if (writer != null)
            //        {
            //            PdfImportedPage page = writer.GetImportedPage(reader, i);
            //            writer.AddPage(page);
            //        }
            //    }
            //    PRAcroForm form = reader.AcroForm;
            //    if (form != null && writer != null)
            //    {
            //        writer.CopyAcroForm(reader);
            //    }

            //    if (reader != null)
            //    {
            //        reader.Close();
            //        reader.Dispose();
            //    }


            //    f++;
            //}

            //// step 5: we close the document
            //if (document != null)
            //{
            //    document.Close();
            //}              

            //writer.Dispose();
        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                btnCercaCOA_Click(null, null);
            }


            if (tabControl1.SelectedIndex == 1)
            {
                LoadItems();
            }


            if (tabControl1.SelectedIndex == 2)
            {
                rGVRichiestaProduzione.DataSource = getRichiesteProduzione();
                rGVDisponibilitaMagazzino.DataSource = getDisponibilita();
                rGVODPLotti.DataSource = getOrdiniProduzioneLOTTI();

                rGVRichiestaProduzione.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.AllCells);
                rGVDisponibilitaMagazzino.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.AllCells);
                rGVODPLotti.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.AllCells);

                this.rGVRichiestaProduzione.SummaryRowsTop.Clear();
                this.rGVDisponibilitaMagazzino.SummaryRowsTop.Clear();
                this.rGVODPLotti.SummaryRowsTop.Clear();

                this.rGVRichiestaProduzione.MasterTemplate.ShowTotals = true;
                GridViewSummaryItem summaryItemP = new GridViewSummaryItem("QTAGESTRES", "{0:0,0}", GridAggregateFunction.Sum);
                GridViewSummaryItem summaryItemFustiP = new GridViewSummaryItem("NRPEZZIIMBALLO", "{0:0,0}", GridAggregateFunction.Sum);

                GridViewSummaryRowItem summaryRowItemP = new GridViewSummaryRowItem();
                summaryRowItemP.Add(summaryItemP);
                summaryRowItemP.Add(summaryItemFustiP);
                this.rGVRichiestaProduzione.SummaryRowsTop.Add(summaryRowItemP);


                this.rGVDisponibilitaMagazzino.MasterTemplate.ShowTotals = true;
                GridViewSummaryItem summaryItemD = new GridViewSummaryItem("QTAGESTRES", "{0:0,0}", GridAggregateFunction.Sum);
                GridViewSummaryItem summaryItemFustiD = new GridViewSummaryItem("NRPEZZIIMBALLO", "{0:0,0}", GridAggregateFunction.Sum);
                GridViewSummaryItem summaryItemDispD = new GridViewSummaryItem("DISPONIBILITA", "{0:0,0}", GridAggregateFunction.Sum);
                GridViewSummaryItem summaryItemColliDaSpedD = new GridViewSummaryItem("COLLIDASPEDIRE", "{0:0,0}", GridAggregateFunction.Sum);
                GridViewSummaryItem summaryItemColliSpedibiliD = new GridViewSummaryItem("COLLISPEDIBILI", "{0:0,0}", GridAggregateFunction.Sum);



                GridViewSummaryRowItem summaryRowItemD = new GridViewSummaryRowItem();
                summaryRowItemD.Add(summaryItemD);
                summaryRowItemD.Add(summaryItemFustiD);
                summaryRowItemD.Add(summaryItemDispD);
                summaryRowItemD.Add(summaryItemColliDaSpedD);
                summaryRowItemD.Add(summaryItemColliSpedibiliD);
                this.rGVDisponibilitaMagazzino.SummaryRowsTop.Add(summaryRowItemD);



                this.rGVODPLotti.MasterTemplate.ShowTotals = true;
                GridViewSummaryItem summaryItemODPD = new GridViewSummaryItem("QTAGESTIONERES", "{0:0,0}", GridAggregateFunction.Sum);



                GridViewSummaryRowItem summaryRowItemODPD = new GridViewSummaryRowItem();
                summaryRowItemODPD.Add(summaryItemODPD);
                this.rGVODPLotti.SummaryRowsTop.Add(summaryRowItemODPD);

            }

        }


        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>

        void LoadItems()
        {

            itaCalendarObject1.chkViewITALIA = itaCalendarObject1.chkViewESTERO = 0;

            if ((cmbTIPIORDINI.SelectedIndex == 0) || (cmbTIPIORDINI.SelectedIndex == 1))
            {
                itaCalendarObject1.chkViewITALIA = 1;
            }

            if ((cmbTIPIORDINI.SelectedIndex == 0) || (cmbTIPIORDINI.SelectedIndex == 2))
            {
                itaCalendarObject1.chkViewESTERO = 1;
            }

            itaCalendarObject1.chkViewNONCISTERNA = itaCalendarObject1.chkViewCISTERNA = 0;

            if ((cmbIMBALLI.SelectedIndex == 0) || (cmbIMBALLI.SelectedIndex == 1))
            {
                itaCalendarObject1.chkViewNONCISTERNA = 1;
            }

            if ((cmbIMBALLI.SelectedIndex == 0) || (cmbIMBALLI.SelectedIndex == 2))
            {
                itaCalendarObject1.chkViewCISTERNA = 1;
            }

            itaCalendarObject1.loadfilter();

            itaCalendarObject1.LoadItems();

            //bindingSource2.DataSource = radGridViewCOA.DataSource;

            ////SchedulerDayView dayView = this.radScheduler1.GetDayView();

            ////dayView.WorkTime = TimeInterval.DefaultWorkTime;

            ////dayView.RulerStartScale = 7;
            ////dayView.RulerEndScale = 19;
            //calendarUserControl1.ExternaItemsList.Clear();

            //DateTime prevdate = System.DateTime.Today;
            //int prevdatecount = 0;

            //for (int i = 0; i < ((DataTable)bindingSource2.DataSource).Rows.Count; i++)
            //{


            //    if (prevdate == Convert.ToDateTime(((DataTable)bindingSource2.DataSource).Rows[i]["DATACARICO"].ToString()))
            //    {
            //        prevdatecount++;
            //    }
            //    else
            //    {
            //        prevdatecount = 0;
            //    }

            //    calendarUserControl1.ExternaItemsList.Add(new System.Windows.Forms.Calendar.ItemInfo(
            //        Convert.ToDateTime(((DataTable)bindingSource2.DataSource).Rows[i]["DATACARICO"].ToString()).AddHours(8+prevdatecount)
            //        , Convert.ToDateTime(((DataTable)bindingSource2.DataSource).Rows[i]["DATACARICO"]).AddHours(9+prevdatecount)
            //        , string.Format("{0} - {1} - {2} - {3} - {4} - {5} - {6}"
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["XLS_SPEDIZIONIERE"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["DESCRIZIONE"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["RAGIONESOCIALE"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["XLS_QTA"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["IMBALLO"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["NRLOTTO"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["DOCUMENTO"].ToString()
            //            , ((DataTable)bindingSource2.DataSource).Rows[i]["COLLISPEDIBILI"].ToString()
            //            )
            //        )
                        
            //    );

            //    prevdate = Convert.ToDateTime(((DataTable)bindingSource2.DataSource).Rows[i]["DATACARICO"].ToString());

            //}


            //calendarUserControl1.mvStartDate = dTP_BOLLEDA.Value;
            //calendarUserControl1.mvEndDate = dTP_BOLLEA.Value;

            ////calendarUserControl1.SetMonthViewStartEnd;
            //calendarUserControl1.LoadCalendar();

            ////foreach (CalendarItem calendarItem in _items)
            ////{
            ////    if (this.calendar1.ViewIntersects(calendarItem))
            ////    {
            ////        this.calendar1.Items.Add(calendarItem);
            ////    }
            ////}

            ////this.calendar1.SetViewRange(dTP_BOLLEDA.Value, dTP_BOLLEA.Value);






        }


        private void DownoladFileFromUrl(string m_uri, string m_filePath)
        {

            HttpWebRequest request;

            HttpWebResponse response = null;

            try
            {

                request = (HttpWebRequest)WebRequest.Create(m_uri);

                request.Timeout = 10000;

                request.AllowWriteStreamBuffering = false;

                response = (HttpWebResponse)request.GetResponse();

                Stream s = response.GetResponseStream();



                //Write to disk

                FileStream fs = new FileStream(m_filePath, FileMode.Create);

                byte[] read = new byte[256];

                int count = s.Read(read, 0, read.Length);

                while (count > 0)
                {

                    fs.Write(read, 0, count);

                    count = s.Read(read, 0, read.Length);

                }

                //Close everything

                fs.Close();

                s.Close();

                response.Close();

            }

            catch (System.Net.WebException)
            {

                if (response != null)

                    response.Close();

            }        
        
        
        
        
        }
        
        //private bool SendNotify(string _address, string _subject, string _body, string file, string _link)
        //{
        //    try
        //    {
        //        toolStripProgressBar1.Step = 1;
        //        toolStripProgressBar1.Minimum = 0;
        //        toolStripProgressBar1.Maximum = 7;
        //        toolStripProgressBar1.Value = 0;
        //        toolStripProgressBar1.Visible = true;
        //        toolStripProgressBar1.Width = statusStrip1.Width - toolStripStatusLabel1.Width - 50;

        //        // Create the Outlook application by using inline initialization.
        //        Outlook.Application oApp = new Outlook.Application();
        //        toolStripStatusLabel1.Text = "inizializzo Outlook....";
        //        toolStripProgressBar1.PerformStep();

        //        // survive to grant access....
        //        Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
        //        Outlook.MAPIFolder f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        //        System.Threading.Thread.Sleep(5000);

        //        toolStripStatusLabel1.Text = "inizializzo Messaggio....";
        //        toolStripProgressBar1.PerformStep();

        //        //Create the new message by using the simplest approach.
        //        Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

        //        //Add a recipient.
        //        // TODO: Change the following recipient where appropriate.
        //        Outlook.Recipient oRecip;
        //        if (_address != "")
        //        {
        //            oRecip = (Outlook.Recipient)oMsg.Recipients.Add(_address);
        //            oRecip.Resolve();
        //        }
        //        toolStripStatusLabel1.Text = "inizializzo Destinatario....";
        //        toolStripProgressBar1.PerformStep();

        //        //Set the basic properties.
        //        oMsg.Subject = _subject;// "This is the subject of the test message";

        //        oMsg.Body = _body; // "This is the text in the message.";

        //        if (_link != "")
        //        {
        //            oMsg.HTMLBody += "\n\r Link diretto alla pubblicazione modifica: \n\r" + string.Format("<a href=\"{0}\">{0}</a>", _link);
        //        }
        //            //toolStripStatusLabel1.Text = "inizializzo titolo e corpo maessaggio....";
        //        //toolStripProgressBar1.PerformStep();


        //        Outlook.Attachment oAttach;

        //        if (file != "")
        //        {
        //            if (File.Exists(SignFiles.startXML) == true)
        //            {
        //                String sSource = file;
        //                String sDisplayName = "Allegato Certificato";
        //                int iPosition = (int)oMsg.Body.Length + 1;
        //                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
        //                oAttach = oMsg.Attachments.Add(sSource, iAttachType, iPosition, sDisplayName);
        //            }
        //            else
        //            {
        //                // aggiungere link al PDL

        //            }
        //        }

        //        toolStripStatusLabel1.Text = "inizializzo Allegato per apertura programma firma....";
        //        toolStripProgressBar1.PerformStep();


        //        // If you want to, display the message.
        //        if (notifyPopUp == true)
        //        {
        //            oMsg.Display(true);  //modal
        //            //oMsg.Save();
        //        }
        //        else
        //        {
        //            //Send the message.
        //            oMsg.Save();
        //            oMsg.Send();
        //        }


        //        //Explicitly release objects.
        //        oRecip = null;
        //        oAttach = null;
        //        oMsg = null;
        //        oApp = null;
        //    }

        //                    // Simple error handler.
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(string.Format("Messaggio da Outlook: \r\n {0} ", e.Message), "Invio Notifica");
        //        toolStripStatusLabel1.Text = "";
        //        toolStripProgressBar1.Visible = false;
        //        return true;
                
        //    }
        //    finally
        //    {
        //        toolStripStatusLabel1.Text = "";
        //        toolStripProgressBar1.Visible = false;
        //    }
        
        //    //Default return value.
        //    return true;        
        
        
        //}


        private bool IsControlAtFront(Control control)
        {
            while (control.Parent != null)
            {
                if (control.Parent.Controls.GetChildIndex(control) == 0)
                {
                    control = control.Parent;
                    if (control.Parent == null)
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        


        private void btnSchedaPDL_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }
        

        private void SaveGridSettings(DataGridView dg)
        {
            // salva le impostazioni della gridview in un file XML per utente
            //string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            DataTable dt = new DataTable("table");

            var query = from DataGridViewColumn col in dg.Columns
                        orderby col.DisplayIndex
                        select col;

            foreach (DataGridViewColumn col in query)
            {
                dt.Columns.Add(col.Name);
            }

            //dt.WriteXmlSchema(pathGridSettings);
        }

        


        private void SaveGridSettings(Telerik.WinControls.UI.RadGridView dg)
        {
            // salva le impostazioni della gridview in un file XML per utente
            //string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            DataTable dt = new DataTable("table");

            var query = from DataGridViewColumn col in dg.Columns
                        orderby col.DisplayIndex
                        select col;

            foreach (DataGridViewColumn col in query)
            {
                dt.Columns.Add(col.Name);
            }

//            dt.WriteXmlSchema(pathGridSettings);
        }

        
        private bool cleanTempFolder(string path)
        {
            bool bOK = true;

            string[] files = Directory.GetFiles(path);

            foreach (var f in files)
            {
                try
                {
                    File.Delete(f);
                }
                catch
                {
                    bOK = false;

                }
            }

            return bOK;
        
        }

        private void btnPathEpy_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog fb = new FolderBrowserDialog();
            //fb.Description = "Apri Cartella File Shede Epy";
            
            //fb.ShowDialog();

            ////if (fb.SelectedPath.Length == 1)
            //    lblPathEpy.Text = fb.SelectedPath;

            //    //getSchede(fb.SelectedPath);

            //    radGridViewEpy.DataSource = getArticoliSchede();



            //kw.GetMyCertificates("");



        }


        private List<string> getArticoloMetodo(string codart)
        { 
            List<string> lOut = new List<string>();

            using(SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(string.Format("SELECT A.CODICE + '_' + A.DESCRIZIONE + '_' + cast(X.IDPUBBLICAZIONEKNOS as varchar) AS ARTICOLOMETODO FROM ANAGRAFICAARTICOLI A INNER JOIN ITA_ARCHIVIO_ARTICOLI X ON A.CODICE = X.CODICE WHERE A.CODICE LIKE '{0}%'", codart)))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataReader dr = cmd.ExecuteReader();

                    while (dr.Read())
                    {
                        lOut.Add(dr[0].ToString());
                    }

                
                }

            }
            return lOut;
        }
        


        private void aggiornaRegistro()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                using (SqlConnection cn = new SqlConnection(sqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(Properties.Settings.Default.SqlUpdateBSC.ToString()))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        cmd.ExecuteNonQuery();

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;

            }      
        }


        
        DataTable getGiacenzeBertello(string codart = "")
        {
            DataTable x = new DataTable();

            string sql = string.Format("SELECT * FROM ZSI_VISTA_GIACENZEBERTELLO");
            if (codart != "")
                sql += string.Format(" WHERE CODICE LIKE '{0}%'", codart.Replace("XXX","").Replace("#000", ""));

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);

                }

            }
            return x;
        }


        DataTable getSpedizionieri()
        {
            DataTable x = new DataTable();

            string strWhere = "";

            if (cmbTIPIORDINI.SelectedIndex == 1)
            {
                strWhere = " OR M12ITA = 1 ";
            }
            else
            {
                strWhere = "  OR isnull(M12ITA, 0) = isnull(M12ITA, 0) ";
            }


            string sql = string.Format("SELECT * FROM ZS_VISTA_SPEDIZIONIERI WHERE CODICE = 0 {0} ORDER BY RAGIONESOCIALE", strWhere);


            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);

                }

            }
            return x;
        }


        private void labeH_Click(object sender, EventArgs e)
        {
            
        }
        



        DataSet getArticoliSchedeCOA()
        {

            updataDatePosizionamentoITA();

            DataSet dsOut = new DataSet();
            DataTable x = new DataTable("RigheOrdine");
            DataTable xC = new DataTable("Lotti");

            string strSQL = Properties.Settings.Default.MAGMetodoCommand + " {0}";

            string strWHERE = " WHERE 1=1";

            strWHERE += String.Format(" AND DATACONSEGNA BETWEEN '{0}' AND '{1}'", dTP_BOLLEDA.Value.Date.ToString("yyyyMMdd"), dTP_BOLLEA.Value.Date.ToString("yyyyMMdd"));


            strWHERE += String.Format(" AND DATACARICO BETWEEN '{0}' AND '{1}'", dTP_COADA.Value.Date.ToString("yyyyMMdd"), dTP_COAA.Value.Date.ToString("yyyyMMdd"));

            if (chkFP.Checked)
            {
                strWHERE += String.Format(" AND DATAPOSIZIONAMENTO BETWEEN '{0}' AND '{1}'", dTP_POSIZDA.Value.Date.ToString("yyyyMMdd"), dTP_POSIZA.Value.Date.ToString("yyyyMMdd"));
            }

            Properties.Settings.Default.MAGTipiImballo = cmbIMBALLI.SelectedIndex;
            Properties.Settings.Default.MAGTipoOrdini = cmbTIPIORDINI.SelectedIndex;
            Properties.Settings.Default.MAGTStatoM05 = cmbSTATOM05.SelectedIndex;
            Properties.Settings.Default.Save();

            if (chkChiusi.Checked == false)
            {
                strWHERE += " AND QTAGESTRES > 0";
            }
            else
            {
                strWHERE += " AND QTAGESTRES >= 0";
            }

            //if (chkInviati.Checked == false)
            //{
            //    strWHERE += " AND ISNULL(DATAINVIOCOA, 1) = 1";
            //}

            //if (txtLotto.Text != "")
            //{
            //    strWHERE += string.Format(" AND LOTTO LIKE '%{0}%'", txtLotto.Text);
            //}

            
            if (cmbTIPIORDINI.SelectedIndex == 0)
            {
                if ((Properties.Settings.Default.MAGTipiDocItalia != "") && (Properties.Settings.Default.MAGTipiDocEstero != ""))
                {
                    strWHERE += string.Format(" AND TIPODOC IN ({0}, {1})", Properties.Settings.Default.MAGTipiDocItalia, Properties.Settings.Default.MAGTipiDocEstero);
                }
            }

            if (cmbTIPIORDINI.SelectedIndex == 1)
            {
                if (Properties.Settings.Default.MAGTipiDocItalia != "")
                {
                    strWHERE += string.Format(" AND TIPODOC IN ({0})", Properties.Settings.Default.MAGTipiDocItalia);
                }
            }

            if (cmbTIPIORDINI.SelectedIndex == 2)
            {
                if (Properties.Settings.Default.MAGTipiDocEstero != "")
                {
                    strWHERE += string.Format(" AND TIPODOC IN ({0})", Properties.Settings.Default.MAGTipiDocEstero);
                }
            }

            if (cmbIMBALLI.SelectedIndex <= 0)
            {
            }

            if (cmbIMBALLI.SelectedIndex == 2)
            {
                if (Properties.Settings.Default.MAGTipiImballiEsclusi != "")
                {
                    strWHERE += string.Format(" AND CODIMBALLO IN ({0})", Properties.Settings.Default.MAGTipiImballiEsclusi);
                }

            }

            if (cmbIMBALLI.SelectedIndex == 1)
            {
                if (Properties.Settings.Default.MAGTipiImballiEsclusi != "")
                {
                    strWHERE += string.Format(" AND CODIMBALLO NOT IN ({0})", Properties.Settings.Default.MAGTipiImballiEsclusi);
                }
            }


            if (cmbSTATOM05.SelectedIndex <= 0)
            {
            }

            if (cmbSTATOM05.SelectedIndex > 0)
            {
                strWHERE += string.Format(" AND MAGSTATOM05 = {0}", cmbSTATOM05.SelectedIndex - 1);
            }

            if (txtClienteCOA.Text != "")
            {
                strWHERE += string.Format(" AND RAGIONESOCIALE LIKE '%{0}%'", txtClienteCOA.Text);
            }

            strSQL = string.Format(strSQL, strWHERE) + " ORDER BY DATADOC DESC, RAGIONESOCIALE";


            



            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);


                    strSQL = "SELECT QTACONFEZIONE, LOTTO, PROGRESSIVO, IDRIGA FROM ZSI_CONFEZSPEDIBILE WHERE IDTESTADDT = 0";
                    cmd.CommandText = strSQL;
                    cmd.CommandType = CommandType.Text;
                    da = new SqlDataAdapter(cmd);

                    da.Fill(xC);

                }

            }

            dsOut.Tables.Add(x);
            dsOut.Tables.Add(xC);


            return dsOut;
        }

        void updataDatePosizionamentoITA()
        {
            string strSQL = "UPDATE e SET E.DATAPOSIZIONAMENTO = e.DATACARICO FROM RIGHEDOCUMENTI r JOIN EXTRARIGHEDOC e ON e.IDTESTA = r.IDTESTA AND e.IDRIGA = r.IDRIGA " +
                        " WHERE r.TIPODOC = 'OCC' AND R.ESERCIZIO >= 2016 AND r.CODART <> '' AND E.DATAPOSIZIONAMENTO IS NULL AND NOT e.DATACARICO IS NULL";

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    
                    cmd.ExecuteNonQuery();
                }

            }

        }


        private void btnCercaCOA_Click(object sender, EventArgs e)
        {
            int currentRowIndex = 0;

            if (this.radGridViewCOA.CurrentRow != null)
            {
                currentRowIndex = this.radGridViewCOA.CurrentRow.Index;
            }

            Cursor.Current = Cursors.WaitCursor;

            panelLOG.Left = this.Width / 2 - panelLOG.Width / 2;
            panelLOG.Top = this.Height / 2 - panelLOG.Height / 2;
            panelLOG.Visible = true;
            txtbLOGOperazioni.Text = "Operazioni in corso.....\r\n";


            comboBoxSpedizionieri.DataSource = comboBoxSpedizionieriPM.DataSource = getSpedizionieri();
            comboBoxSpedizionieri.DisplayMember = comboBoxSpedizionieriPM.DisplayMember = "RAGIONESOCIALE";
            comboBoxSpedizionieri.ValueMember = comboBoxSpedizionieriPM.ValueMember = "CODICE";

            toolStripStatusLabelLOG.Text = string.Format("{0}","Caricamento dati in corso......");

            Application.DoEvents();

            DataSet xx = getArticoliSchedeCOA();

            toolStripStatusLabelLOG.Text = string.Format("{0}","Riempio la grid");

            Application.DoEvents();

            string path = Properties.Settings.Default.PathCOA;// lblPathH.Text;

            
            DateTime datamodifica = System.DateTime.Now;

            try
            {

                if (radGridViewCOA.ChildRows.Count == 0)
                {
                    radGridViewCOA.FilterDescriptors.Clear();
                }


                radGridViewCOA.DataSource = xx.Tables[0];




                //bindingSource1.DataSource = xx;

                //radGridViewCOA.DataSource = null;
                //radGridViewCOA.Rows.Clear();
                //radGridViewCOA.DataSource = xx;
                radGridViewCOA.EnableSorting = true;
                radGridViewCOA.EnableFiltering = true;
                radGridViewCOA.ShowFilteringRow = true;
                radGridViewCOA.SortDescriptors.Add("DATADOC", ListSortDirection.Descending);

                //radGridViewCOA.SortDescriptors.Add("CLIENTE", ListSortDirection.Ascending);
                //radGridViewCOA.SortDescriptors.Add("DOCUMENTO", ListSortDirection.Ascending);

                this.radGridViewCOA.SummaryRowsTop.Clear();

                if (bolShowTotali == true)
                {

                    this.radGridViewCOA.MasterTemplate.ShowTotals = true;
                    GridViewSummaryItem summaryItem = new GridViewSummaryItem("QTAGESTRES", "{0:0,0}", GridAggregateFunction.Sum);
                    GridViewSummaryItem summaryItemFusti = new GridViewSummaryItem("FUSTI", "{0:0,0}", GridAggregateFunction.Sum);
                    GridViewSummaryItem summaryItemQtaSped = new GridViewSummaryItem("QTASPEDIBILE", "{0:0,0}", GridAggregateFunction.Sum);
                    GridViewSummaryItem summaryItemNR = new GridViewSummaryItem("DOCUMENTO", "{0:0,0}", GridAggregateFunction.Count);
                    GridViewSummaryItem summaryItemDispQta = new GridViewSummaryItem("DISPONIBILITA", "{0:0,0}", GridAggregateFunction.Sum);
                    GridViewSummaryItem summaryItemDispImb = new GridViewSummaryItem("DISPONIBILITAIMB", "{0:0,0}", GridAggregateFunction.Sum);


                    GridViewSummaryRowItem summaryRowItem = new GridViewSummaryRowItem();
                    summaryRowItem.Add(summaryItem);
                    summaryRowItem.Add(summaryItemFusti);
                    summaryRowItem.Add(summaryItemQtaSped);
                    summaryRowItem.Add(summaryItemNR);
                    summaryRowItem.Add(summaryItemDispQta);
                    summaryRowItem.Add(summaryItemDispImb);
                    this.radGridViewCOA.SummaryRowsTop.Add(summaryRowItem);

                }


                Cursor.Current = Cursors.Default;
                //Application.DoEvents();

                //this.radGridViewCOA.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.HeaderCells);
                radGridViewCOA.AutoSizeRows = false;
                radGridViewCOA.TableElement.RowHeight = hRows; // 40;

                radGridViewCOA.ShowGroupedColumns = true;


                lblNrFilesCOA.Text = "Nr Righe Ordine: " + radGridViewCOA.Rows.Count.ToString();

                if (radGridViewCOA.ChildRows.Count > 0)
                {
                   
                        lblNrFilesCOA.Text += string.Format(" filtrate {0}", radGridViewCOA.ChildRows.Count);
                }

                //MessageBox.Show("Caricamento effettuato!");

                //perform changes as refreshing/rebinding
                //this.radGridViewCOA.CurrentRow = this.radGridViewCOA.Rows[currentRowIndex];

                radGridViewCOA.ClearSelection();

                if (radGridViewCOA.Rows.Count > 0)
                {
                    if (currentRowIndex >= 0)
                    {
                        if (radGridViewCOA.ChildRows.Count > 0)
                        {
                            radGridViewCOA.ChildRows[currentRowIndex].IsSelected = true;
                            radGridViewCOA.ChildRows[currentRowIndex].IsCurrent = true;
                        }
                        else
                        {
                            radGridViewCOA.Rows[currentRowIndex].IsSelected = true;
                            radGridViewCOA.Rows[currentRowIndex].IsCurrent = true;
                        }

                        datirigaCOA(0);
                    }
                }


                //DIPENDENZE
                radGridViewCOA.MasterTemplate.Templates.Clear();
                if (radGridViewCOA.MasterTemplate.Templates.Count == 0)
                {
                    GridViewTemplate template = new GridViewTemplate();
                    template.DataSource = xx.Tables["Lotti"];
                    template.Columns["PROGRESSIVO"].IsVisible = false;
                    template.Columns["IDRIGA"].IsVisible = false;
                    template.Columns["QTACONFEZIONE"].HeaderText = "Quantità per lotto";
                    template.Columns["QTACONFEZIONE"].FormatString = "{0:0,0.00}";
                    radGridViewCOA.MasterTemplate.Templates.Add(template);

                    GridViewRelation relation = new GridViewRelation(radGridViewCOA.MasterTemplate);
                    relation.ChildTemplate = template;
                    relation.RelationName = "Lotti";
                    relation.ParentColumnNames.Add("IDTESTA");
                    relation.ChildColumnNames.Add("PROGRESSIVO");
                    relation.ParentColumnNames.Add("IDRIGA");
                    relation.ChildColumnNames.Add("IDRIGA");
                    radGridViewCOA.Relations.Add(relation);
                }

                toolStripStatusLabelLOG.Text = string.Format("{0}", "Caricamento effettuato!");

                panelLOG.Visible = false;

                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                panelLOG.Visible = false;
            }

            // calendario

            //calendarUserControl1.mvStartDate = dTP_BOLLEDA.Value;
            //calendarUserControl1.mvEndDate = dTP_BOLLEA.Value;
            //calendarUserControl1.SetMonthViewStartEnd();


        }


        void radGridViewCOA_CellFormatting2(object sender, Telerik.WinControls.UI.CellFormattingEventArgs e)
        {
            Font x = new Font("Arial",12,FontStyle.Bold);



            if ((e.ColumnIndex >= 0) && (e.RowIndex >= 0))
            {



                e.CellElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local);
                e.CellElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local);
                e.CellElement.ResetValue(LightVisualElement.NumberOfColorsProperty, ValueResetFlags.Local);
                e.CellElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local);

                if (e.CellElement.ColumnInfo.HeaderText == "I")
                {
                    if (e.CellElement.RowInfo.Cells["ROWHEADERIMAGE"].Value != null)
                        Debug.Print(e.CellElement.RowInfo.Cells["ROWHEADERIMAGE"].Value.ToString());
                    if (Properties.Settings.Default.RowHeaderImageList.ToString() == "16")
                    {
                        e.CellElement.Image = imageList16.Images[e.CellElement.RowInfo.Cells["ROWHEADERIMAGE"].Value.ToString()];
                    }
                    else
                    {
                        e.CellElement.Image = imageList32.Images[e.CellElement.RowInfo.Cells["ROWHEADERIMAGE"].Value.ToString()];
                    }

                }

                if (e.CellElement.ColumnInfo.HeaderText == "IM")
                {
                    if (e.CellElement.RowInfo.Cells["ROWHEADERIMAGEM"].Value != null)
                        Debug.Print(e.CellElement.RowInfo.Cells["ROWHEADERIMAGEM"].Value.ToString());
                    if (Properties.Settings.Default.RowHeaderImageList.ToString() == "16")
                    {
                        e.CellElement.Image = imageList16.Images[e.CellElement.RowInfo.Cells["ROWHEADERIMAGEM"].Value.ToString()];
                    }
                    else
                    {
                        e.CellElement.Image = imageList32.Images[e.CellElement.RowInfo.Cells["ROWHEADERIMAGEM"].Value.ToString()];
                    }

                }

                if (e.CellElement.ColumnInfo.HeaderText.ToUpper().StartsWith("PRODOTTO"))
                {

                    e.CellElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local);
                    e.CellElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local);
                    e.CellElement.ResetValue(LightVisualElement.NumberOfColorsProperty, ValueResetFlags.Local);
                    e.CellElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local);


                    if (e.CellElement.RowInfo.Cells["DSCMAGSTATOM05"].Value.ToString() == "PREVISIONE")
                    {
                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Black;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Yellow;

                    }

                    if (e.CellElement.RowInfo.Cells["DSCMAGSTATOM05"].Value.ToString() == "POSIZIONATO")
                    {

                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Black;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Lime;
                            
                    }

                }


                if (e.CellElement.ColumnInfo.HeaderText.ToUpper() == "ORDINE")
                {
                    int zc = 0;

                    int.TryParse(e.CellElement.RowInfo.Cells["ZCOUNT"].Value.ToString(), out zc);

                    if (zc > 0)
                    {
                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Black;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Yellow;
                    }
                    else
                    {
                        e.CellElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local);
                        e.CellElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local);
                        e.CellElement.ResetValue(LightVisualElement.NumberOfColorsProperty, ValueResetFlags.Local);
                        e.CellElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local);
                    }
                }


                if (e.CellElement.ColumnInfo.HeaderText == "gg_dataconsegna")
                {
                    int gg = 0;

                    int.TryParse(e.CellElement.RowInfo.Cells["gg_dataconsegna"].Value.ToString(), out gg);

                    if (gg < 0)
                    {
                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Blue;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Pink;
                    }

                    if (gg <= -7)
                    {
                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Blue;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Red;
                    }

                    if (gg >= 0)
                    {
                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Yellow;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Lime;
                    }

                    if (gg >= 7)
                    {
                        e.CellElement.DrawFill = true;
                        e.CellElement.ForeColor = Color.Yellow;
                        e.CellElement.NumberOfColors = 1;
                        e.CellElement.BackColor = Color.Green;
                    }

                }

                if (e.Row is GridViewDataRowInfo)
                {
                    e.CellElement.ToolTipText = "" + e.CellElement.Text;
                }
            }
        }

        private void btnSendMailCOA_Click(object sender, EventArgs e)
        {
            string address = Properties.Settings.Default.sendMailBCCSimulazioneCOA; //;kavanzi@italcom.biz";
            string addressCC = "";  //Properties.Settings.Default.sendMailBCCSimulazione; //"alfredo.deangelo@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string addressBCC = Properties.Settings.Default.sendMailBCCCOA; // "knosmail@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string body = "";
            string subject = "";
            string dettaglioM05 = "";

            //bool bOKUpload = true;

            //string codclifor = "";
            //string codart = "";


            int IdObjectDOC = 0;

            //string localfilenameCOA = "";
            string localfileCOA = "";
            string subjectM05 = Properties.Settings.Default.sendMailM05Subject;


            string msg = "";

            //int IdObjectSentMail = 0;



            //Knos
            //if (refreshKnosLogin() == false)
            //    return;

            if (chkSimulazioneCOA.Checked == false)
                address = "";

            foreach (string c in checkedListBoxMailTo.CheckedItems)
            {
                address += c + ";";
            }

            toolStripProgressBarLOG.Minimum = 0;
            toolStripProgressBarLOG.Maximum = radGridViewCOA.SelectedRows.Count + 1;
            toolStripProgressBarLOG.Value = 1;
            toolStripProgressBarLOG.Step = 1;
            toolStripProgressBarLOG.Visible = true;

            log.LogSomething(string.Format("Nr mail da inviare: {0}", radGridViewCOA.SelectedRows.Count));

            //checkBoxInterrompiInvio.Enabled = true;

            try
            {

                dettaglioM05 += "<table border=\"1\"  cellspacing=\"0\" cellpadding=\"2\">";
                dettaglioM05 += string.Format("<tr>" +
                    "<th width=\"110\"><b><font size=\"2\"  face=\"Arial\">{0}</font></b></th>" +
                    "<th width=\"240\"><b><font size=\"2\"  face=\"Arial\">{1}</font></b></th>" +
                    "<th width=\"120\"><b><font size=\"2\"  face=\"Arial\">{2}</font></b></th>" +
                    "<th width=\"250\"><b><font size=\"2\"  face=\"Arial\">{3}</font></b></th>" +
                    "<th width=\"100\"><b><font size=\"2\"  face=\"Arial\">{4}</font></b></th>" +
                    "<th width=\"140\"><b><font size=\"2\"  face=\"Arial\">{5}</font></b></th>" +
                    "<th width=\"90\"><b><font size=\"2\"  face=\"Arial\">{6}</font></b></th>" +
                    "<th width=\"90\"><b><font size=\"2\"  face=\"Arial\">{7}</font></b></th>" +
                    "<th width=\"90\"><b><font size=\"2\"  face=\"Arial\">{8}</font></b></th>" +
                    "<th width=\"90\"><b><font size=\"2\"  face=\"Arial\">{9}</font></b></th>" +
                    "</tr>"
                    , "ARTICOLO\r\nITEM"
                    , "DESCRIZIONE\r\nDESCRIPTION"
                    , "DOCUMENTO\r\nORDER NO."
                    , "RAGIONE SOCIALE\r\nCUSTOMER"
                    , "DATA POS./CARICO\r\nEXPECTED LOADING DATE"
                    , "IMBALLO\r\nPACKING"
                    , "Q.TA' ORDINE\r\nQ.TY"
                    , "Q.TA' SPEDIBILE\r\nAVALAIBLE Q.TY"
                    , "LOTTO\r\nBATCH"
                    , "DISPONIBILITA'\r\n"
                    );

                // preparazione del body
                if (radGridViewCOA.SelectedRows.Count > 0)
                {
                    for (int r = 0; r < radGridViewCOA.SelectedRows.Count; r++)
                    {


                        string dtPosizionamento = "NON POSIZIONATO";

                        if (radGridViewCOA.SelectedRows[r].Cells["DATAPOSIZIONAMENTO"].Value.ToString().Length >= 10)
                        {
                            dtPosizionamento = radGridViewCOA.SelectedRows[r].Cells["DATAPOSIZIONAMENTO"].Value.ToString().Substring(0, 10);
                        }
                        else
                        {
                            if (radGridViewCOA.SelectedRows[r].Cells["DATACARICO"].Value.ToString().Length >= 10)
                            {
                                dtPosizionamento += "\r\n Carico: " + radGridViewCOA.SelectedRows[r].Cells["DATACARICO"].Value.ToString().Substring(0, 10);
                            }
                        }

                        dettaglioM05 += string.Format("<tr><td><font size=\"2\"  face=\"Arial\">{0}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{1}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{2}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{3}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{4}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{5}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{6}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{7}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{8}</font></td>" +
                            "<td><font size=\"2\"  face=\"Arial\">{9}</font></td>" +
                            "</tr>"
                            , radGridViewCOA.SelectedRows[r].Cells["ARTICOLO"].Value.ToString()
                            , radGridViewCOA.SelectedRows[r].Cells["DESCRIZIONE"].Value.ToString()
                            , string.Format("{0} (Rif. {1})", radGridViewCOA.SelectedRows[r].Cells["DOCUMENTO"].Value.ToString(), radGridViewCOA.SelectedRows[r].Cells["NUMRIFDOC"].Value.ToString())
                            , string.Format("{0} {1}", radGridViewCOA.SelectedRows[r].Cells["RAGIONESOCIALE"].Value.ToString(), radGridViewCOA.SelectedRows[r].Cells["DESTINAZIONEMAILCOMM"].Value.ToString())
                            , dtPosizionamento
                            , radGridViewCOA.SelectedRows[r].Cells["IMBALLO"].Value.ToString()
                            , String.Format("{0:0.##}", radGridViewCOA.SelectedRows[r].Cells["QTAGESTRES"].Value.ToString())
                            , String.Format("{0:0.##}", radGridViewCOA.SelectedRows[r].Cells["QTASPEDIBILECOMM"].Value.ToString())
                            , radGridViewCOA.SelectedRows[r].Cells["XLS_NRLOTTO"].Value.ToString()
                            , radGridViewCOA.SelectedRows[r].Cells["XLS_M05_ESTERO_DISPONIBILITA"].Value.ToString()
                            );

                    }
                    dettaglioM05 += "</table>";
                }
                else
                {
                    MessageBox.Show("Nessuna riga selezionata");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Verificare le date di posizionamento delle righe selezionate");
                return;
            }


            msg = string.Format("Procedo con l'invio delle notifiche {0}", radGridViewCOA.SelectedRows.Count);


            if (MessageBox.Show(msg, "Invio mail ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {

                toolStripProgressBarLOG.Value += 1;
                toolStripProgressBarLOG.Text = string.Format("Record {0}/{1}", toolStripProgressBarLOG.Value, radGridViewCOA.SelectedRows.Count);
                log.LogSomething(string.Format("Record {0}/{1}", toolStripProgressBarLOG.Value, radGridViewCOA.SelectedRows.Count));

                var attachments = new List<string>();

                localfileCOA = "";
                IdObjectDOC = 0;

                log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));


                // invio singolo
                Application.DoEvents();

                subjectM05 = string.Format(Properties.Settings.Default.sendMailM05Subject);

                body = string.Format(Properties.Settings.Default.sendMailM05, string.Format("<p>{0}</p><br/>{1}", txtTestoMail.Text, dettaglioM05));

                log.LogSomething(string.Format("Invio mail {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));


                if (Notifica.SendNotifyOtlk(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                {
                    log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                    textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);

                    radGridViewCOA.ShowRowHeaderColumn = true;

                    Application.DoEvents();
                }
                else
                {
                    log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                    textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);
                    //radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                }
                //if (Properties.Settings.Default.UseLotus)
                //{

                //    subject = string.Format("{0}", subjectM05);


                //    Notifica cNotifica = new ToDoNotificheBSC.Notifica();
                //    if (cNotifica.SendNotifyLotus(address, subject, body, attachments, null, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                //    {
                //        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                //        textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);

                //        radGridViewCOA.ShowRowHeaderColumn = true;

                //        Application.DoEvents();
                //    }
                //    else
                //    {
                //        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                //        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);
                //        //radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                //    }


                //}
                //else
                //{
                //    if (Properties.Settings.Default.UseCdo)
                //    {
                //        subject = string.Format("{0}", subjectM05);

                //        Notifica cNotifica = new ToDoNotificheBSC.Notifica();
                //        if (cNotifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                //        {
                //            log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                //            textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);

                //            radGridViewCOA.ShowRowHeaderColumn = true;

                //            Application.DoEvents();
                //        }
                //        else
                //        {
                //            log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                //            textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);
                //            //radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                //        }
                //    }
                //    else
                //    {
                //        if (Notifica.SendNotifyMAPI(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                //        //if (Notifica.SendNotifyMAPILotus(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                //        {
                //            radGridViewCOA.ShowRowHeaderColumn = true;

                //            Application.DoEvents();
                //            log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                //            textBoxLOG.Text += string.Format("\r\n - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC);

                //        }
                //        else
                //        {
                //            log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                //            textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, "", addressCC, addressBCC);
                //        }
                //    }
                //}

            }

            toolStripProgressBarLOG.Visible = false;
            //checkBoxInterrompiInvio.Enabled = false;

            MessageBox.Show("Invio completato!");
        }



        private void updateDatiRiga(int idt, int idr, string user, string machine)
        {
            if (CurrentUser != "")
            {
                user = CurrentUser;
            }

            if (idtesta == 0)
            {
                MessageBox.Show("Selezionare una riga", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }


            string dt = System.DateTime.Today.ToShortDateString();

            double qtaspedibile = 0;
            double.TryParse(txtbQTASPEDIBILE.Text, out qtaspedibile);

            double disponibilita = 0;
            double.TryParse(txtbDISPONIBILITA.Text, out disponibilita);

            double coddestdiv = 0;
            double.TryParse(lblNUMDESTDIVERSAMERCI.Text, out coddestdiv);

            //string strSQL = string.Format("UPDATE EXTRARIGHEDOC SET DATAMODIFICA = getdate(), UTENTEMODIFICA = 'ITALCOM', PRZCARICO = {2}, VALFORNITURA = {3}, NOTEMAG = '{4}' WHERE IDTESTA = {0} AND IDRIGA = {1}"
            //    , idt, idr
            //    , qtaspedibile
            //    , chkCONFEZIONATO.SelectedIndex
            //    , txtbNOTEMAG.Text
            //    );

            string strSQL = string.Format("ITA_SP_UPDATE_DATISPEDIZIONE");

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("idtesta", idtesta);
                    cmd.Parameters.Add("idriga", idriga);
                    cmd.Parameters.Add("codsped", decimal.Parse(comboBoxSpedizionieri.SelectedValue.ToString()));
                    cmd.Parameters.Add("nrlotto", txtbLotto.Text);
                    cmd.Parameters.Add("qtaspedibile", qtaspedibile);
                    cmd.Parameters.Add("disponibilita", disponibilita);
                    cmd.Parameters.Add("posizionamento", txtbPOSIZIONAMENTO.Text);
                    cmd.Parameters.Add("magstatoriga", chkCONFEZIONATO.SelectedIndex);
                    cmd.Parameters.Add("notecli", txtbNOTECLIENTE.Text);
                    cmd.Parameters.Add("noteart", txtbNOTEARTICOLO.Text);
                    cmd.Parameters.Add("notemag", txtbNOTEMAG.Text);
                    cmd.Parameters.Add("noteconsignee", txtbNOTECONSIGNEE.Text);
                    cmd.Parameters.Add("notenotify", txtbNOTENOTIFY.Text);
                    cmd.Parameters.Add("annotazioni", txtbNOTERIGA.Text);
                    cmd.Parameters.Add("coddestdiv", coddestdiv);
                    cmd.Parameters.Add("notecontainer", txtbNOTECONTAINER.Text);
                    cmd.Parameters.Add("datacarico", dTPDataCarico.Value);
                    cmd.Parameters.Add("dataconsegna", dTPDataConsegna.Value);

                    if (chkMAGSTATOM05.SelectedIndex < 2)
                    {
                        if (MessageBox.Show("Attenzione, devo aggiornare la data posizionamento?", "Controllo data posizionamento", MessageBoxButtons.YesNo, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                        {
                            cmd.Parameters.Add("dataposizionamento", dTPPosizionamento.Value);
                            chkMAGSTATOM05.SelectedIndex = 2;
                        }
                        else
                        {
                            cmd.Parameters.Add("dataposizionamento", DBNull.Value);
                            chkMAGSTATOM05.SelectedIndex = 1;
                        }


                    }
                    else
                    {
                        cmd.Parameters.Add("dataposizionamento", dTPPosizionamento.Value);

                    }


                    cmd.Parameters.Add("magstatom05", chkMAGSTATOM05.SelectedIndex);
                    cmd.Parameters.Add("user", user);
                    cmd.Parameters.Add("machine", machine);
                    cmd.Parameters.Add("m_mezzo", txtbNOTECLIENTEMEZZO.Text);
                    cmd.Parameters.Add("m_costo", txtbNOTECLIENTECOSTO.Text);
                    cmd.Parameters.Add("noteshipper", txtbNOTESHIPPER.Text);
                    cmd.Parameters.Add("notedocumenti", txtbNOTEDOCUMENTI.Text);
                    cmd.Parameters.Add("noteddc", txtbNOTEDDC.Text);
                    cmd.Parameters.Add("raggruppamento", txtbRAGGRUPPAMENTO.Text);



                    try
                    {

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("{0}", ex.Message), "Errore");
                    }
                }

            }
        }


        private void radGridViewCOA_CurrentRowChanging(object sender, Telerik.WinControls.UI.CurrentRowChangingEventArgs e)
        {
            if (radGridViewCOA.Rows.Count > 0)
            {
                if ( (e.NewRow != null))
                {
                    if (e.NewRow.Index >= 0) 
                    datirigaCOA(0);
                }
            }
        }


        private string datirigaCOA(int r)
        {
            try
            {

                if (radGridViewCOA.CurrentCell == null)
                    return ("");

                if ((radGridViewCOA.CurrentCell.ColumnIndex < 0) || (radGridViewCOA.CurrentCell.RowIndex < 0))
                {
                    //MessageBox.Show("Selezionare una riga!");
                    return("");
                }

                dGVConfezionamentoLotti.DataSource = getConfezionamentoLotti(0, 0);
                panelConfezionamentoLotti.Visible = false;
                panelConfezionamentoLotti.Visible = false;


                idtesta = idriga = 0;
                nrpezziimballo = 1;
                articolo = "";


                int.TryParse(radGridViewCOA.CurrentCell.RowInfo.Cells["NRPEZZIIMBALLO"].Value.ToString(), out nrpezziimballo);

                string tooltip = "Cliente: {0} \r\n" +
                    "Articolo: {1} - " +
                    "Data Consegna: {2} - " +
                    "Quantità: {3} \r\n" +
                    "Destinazione: {4} \r\n" +
                    "Spedibile: {5} \r\n" +
                    "Documento: {6} {9:MM/dd/yyyy}\r\n\r\n" +
                    "Note Cliente: {7} \r\n\r\n" +
                    "Note Articolo: {8} \r\n\r\n" +
                    "Note Magazzino: {9} \r\n\r\n" +
                    "Ultima Modifica: {10} \r\n\r\n";

                string cliente = string.Format("{0} - {1}", radGridViewCOA.CurrentCell.RowInfo.Cells["RAGIONESOCIALE"].Value.ToString(), radGridViewCOA.CurrentCell.RowInfo.Cells["CODCLIFOR"].Value.ToString());
                articolo = radGridViewCOA.CurrentCell.RowInfo.Cells["ARTICOLO"].Value.ToString();
                string datadoc = radGridViewCOA.CurrentCell.RowInfo.Cells["DATADOC"].Value.ToString();
                string dataconsegna = radGridViewCOA.CurrentCell.RowInfo.Cells["DATACONSEGNA"].Value.ToString();
                string datacarico = radGridViewCOA.CurrentCell.RowInfo.Cells["DATACARICO"].Value.ToString();

                if (radGridViewCOA.CurrentCell.RowInfo.Cells["DATACONSEGNA"].Value != null)
                    dTPDataConsegna.Value = DateTime.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["DATACONSEGNA"].Value.ToString());

                if (radGridViewCOA.CurrentCell.RowInfo.Cells["DATACARICO"].Value != null)
                    dTPDataCarico.Value = DateTime.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["DATACARICO"].Value.ToString());

                if ((radGridViewCOA.CurrentCell.RowInfo.Cells["DATAPOSIZIONAMENTO"].Value != null) && (radGridViewCOA.CurrentCell.RowInfo.Cells["DATAPOSIZIONAMENTO"].Value.ToString() != ""))
                {
                    dTPPosizionamento.Value = DateTime.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["DATAPOSIZIONAMENTO"].Value.ToString());
                }
                else
                {
                    dTPPosizionamento.Value = dTPDataCarico.Value;
                }

                lblDataCarico.BackColor = lblDataConsegna.BackColor;
                if (dTPDataCarico.Value == dTPDataConsegna.Value)
                {
                    lblDataCarico.BackColor = Color.Red;
                }

                string qta = radGridViewCOA.CurrentCell.RowInfo.Cells["QTAGESTRES"].Value.ToString();
                string qtaspedibile = radGridViewCOA.CurrentCell.RowInfo.Cells["QTASPEDIBILE"].Value.ToString();
                string documento = radGridViewCOA.CurrentCell.RowInfo.Cells["DOCUMENTO"].Value.ToString();
                string notecli = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTECLI"].Value.ToString();
                string noteart = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTEART"].Value.ToString();
                string notemag = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTEMAG"].Value.ToString();

                txtbCodice.Text = articolo;
                txtbDescrizione.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["DESCRIZIONE"].Value.ToString();
                txtbQTAGESTRES.Text = qta;
                txtbQTASPEDIBILE.Text = qtaspedibile;
                txtbDISPONIBILITA.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["DISPONIBILITA"].Value.ToString();

                try
                {
                    double disp = 0;
                    double.TryParse(txtbDISPONIBILITA.Text, out disp);
                    txtbDISPONIBILITA_IMB.Text = (disp / nrpezziimballo).ToString();
                }
                catch (Exception ex)
                { }


                txtbIMBALLO.Text = string.Format("{0} {1}", radGridViewCOA.CurrentCell.RowInfo.Cells["CODIMBALLO"].Value.ToString(), radGridViewCOA.CurrentCell.RowInfo.Cells["IMBALLO"].Value.ToString());
                txtbNRFUSTI.Text = string.Format("{0}", radGridViewCOA.CurrentCell.RowInfo.Cells["FUSTI"].Value.ToString());
                chkCONFEZIONATO.SelectedIndex = int.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["MAGSTATORIGA"].Value.ToString());
                chkMAGSTATOM05.SelectedIndex =  int.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["MAGSTATOM05"].Value.ToString());
                txtbLotto.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["NRLOTTO"].Value.ToString();

                chkCONFEZIONATO.SetItemChecked(chkCONFEZIONATO.SelectedIndex, true);
                chkMAGSTATOM05.SetItemChecked(chkMAGSTATOM05.SelectedIndex, true);
                //chkMAGSTATOM05.SetItemChecked(chkMAGSTATOM05.SelectedIndex, true);

                

                //chkMAGSTATOM05.UncheckAllItems();
                //chkMAGSTATOM05.SelectedItem.CheckState = Telerik.WinControls.Enumerations.ToggleState.On;// = int.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["MAGSTATORIGA"].Value.ToString()); 
                chkMAGSTATOM05.SelectedIndex = int.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["MAGSTATOM05"].Value.ToString());
                //chkMAGSTATOM05_SelectedIndexChanged(null, null);


                txtbNOTEMAG.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTEMAG"].Value.ToString();
                txtbPOSIZIONAMENTO.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["POSIZIONAMENTO"].Value.ToString();

                comboBoxSpedizionieri.SelectedValue = int.Parse(radGridViewCOA.CurrentCell.RowInfo.Cells["CODSPED"].Value.ToString());

                txtbDATACARICO.Text = datacarico;
                txtbDATACONSEGNA.Text = dataconsegna;

                txtbNOTECLIENTE.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTECLI"].Value.ToString();
                txtbNOTECLIENTEMEZZO.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["M_MEZZO"].Value.ToString();
                txtbNOTECLIENTECOSTO.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["M_COSTO"].Value.ToString();
                txtbNOTEARTICOLO.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTEART"].Value.ToString();
                txtbNOTERIGA.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["ANNOTAZIONI"].Value.ToString();
                txtbNOTECONSIGNEE.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["XLS_NOTECONSIGNEE"].Value.ToString();
                txtbNOTENOTIFY.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["XLS_NOTENOTIFY"].Value.ToString();

                txtbNOTESHIPPER.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["XLS_NOTESHIPPER"].Value.ToString();
                txtbNOTEDOCUMENTI.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["XLS_NOTEDOCUMENTI"].Value.ToString();
                txtbNOTECFINDOC.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["XLS_NOTECFINDOC"].Value.ToString();

                int coddestdiv = 0;

                int.TryParse(radGridViewCOA.CurrentCell.RowInfo.Cells["NUMDESTDIVERSAMERCI"].Value.ToString(), out coddestdiv);


                txtbNOTEDDC.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["XLS_NOTEDESTINAZIONEDIVERSA"].Value.ToString();

                txtbNOTEDDC.Enabled = (coddestdiv > 0);



                txtbNOTECONTAINER.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["NOTECONTAINER"].Value.ToString();
                int.TryParse(radGridViewCOA.CurrentCell.RowInfo.Cells["IDTESTA"].Value.ToString(), out idtesta);
                int.TryParse(radGridViewCOA.CurrentCell.RowInfo.Cells["IDRIGA"].Value.ToString(), out idriga);

                lblNUMDESTDIVERSAMERCI.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["NUMDESTDIVERSAMERCI"].Value.ToString();
                lblDOCUMENTO.Text = string.Format("{0} {1:MM/dd/yyyy}", documento, datadoc);

                //string address = "";// radGridViewCOA.Rows[r].Cells["EMAIL_CLIENTE"].Value.ToString();
                //string addressCC = ""; //radGridViewCOA.Rows[r].Cells["DOCUMENTO"].Value.ToString();

                //string fileSCH = radGridViewCOA.Rows[r].Cells["FILENAME"].Value.ToString();
                //string ultimoinvioSCH = "";
                //if (radGridViewCOA.Rows[r].Cells["DATAINVIOCOA"].Value != null)
                //    ultimoinvioSCH = radGridViewCOA.Rows[r].Cells["DATAINVIOCOA"].Value.ToString();

                string destinazione = radGridViewCOA.CurrentCell.RowInfo.Cells["DESTINAZIONE"].Value.ToString();

                textBoxToolTipCOA.Text = string.Format(tooltip, cliente, articolo, dataconsegna, qta, destinazione, qtaspedibile + " " + chkCONFEZIONATO.SelectedItem.ToString(), documento, notecli, noteart, notemag, datadoc, radGridViewCOA.CurrentCell.RowInfo.Cells["INFOMODIFICA"].Value.ToString());

                toolTip1.SetToolTip(textBoxToolTipCOA, textBoxToolTipCOA.Text);

                //lotti
                cmbLottiODP.DataSource = getOrdiniProduzioneLOTTI(articolo);
                cmbLottiODP.DisplayMember = "NRLOTTO";
                cmbLottiODP.ValueMember = "NRLOTTO";

                pbROWHEADERIMAGE.Image = imageList16.Images[radGridViewCOA.CurrentCell.RowInfo.Cells["ROWHEADERIMAGE"].Value.ToString()];

                txtbRAGGRUPPAMENTO.Text = radGridViewCOA.CurrentCell.RowInfo.Cells["RAGGRUPPAMENTO"].Value.ToString();

                return string.Format(textBoxToolTipCOA.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return ex.Message; //""


            }
        }

        private void btnAllegati_Click(object sender, EventArgs e)
        {
            List<Allegato> allegati = new List<Allegato>();
            frmUpload f = new frmUpload();
            
            f.ShowDialog();

            foreach (DataGridViewRow r in f.dataGridView1.Rows)
            {
                if (r.Cells[0].Value != null)
                {
                    Allegato a = new Allegato(r.Cells[0].Value.ToString(), r.Cells[0].Value.ToString(), r.Cells[2].Value.ToString());

                    allegati.Add(a);
                }

            }

            creaAllegatiKnos(allegati);

        }

        void creaAllegatiKnos(List<Allegato> listaallegati)
        { 
        
            string commandtext = "INSERT INTO KNOS_ZSI.dbo.Object_Link (IdObjectFrom, IdObjectTo, Url, ExternalLink, Pos, LinkDescr)  SELECT {0}, 0, '{1}', '1', 0, '{2}' FROM TABDITTE WHERE NOT EXISTS (SELECT 1 FROM KNOS_ZSI.dbo.Object_Link WHERE IDOBJECTFROM = {0} AND URL = '{1}')";

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {

                try
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand(commandtext, cn))
                    {

                        //for (int i = 0; i < radGridViewAllegatiSS.SelectedRows.Count; i++)
                        //{
                        //    //if (radGridViewAllegatiSS.SelectedRows[i].Cells[0].Value.ToString() != null)
                        //    //{
                        //    //    foreach(Allegato a in listaallegati)
                        //    //    {
                                    
                        //    //        Application.DoEvents();

                        //    //        toolStripStatusLabel1.Text = string.Format("Allego file: {0} alla pubblicazione con IdObject: {1}", a.Path, radGridViewAllegatiSS.SelectedRows[i].Cells["IDOBJECT_SCH"].Value.ToString());

                        //    //        cmd.CommandText = string.Format(commandtext, radGridViewAllegatiSS.SelectedRows[i].Cells["IDOBJECT_SCH"].Value.ToString(), a.Path, a.Descrizione);
                        //    //        cmd.ExecuteNonQuery();
                        //    //    }
                        //    //}

                        //}

                        toolStripStatusLabel1.Text = string.Format("Caricamento completato");
                    }

                }

                catch (SqlException ex)
                {
                    MessageBox.Show(string.Format("Errore SQL SERVER: {0} - {1}", sqlConnectionString, ex.Message));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore : {0}", ex.Message));
                }

            }
        
        }


        private void btnCaricaSchedeAllegati_Click(object sender, EventArgs e)
        {
            string strW = "";


            string commandtext = Properties.Settings.Default.MetodoCommandAllegati;

            //commandtext += " WHERE DATADOC >= @DATADOC";

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {

                try
                {
                    cn.Open();

                    toolStripStatusLabel1.Text = string.Format("caricamento dati in corso..........");

                    //radGridViewAllegatiSS.EnableFiltering = false;
                    //radGridViewAllegatiSS.ShowFilteringRow = false;

                    //using (SqlCommand cmd = new SqlCommand(commandtext, cn))
                    //{

                    //    //cmd.Parameters.AddWithValue("DATADOC", dateTimePickerDa.Value);
                    //    //cmd.ExecuteNonQuery();
                    //    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    //    DataTable dt = new DataTable();
                    //    da.Fill(dt);

                    //    radGridViewAllegatiSS.DataSource = dt;

                    //    for (int i = 0; i < radGridViewAllegatiSS.Columns.Count; i++)
                    //    {
                    //        if (radGridViewAllegatiSS.Columns[i].FieldName.StartsWith("IDOBJECT"))
                    //        {
                    //            radGridViewAllegatiSS.Columns[i].IsVisible = false;
                    //        }
                    //        else
                    //        {
                    //            radGridViewAllegatiSS.Columns[i].BestFit();
                    //        }
                    //    }

                    //    radGridViewAllegatiSS.AutoScroll = true;
                    //    radGridViewAllegatiSS.Refresh();

                    //    toolStripStatusLabel1.Text = string.Format("Caricamento completato");

                    //}


                    //radGridViewAllegatiSS.EnableFiltering = true;
                    //radGridViewAllegatiSS.ShowFilteringRow = true;
                    //radGridViewAllegatiSS.EnableAlternatingRowColor = true;
                    //radGridViewAllegatiSS.MultiSelect = true;

                }

                catch (SqlException ex)
                {
                    MessageBox.Show(string.Format("Errore SQL SERVER: {0} - {1}", sqlConnectionString, ex.Message));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore : {0}", ex.Message));
                }

            }
        }

        private void frmToDoNotificheBSC_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);

                string p = fi.FullName;

                if (!p.StartsWith("\\"))
                {
                    p = "file://" + p;
                }

                //dataGridView1.Rows.Add(fi.Name, fi.Name, p);
            }
        }

        private void frmToDoNotificheBSC_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }



        private void labelUpload_DoubleClick(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", @"H:\COMUNICAZIONI");

        }

        private void chkAllegati_Click(object sender, EventArgs e)
        {
            //panel1.Visible = chkAllegati.Checked;
        }

        private void radGridViewCOA_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            int IdObjectBOL = 0;
            int IdDocBOL = 0;
            int IdObjectSCH = 0;
            int IdDocSCH = 0;


            ////http://vsrv2k8bsn2:8780/KnoS_Catalog/0/0000035964/0001/1426089157015/Zetesol%20MGS.doc
            //string url = "{0}";

            //if (e.ColumnIndex >= 0 && e.ColumnIndex >= 0)
            //{
            //    if (radGridViewCOA.Rows[e.RowIndex].Cells["filename"].ColumnInfo.Index == e.ColumnIndex)
            //    {
            //        url = string.Format(url, radGridViewCOA.Rows[e.RowIndex].Cells["filename"].Value.ToString());
            //        //url = url.Replace("#", "_");

            //        webBrowserCOA.Navigate(url);
            //    }


            //}
            //else
            //{
            //    // selezione celle
            //    Debug.Print(string.Format("ci: {0}  - ri: {1}", e.ColumnIndex, e.RowIndex));

            //    if ((e.RowIndex == -1) && (e.ColumnIndex == -1))
            //    {
            //        radGridView1.SelectAll();
            //    }

            //    if ((e.RowIndex > -1) && (e.ColumnIndex == -1))
            //    {

            //        if (e.Row.Group != null)
            //        {
            //            for (int x = 0; x < e.Row.Parent.ChildRows.Count; x++)
            //            {

            //                e.Row.Parent.ChildRows[x].IsSelected = true;
            //                Application.DoEvents();

            //            }


            //        }

            //    }




            //}
        }


        private bool refreshKnosLogin()
        {

            //Knos
            if (kw.Inizializza(Properties.Settings.Default.KnoS_URL) == true)
            {

                txtKnosUrl.Text = Properties.Settings.Default.KnoS_URL;
                txtKnoSUser.Text = kw.CurrentUser;

                Application.DoEvents();


                txtKnosUrl.ReadOnly = txtKnoSUser.ReadOnly = txtKnoSPassword.ReadOnly = true;
                btnKnoSLogin.Enabled = false;

                statusStrip1.Text = string.Format("");
                return true;
            }
            else
            {
                txtKnosUrl.Text = Properties.Settings.Default.KnoS_URL;
                statusStrip1.Text =string.Format("Sito KnoS {0} non trovato o non accessibile!", Properties.Settings.Default.KnoS_URL);
                btnKnos.BackColor = Color.Red;
                return false;
            }

            return true;
        }

        private void btnKnoSLogin_Click(object sender, EventArgs e)
        {
            kw.CurrentUser = txtKnoSUser.Text;
            kw.PWD = txtKnoSPassword.Text;
            if (kw.Inizializza(txtKnosUrl.Text) == true)
            {
                btnKnoSLogin.BackColor = Color.LightGreen;
                Properties.Settings.Default.KnoS_URL = txtKnosUrl.Text;
                Properties.Settings.Default.KnoS_User = txtKnoSUser.Text;
                Properties.Settings.Default.KnoS_PWD = txtKnoSPassword.Text;
                Properties.Settings.Default.Save();

            }

        }

        private void frmMagazzino_ResizeEnd(object sender, EventArgs e)
        {
            splitContainer2.SplitterDistance = splitContainer2.Width - 400;
        }

        private void radGridViewCOA_Click(object sender, EventArgs e)
        {
            datirigaCOA(0);
        }

        private void txtbQTAGESTRES_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSalvaDati_Click(object sender, EventArgs e)
        {
            //    int idtesta = 0;
            //    int idriga = 0;

            //    int.TryParse(radGridViewCOA.CurrentRow.Cells["IDTESTA"].Value.ToString(), out idtesta);
            //    int.TryParse(radGridViewCOA.CurrentRow.Cells["IDRIGA"].Value.ToString(), out idriga);

            panelLOG.Left = this.Width / 2 - panelLOG.Width / 2;
            panelLOG.Top = this.Height / 2 - panelLOG.Height / 2;
            panelLOG.Visible = true;
            txtbLOGOperazioni.Text = "Operazioni in corso.....\r\n";

            toolStripStatusLabelLOG.Text = string.Format("{0}","Savataggio dati in corso....");

            updateDatiRiga(idtesta, idriga, CurrentUser, Environment.MachineName);

            //MessageBox.Show("Dati Salvati!");

            panelLOG.Visible = false;

            toolStripStatusLabelLOG.Text = string.Format("{0}","");

            //if (checkBox2.Checked == true)
            //{
            //    invionotifiche(idtesta, idriga);
            //}

            btnCercaCOA_Click(null, null);


        }

        private void txtbQTASPEDIBILE_DoubleClick(object sender, EventArgs e)
        {
            txtbQTASPEDIBILE.Text = txtbQTAGESTRES.Text;
        }


        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (chkCONFEZIONATO.SelectedIndex == 1)
            {
                txtbQTASPEDIBILE.Text = txtbQTAGESTRES.Text;
                
            }

            if (chkCONFEZIONATO.SelectedIndex == 2)
            {
                txtbQTASPEDIBILE.Focus();
                txtbNOTEMAG.Text = "NON SPEDIBILE COMPLETAMENTE";
            }

            if (chkCONFEZIONATO.SelectedIndex == 3)
            {
                txtbQTASPEDIBILE.Text = "0";
                txtbNOTEMAG.Text = "NON SPEDIBILE ";
                checkBox2.Checked = true;
            }

        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            for (int ix = 0; ix < chkCONFEZIONATO.Items.Count; ++ix)
                if (ix != e.Index) chkCONFEZIONATO.SetItemChecked(ix, false);
        }

        private void btnLottoBertello_Click(object sender, EventArgs e)
        {
            if (panelConfezionamentoLotti.Visible == false)
            {
                panelConfezionamentoLotti.Visible = checkBoxFiltroGiac.Visible = true;
                //panelGiacenze.Height = 185;
                panelConfezionamentoLotti.BringToFront();

                

                // carico giacenze bertello

                string filtro = "";

                if (checkBoxFiltroGiac.Checked)
                    filtro = txtbCodice.Text;

                dGVConfezionamentoLotti.DataSource = getGiacenzeBertello(filtro);
            }
            else {
                dGVConfezionamentoLotti.DataSource = null;
                dGVConfezionamentoLotti.Rows.Clear();
                panelConfezionamentoLotti.Visible = checkBoxFiltroGiac.Visible = false;
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void txtbQTASPEDIBILE_TextChanged(object sender, EventArgs e)
        {
            calcolacollidaqta();
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void btnExportExcelTrasp_Click(object sender, EventArgs e)
        {
            ExportModuloTrasportatore();
        }

        private void btnExportExcelTraspM14_Click(object sender, EventArgs e)
        {

            if (cmbTIPIORDINI.SelectedIndex == 1)
            {
                ExportModuloTrasportatoreITA("M12ITA");
            }
            else
            {
                MessageBox.Show("Imposta il tipo ordini ITALIA");
                cmbTIPIORDINI.Focus();
                return;
            }

        }


        private void btnExportExcelTraspM15_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("Vuoi stampare il modulo M15 GROUPAGE?", "Scelta modulo M15", MessageBoxButtons.YesNo)  == DialogResult.Yes)
            {
                ExportModuloTrasportatoreM15("M15GR");
            }
            else
            {
                ExportModuloTrasportatoreM15("M15");
            }
        }


        private void btnExportExcelCOA_Click(object sender, EventArgs e)
        {
            if (cmbTIPIORDINI.SelectedIndex == 2)
            {
                ExportModuloTrasportatoreITA("M10EST");
            }
            else
            {
                MessageBox.Show("Imposta il tipo ordini ESTERO");
                cmbTIPIORDINI.Focus();
                return;
            }

        }


        private void ExportModuloTrasportatore(string Modulo = "", bool all = false, bool knos = true)
        {
            Cursor.Current = Cursors.WaitCursor;

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            List<int> idobjects = new List<int>();


            bool bOK = true;
            string modulo = "";


            if (all)
            {

            }
            else
            {
                // controllo consistenza modulo
                for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
                {
                    //if (radGridViewCOA.SelectedRows[i].Cells["ROWHEADERIMAGE"].Value.ToString() == "check-no.png")
                    //{
                    //    bOK = false;

                    //    MessageBox.Show(string.Format("verificare la data di posizionamento della riga con articolo {0} e salvare prima di aprire il modulo trasportatore", (radGridViewCOA.SelectedRows[i].Cells["ARTICOLO"].Value.ToString())));

                    //    break;
                    //}


                    if ((modulo != "") && (modulo != radGridViewCOA.SelectedRows[i].Cells["MODULO"].Value.ToString()))
                        bOK = false;

                    modulo = radGridViewCOA.SelectedRows[i].Cells["MODULO"].Value.ToString();
                    idobjects.Add(int.Parse(radGridViewCOA.SelectedRows[i].Cells["IDOBJECT"].Value.ToString()));
                }

                if (!bOK)
                {
                    MessageBox.Show("Sono state selezionate delle righe che NON hanno modulo congrurente tra loro a causa del tipo di imballaggio");
                    return;
                }

            }
            if (Modulo != "")
                modulo = Modulo;

            try
            {

                #region oldcode
                //for (int i = 0; i < radGridViewCOA.Columns.Count; i++)
                //{
                //    if (Properties.Settings.Default.MAGModuloTrasportatoreColonne.Contains(radGridViewCOA.Columns[i].Name))
                //    {

                //        DataColumn dc = new DataColumn();
                //        dc.ColumnName = radGridViewCOA.Columns[i].Name;
                //        dc.DataType = radGridViewCOA.Columns[i].DataType;
                //        dt.Columns.Add(dc);
                //    }
                //}

                //for (int j = 0; j < radGridViewCOA.SelectedRows.Count; j++)
                //{
                //    DataRow dr;

                //    dr = dt.NewRow();
                //    for (int i = 0; i < radGridViewCOA.SelectedRows[j].Cells.Count; i++)
                //    {
                //        if (dt.Columns.Contains(radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName))
                //        {
                //            if (radGridViewCOA.SelectedRows[j].Cells[i].Value != null)
                //            {
                //                decimal xc = 0;
                //                if (decimal.TryParse(radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString(), out xc))
                //                {
                //                    dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = xc;
                //                }
                //                else
                //                {
                //                    dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString();
                //                }

                //            }
                //        }
                //    }

                //    dt.Rows.Add(dr);

                //}

                //ds.Tables.Add(dt);
                #endregion
                string xslxmodello = Path.Combine(Application.StartupPath, "XLSXModelli", string.Format("{0}.xlsx", modulo));
                string outfile = Path.Combine(Application.StartupPath, string.Format("{0}.xlsx", modulo));

                string tmpfile = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".xlsx");

                toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", tmpfile);

                ExportExcel.excelFile = tmpfile;
                ExportExcel.creaModulo(dtselected(modulo), xslxmodello, outfile, 0, 9);

                toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", outfile);

                // We will open Filename with wordpad.exe.
                ProcessStartInfo start_info =
                    new ProcessStartInfo("excel.exe", outfile);
                start_info.WindowStyle = ProcessWindowStyle.Maximized;

                // Open wordpad.
                Process proc = new Process();
                proc.StartInfo = start_info;
                proc.Start();

                // Wait for wordpad to finish.
                proc.WaitForExit();

                toolStripStatusLabelLOG.Text = string.Format("{0}", "");

                if (knos)
                {

                    if (MessageBox.Show("Vuoi pubblicare il modulo trasportatore su Knos?", "Pubblicazione su Knos", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {

                        Cursor.Current = Cursors.WaitCursor;
                        ExportExcel.exportToPDF(outfile, outfile + ".pdf");

                        toolStripProgressBarLOG.Visible = true;
                        toolStripProgressBarLOG.Maximum = idobjects.Count;
                        toolStripProgressBarLOG.Minimum = 1;

                        for (int i = 0; i < idobjects.Count; i++)
                        {

                            string docName = string.Format("{0}.xlsx.pdf", modulo);

                            if (Properties.Settings.Default.MNNDataNomeFile == true)
                            {
                                docName = string.Format("{0}_{1}.xlsx.pdf", modulo, System.DateTime.Today.ToString("yyyyMMdd"));
                            }


                            toolStripStatusLabelLOG.Text = string.Format("pubblicazione {0} - doc: {1}", idobjects[i], docName);
                            kw.UploadFileCertificato(idobjects[i], 0, outfile + ".pdf", "MODULO TRASPORTATORE", docName, 0, "", "");
                            toolStripProgressBarLOG.Increment(1);
                        }

                        toolStripStatusLabelLOG.Text = "";

                        Cursor.Current = Cursors.Default;

                        MessageBox.Show("Il modulo è stato pubblicato su Knos");

                        toolStripProgressBarLOG.Visible = false;

                    }
                }
                // AGGIORNA STATO RIGA A MODULO TRASPORTATORE PREPARATO
                // TODO

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Cursor.Current = Cursors.Default;
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");

            }
            finally {
                Cursor.Current = Cursors.Default;
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");
            }
            //try {
            //    ExportToExcelML exporter = new ExportToExcelML(this.radGridViewCOA);
            //    exporter.RunExport(@"C:\\temp\\test.xlsx");
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);

            //}

        }


        private void ExportModuloTrasportatoreM15(string Modulo = "", bool all = false, bool knos = true)
        {
            Cursor.Current = Cursors.WaitCursor;

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            List<int> idobjects = new List<int>();


            bool bOK = true;
            string modulo = "";


            if (all)
            {

            }
            else
            {
                // controllo consistenza modulo
                for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
                {
                    //if (radGridViewCOA.SelectedRows[i].Cells["ROWHEADERIMAGE"].Value.ToString() == "check-no.png")
                    //{
                    //    bOK = false;

                    //    MessageBox.Show(string.Format("verificare la data di posizionamento della riga con articolo {0} e salvare prima di aprire il modulo trasportatore", (radGridViewCOA.SelectedRows[i].Cells["ARTICOLO"].Value.ToString())));

                    //    break;
                    //}


                    //if ((modulo != "") && (modulo != radGridViewCOA.SelectedRows[i].Cells["MODULO"].Value.ToString()))
                    //    bOK = false;

                    //modulo = radGridViewCOA.SelectedRows[i].Cells["MODULO"].Value.ToString();
                    idobjects.Add(int.Parse(radGridViewCOA.SelectedRows[i].Cells["IDOBJECT"].Value.ToString()));
                }

                if (!bOK)
                {
                    MessageBox.Show("Sono state selezionate delle righe che NON hanno modulo congrurente tra loro a causa del tipo di imballaggio");
                    return;
                }

            }


            if (Modulo != "")
                modulo = Modulo;

            try
            {

                #region oldcode
                //for (int i = 0; i < radGridViewCOA.Columns.Count; i++)
                //{
                //    if (Properties.Settings.Default.MAGModuloTrasportatoreColonne.Contains(radGridViewCOA.Columns[i].Name))
                //    {

                //        DataColumn dc = new DataColumn();
                //        dc.ColumnName = radGridViewCOA.Columns[i].Name;
                //        dc.DataType = radGridViewCOA.Columns[i].DataType;
                //        dt.Columns.Add(dc);
                //    }
                //}

                //for (int j = 0; j < radGridViewCOA.SelectedRows.Count; j++)
                //{
                //    DataRow dr;

                //    dr = dt.NewRow();
                //    for (int i = 0; i < radGridViewCOA.SelectedRows[j].Cells.Count; i++)
                //    {
                //        if (dt.Columns.Contains(radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName))
                //        {
                //            if (radGridViewCOA.SelectedRows[j].Cells[i].Value != null)
                //            {
                //                decimal xc = 0;
                //                if (decimal.TryParse(radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString(), out xc))
                //                {
                //                    dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = xc;
                //                }
                //                else
                //                {
                //                    dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString();
                //                }

                //            }
                //        }
                //    }

                //    dt.Rows.Add(dr);

                //}

                //ds.Tables.Add(dt);
                #endregion
                string xslxmodello = Path.Combine(Application.StartupPath, "XLSXModelli", string.Format("{0}.xlsx", modulo));
                string outfile = Path.Combine(Application.StartupPath, string.Format("{0}.xlsx", modulo));

                string tmpfile = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".xlsx");

                toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", tmpfile);

                ExportExcel.excelFile = tmpfile;
                ExportExcel.creaModulo(dtselected(modulo), xslxmodello, outfile, 0, 9);

                toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", outfile);

                // We will open Filename with wordpad.exe.
                ProcessStartInfo start_info =
                    new ProcessStartInfo("excel.exe", outfile);
                start_info.WindowStyle = ProcessWindowStyle.Maximized;

                // Open wordpad.
                Process proc = new Process();
                proc.StartInfo = start_info;
                proc.Start();

                // Wait for wordpad to finish.
                proc.WaitForExit();

                toolStripStatusLabelLOG.Text = string.Format("{0}", "");

                if (knos)
                {

                    if (MessageBox.Show("Vuoi pubblicare il modulo trasportatore su Knos?", "Pubblicazione su Knos", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {

                        Cursor.Current = Cursors.WaitCursor;
                        ExportExcel.exportToPDF(outfile, outfile + ".pdf");

                        toolStripProgressBarLOG.Visible = true;
                        toolStripProgressBarLOG.Maximum = idobjects.Count;
                        toolStripProgressBarLOG.Minimum = 1;

                        for (int i = 0; i < idobjects.Count; i++)
                        {

                            string docName = string.Format("{0}.xlsx.pdf", modulo);

                            if (Properties.Settings.Default.MNNDataNomeFile == true)
                            {
                                docName = string.Format("{0}_{1}.xlsx.pdf", modulo, System.DateTime.Today.ToString("yyyyMMdd"));
                            }


                            toolStripStatusLabelLOG.Text = string.Format("pubblicazione {0} - doc: {1}", idobjects[i], docName);
                            kw.UploadFileCertificato(idobjects[i], 0, outfile + ".pdf", "MODULO TRASPORTATORE", docName, 0, "", "");
                            toolStripProgressBarLOG.Increment(1);
                        }

                        toolStripStatusLabelLOG.Text = "";

                        Cursor.Current = Cursors.Default;

                        MessageBox.Show("Il modulo è stato pubblicato su Knos");

                        toolStripProgressBarLOG.Visible = false;

                    }
                }
                // AGGIORNA STATO RIGA A MODULO TRASPORTATORE PREPARATO
                // TODO

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Cursor.Current = Cursors.Default;
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");

            }
            finally
            {
                Cursor.Current = Cursors.Default;
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");
            }
            //try {
            //    ExportToExcelML exporter = new ExportToExcelML(this.radGridViewCOA);
            //    exporter.RunExport(@"C:\\temp\\test.xlsx");
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);

            //}

        }


        private void ExportModuloTrasportatoreITA(string Modulo = "", bool all = false, bool knos = true)
        {
            Cursor.Current = Cursors.WaitCursor;

            bool bOK = true;
            string modulo = "";

            if (Modulo != "")
                modulo = Modulo;

            try
            {

                string xslxmodello = Path.Combine(Application.StartupPath, "XLSXModelli", string.Format("{0}.xlsm", modulo));
                string outfile = Path.Combine(Application.StartupPath, string.Format("{0}.xlsm", modulo));

                string tmpfile = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".xlsm");

                toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", tmpfile);

                ExportExcel.excelFile = tmpfile;

                DataRowView r = ((DataRowView)comboBoxSpedizionieri.SelectedItem);
                

                ExportExcel.creaModuloM12ITA(System.DateTime.Today, r["CODICE"].ToString(), xslxmodello, outfile, 0, 9);

                toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", outfile);
                
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Cursor.Current = Cursors.Default;
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");

            }
            finally
            {
                Cursor.Current = Cursors.Default;
                toolStripStatusLabelLOG.Text = string.Format("{0}", "");
            }
        }


        DataTable dtselected(string Modulo)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            try
            {

                System.Collections.Specialized.StringCollection xlscolumns = null;

                if (Modulo == "M05")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM05Export;
                }

                if (Modulo == "M12")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM12TrasportatoreColonne;
                }

                if (Modulo == "M12ITA")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM12ITATrasportatoreColonne;
                }

                if (Modulo == "M13")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM13TrasportatoreColonne;
                }

                if (Modulo == "M14")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM14TrasportatoreColonne;
                }

                if (Modulo == "M15GR")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM15GRTrasportatoreColonne;
                }

                if (Modulo == "M15")
                {
                    xlscolumns = Properties.Settings.Default.MAGModuloM15TrasportatoreColonne;
                }

                foreach (string f in xlscolumns)
                {

                    for (int i = 0; i < radGridViewCOA.Columns.Count; i++)
                    {

                       if (f == radGridViewCOA.Columns[i].Name)
                        {
                            Debug.Print(string.Format("{0} trovato!", f));
                            DataColumn dc = new DataColumn();
                            dc.ColumnName = radGridViewCOA.Columns[i].Name;
                            if (radGridViewCOA.Columns[i].DataType.FullName == "System.Int32")
                            {
                                dc.DataType = typeof(Int32);
                            }
                            else
                            {
                                dc.DataType = radGridViewCOA.Columns[i].DataType;
                            }

                            dt.Columns.Add(dc);
                        }
                    }
                }


                if (radGridViewCOA.SelectedRows.Count == 0)
                {
                    radGridViewCOA.SelectAll();
                }

                string prevIDTESTA = "0";
                string prevADR = "";
                List<string> m14SelectedIndex = new List<string>();

                if (Modulo == "M14")
                {
                    for (int j = 0; j < radGridViewCOA.SelectedRows.Count; j++)
                    {

                        if ((Modulo == "M14") && ((radGridViewCOA.SelectedRows[j].Cells["IDTESTA"].Value.ToString() != prevIDTESTA) || (radGridViewCOA.SelectedRows[j].Cells["CLASSEADR"].Value.ToString() != "")))
                        {
                            m14SelectedIndex.Add(string.Format("{0}|{1}|{2}", radGridViewCOA.SelectedRows[j].Index, radGridViewCOA.SelectedRows[j].Cells["IDTESTA"].Value.ToString(), radGridViewCOA.SelectedRows[j].Cells["CLASSEADR"].Value.ToString()));

                            prevIDTESTA = radGridViewCOA.SelectedRows[j].Cells["IDTESTA"].Value.ToString();
                            prevADR = radGridViewCOA.SelectedRows[j].Cells["CLASSEADR"].Value.ToString();
                        }
                    }
                }

                prevIDTESTA = prevADR = "";

                    for (int j = 0; j < radGridViewCOA.SelectedRows.Count; j++)
                    {
                        DataRow dr;

                        if ((Modulo == "M14") && ((radGridViewCOA.SelectedRows[j].Cells["IDTESTA"].Value.ToString() != prevIDTESTA) || (radGridViewCOA.SelectedRows[j].Cells["CLASSEADR"].Value.ToString() != "")))
                        {
                            dr = dt.NewRow();

                            for (int i = 0; i < radGridViewCOA.SelectedRows[j].Cells.Count; i++)
                            {
                                if (dt.Columns.Contains(radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName))
                                {
                                    if (radGridViewCOA.SelectedRows[j].Cells[i].Value != null)
                                    {
                                        decimal xc = 0;
                                        if (decimal.TryParse(radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString(), out xc))
                                        {
                                            dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = xc;
                                        }
                                        else
                                        {
                                            dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString();
                                        }

                                    }
                                }
                            }

                            dt.Rows.Add(dr);

                            prevIDTESTA = radGridViewCOA.SelectedRows[j].Cells["IDTESTA"].Value.ToString();
                            prevADR = radGridViewCOA.SelectedRows[j].Cells["CLASSEADR"].Value.ToString();
                        }

                        if (Modulo != "M14")
                        {

                            dr = dt.NewRow();

                            for (int i = 0; i < radGridViewCOA.SelectedRows[j].Cells.Count; i++)
                            {
                                if (dt.Columns.Contains(radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName))
                                {
                                    if (radGridViewCOA.SelectedRows[j].Cells[i].Value != null)
                                    {
                                        decimal xc = 0;
                                        if (decimal.TryParse(radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString(), out xc))
                                        {
                                            dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = xc;
                                        }
                                        else
                                        {
                                            dr[radGridViewCOA.SelectedRows[j].Cells[i].ColumnInfo.FieldName] = radGridViewCOA.SelectedRows[j].Cells[i].Value.ToString();
                                        }

                                    }
                                }
                            }

                            dt.Rows.Add(dr);
                        }

                    }
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            return dt;
        }


        private void btnLottoZSI_Click(object sender, EventArgs e)
        {
            if (panelConfezionamentoLotti.Visible == false)
            {
                panelConfezionamentoLotti.Visible = true;
                //panelConfezionamentoLotti.Height = 185;
                panelConfezionamentoLotti.BringToFront();

                dGVConfezionamentoLotti.DataSource = getConfezionamentoLotti(idtesta, idriga);
            }
            else
            {
                dGVConfezionamentoLotti.DataSource = getConfezionamentoLotti(0, 0);
                panelConfezionamentoLotti.Visible = false;
            }
        }




        DataTable getConfezionamentoLotti(int idt, int idr)
        {
            DataTable x = new DataTable();

            string sql = string.Format("SELECT * FROM EXCEL_CONFEZIONATOSPEDIBILE WHERE PROGRESSIVO = {0} AND IDRIGA = {1} AND IDTESTADDT = 0 AND IDRIGADDT = 0", idt, idr);

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);
                }
            }
            return x;
        }

        private void btnSalvaConfezionamentoLotti_Click(object sender, EventArgs e)
        {
            // STORED PROCEDURE CHE AGGIORNA IL FRAZIONAMENTO

            double totQtaSpedibile = 0;
            double totQtaDaSpedire = 0;
            double totQtaResidua = 0;

            double.TryParse(txtbQTAGESTRES.Text, out totQtaDaSpedire);

            for (int i = 0; i < dGVConfezionamentoLotti.Rows.Count-1; i++)
            {
                double qtaspedibile = 0;
                double numcolli = 0;

                if (dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_QTA"].Value == null)
                    dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_QTA"].Value = 0;
                if (dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NUMCOLLI"].Value == null)
                    dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NUMCOLLI"].Value = 0;

                if ((dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value == null) || (dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value.ToString() == ""))
                    dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value = "--- L" + i.ToString();

                double.TryParse(dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_QTA"].Value.ToString(), out qtaspedibile);
                double.TryParse(dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NUMCOLLI"].Value.ToString(), out numcolli);

                

                if (nrpezziimballo > 0)
                {
                    if ((qtaspedibile == 0) && (numcolli > 0))
                    {
                        dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_QTA"].Value = numcolli * nrpezziimballo;
                        qtaspedibile =  numcolli * nrpezziimballo;
                    }

                    if ((qtaspedibile > 0) && (numcolli >= 0))
                    {
                        dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NUMCOLLI"].Value = qtaspedibile / nrpezziimballo;
                    }
                }

                if ((dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_DATACONSEGNA"].Value == null) || (dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_DATACONSEGNA"].Value.ToString() == ""))
                    dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_DATACONSEGNA"].Value = System.DateTime.Parse(dTPPosizionamento.Value.ToShortDateString());


                if ((dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value.ToString() == "NESSUN LOTTO") || dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value.ToString().StartsWith("---"))
                {
                    totQtaResidua += qtaspedibile;

                    // lotto non assegnato
                    qtaspedibile = 0;
                }

                totQtaSpedibile += qtaspedibile;
            }


            if ((totQtaSpedibile + totQtaResidua) > totQtaDaSpedire)
            {
                MessageBox.Show("NON è possibile spedire più di quanto previsto in ordine!", "Controllo frazionamento in lotti", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            if ((totQtaSpedibile + totQtaResidua) < totQtaDaSpedire)
            {
                // aggiunge la riga residua
                //dGVConfezionamentoLotti.Rows.Add();

                DataGridViewRow r = dGVConfezionamentoLotti.Rows[dGVConfezionamentoLotti.Rows.Count-1];

                r.Cells["LOTTI_QTA"].Value = totQtaDaSpedire - totQtaSpedibile;
                r.Cells["LOTTI_NUMCOLLI"].Value = (totQtaDaSpedire - totQtaSpedibile)/nrpezziimballo;
                r.Cells["LOTTI_DATACONSEGNA"].Value = dTPDataCarico.Value;
                r.Cells["LOTTI_NRLOTTO"].Value = "---";


            }


            if (MessageBox.Show("Salvo la assegnazione dei lotti?", "Assegnazione Lotti", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string sql = string.Format("DELETE FROM  ZSI_CONFEZSPEDIBILE WHERE PROGRESSIVO = {0} AND IDRIGA = {1} AND IDTESTADDT = 0 AND IDRIGADDT = 0", idtesta, idriga);

                using (SqlConnection cn = new SqlConnection(sqlConnectionString))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand(sql))
                    {
                        cmd.Connection = cn;
                        cmd.ExecuteNonQuery();
                    }


                    for (int i = 0; i < dGVConfezionamentoLotti.Rows.Count; i++)
                    {
                        //dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_IDTESTA"].Value = idtesta;
                        //dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_IDRIGA"].Value = idriga;

                        sql = string.Format("ITA_SP_UPDATE_DATISPEDIZIONE_LOTTO"); // INSERT INTO dbo.ZS_SPLITLOTTI (IDTESTA, IDRIGA, QTAGEST, DATACONSEGNA, NUMCOLLI, NRLOTTO, UTENTEMODIFICA, DATAMODIFICA) VALUES ({0}, {1}, {2}, {3}, {4}, {5}, {6}, GETDATE())",
                            //idtesta,
                            //idriga,
                            //dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_QTA"].Value,
                            //null,
                            //dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NUMCOLLI"].Value,
                            //dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value,
                            //"trm");

                        if (dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value != null)
                        {

                            using (SqlCommand cmd = new SqlCommand(sql))
                            {
                                //cn.Open();
                                cmd.Connection = cn;
                                cmd.CommandType = CommandType.StoredProcedure;

                                cmd.Parameters.AddWithValue("idtesta", idtesta);
                                cmd.Parameters.AddWithValue("idriga", idriga);
                                cmd.Parameters.AddWithValue("qtaspedibile", dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_QTA"].Value);
                                cmd.Parameters.AddWithValue("nrlotto", dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_NRLOTTO"].Value.ToString());
                                cmd.Parameters.AddWithValue("dataconsegna", dGVConfezionamentoLotti.Rows[i].Cells["LOTTI_DATACONSEGNA"].Value);

                                cmd.ExecuteNonQuery();

                            }
                        }

                    }



                }

                panelConfezionamentoLotti.Visible = false;


            }

            txtbQTASPEDIBILE.Text = totQtaSpedibile.ToString();
            txtbQTASPEDIBILE.Focus();

            //updateDatiRiga(idtesta, idriga, CurrentUser, Environment.MachineName);

        }

        private void txtNRCOLLISPEDIBILI_Leave(object sender, EventArgs e)
        {
            int nrcollispedibili = 0;

            int.TryParse(txtNRCOLLISPEDIBILI.Text, out nrcollispedibili);

            txtbQTASPEDIBILE.Text = (nrcollispedibili * nrpezziimballo).ToString();
        }

        void calcolacollidaqta()
        {
            int nrcollispedibili = 0;
            double qtadaspedire = 0;

            double.TryParse(txtbQTASPEDIBILE.Text, out qtadaspedire);

            if (nrpezziimballo > 0)
            {

                txtNRCOLLISPEDIBILI.Text = (qtadaspedire / nrpezziimballo).ToString();

            }
        }


        private void btnRichiestaProduzione_Click(object sender, EventArgs e)
        {
            // aggirono registro
            int idtesta = 0;
            int idriga = 0;

            for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
            {



                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);




                string strSQL = string.Format("ITA_SP_UPDATE_RICHIESTAPRODUZIONE");

                using (SqlConnection cn = new SqlConnection(sqlConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(strSQL))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("idtesta", idtesta);
                        cmd.Parameters.Add("idriga", idriga);
                        cmd.ExecuteNonQuery();
                    }

                }

                tabControl1.SelectedIndex = 2;

            }
        }

        private DataTable getRichiesteProduzione()
        {
            DataTable x = new DataTable();

            string strSQL = "SELECT * FROM EXCEL_RICHIESTAPRODUZIONE";
//            strSQL = string.Format(strSQL, strWHERE) + " ORDER BY DATADOC DESC, RAGIONESOCIALE";

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);

                }

            }
            return x;


        }

        private void splitContainer2_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bindingSource2_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void chkMAGSTATOM05_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            for (int ix = 0; ix < chkMAGSTATOM05.Items.Count; ++ix)
                if (ix != e.Index) chkMAGSTATOM05.SetItemChecked(ix, false);
        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void btnSalvaImpostazioni_Click(object sender, EventArgs e)
        {

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            string sf = Path.Combine(s, System.Environment.UserName + ".xml");


            radGridViewCOA.SaveLayout(sf);

            if (File.Exists(sf))
            {
                try
                {
                    if (Directory.Exists(Path.Combine(Properties.Settings.Default.pathServerApp, "GridLayout")))
                    {
                        string df = Path.Combine(Path.Combine(Properties.Settings.Default.pathServerApp, "GridLayout"), System.Environment.UserName + ".xml");
                        File.Copy(sf, df, true);
                    }
                }
                catch (Exception ex)
                {

                }

            }


        }

        void getImpostazioniGriglia()
        {

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            string[] fi = Directory.GetFiles(s);

            cmbImpostazioni.DataSource = fi;


            try
            {
                for (int i = 0; i < cmbImpostazioni.Items.Count; i++)
                {
                    if (cmbImpostazioni.Items[i].ToString().Contains(System.Environment.UserName + ".xml"))
                    {
                        cmbImpostazioni.SelectedIndex = i;
                        string sf = Path.Combine(s, System.Environment.UserName + ".xml");

                        check_columns();


                        radGridViewCOA.LoadLayout(cmbImpostazioni.Items[i].ToString());
                        break;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delel righe ordine {0}", ex.Message));
            }

        }

        void check_columns()
        {
            try
            {
                string sXML = Path.Combine(s, System.Environment.UserName + ".xml");
                string tXML = Path.Combine(s, "template_magazzino" + ".xml");

                try
                {
                    if (Directory.Exists(Path.Combine(Properties.Settings.Default.pathServerApp, "GridLayout")))
                    {
                        string sf = Path.Combine(Path.Combine(Properties.Settings.Default.pathServerApp, "GridLayout"), System.Environment.UserName + ".xml");
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

                    xmlDOC.Save(sXML);
                }
            }
            catch(Exception ex)
            {

                log.LogSomething(string.Format("ERRORE IN CARICAMENTO IMPOSTAZIONI COLONNE \r\n{0}", ex.Message));

            }
            
        }

        private void cmbImpostazioni_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                string sf = cmbImpostazioni.SelectedItem.ToString();
                radGridViewCOA.LoadLayout(sf);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delel righe ordine {0}", ex.Message));
            }
        }

        private void btnDisponibilita_Click(object sender, EventArgs e)
        {
            if (gBDisponibilita.Visible == false)
            {
                gBDisponibilita.Visible = true;
                //panelConfezionamentoLotti.Height = 185;
                gBDisponibilita.BringToFront();

                rGVDisponibilita.DataSource = getDisponibilita(articolo);
                rGVDisponibilita.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.AllCells);

                rGVMetodo.DataSource = getDisponibilitaMetodo(articolo);
                rGVMetodo.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.AllCells);

                btnDisponibilita.BackColor = Color.Lime;
            }
            else
            {
                rGVDisponibilita.DataSource = null;
                rGVDisponibilita.Rows.Clear();
                gBDisponibilita.Visible = false;

                btnDisponibilita.BackColor = SystemColors.Control;
            }
        }




        DataTable getDisponibilita(string articolo = "")
        {
            DataTable x = new DataTable();

            string sql = string.Format("SELECT * FROM ZSI_VISTA_DISPONIBILITAM05");
            
                
            if (articolo != "")
                sql += string.Format(" WHERE ARTICOLO  = '{0}' ", articolo);

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);
                }
            }
            return x;
        }




        DataTable getDisponibilitaMetodo(string articolo = "")
        {
            DataTable x = new DataTable();

            string sql = string.Format("SELECT * FROM GET_DISPONIBILITAMETODO('{0}')", articolo);

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);
                }
            }
            return x;
        }


        // produzione lotti

        DataTable getOrdiniProduzioneLOTTI(string articolo = "")
        {
            DataTable x = new DataTable();

            string sql = string.Format("SELECT * FROM GET_ORDINIPRODUZIONELOTTIMETODO('{0}')", articolo);

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);
                }
            }
            return x;
        }



        private void label34_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void btnKnos_Click(object sender, EventArgs e)
        {
            panel1.Visible = !panel1.Visible;
            panel1.BringToFront();
        }



        private void btnPosizionamenti_Click(object sender, EventArgs e)
        {
            // filtro per data posizionamento in ZSI_LOGAGGIORNAMENTOPOSIZIONAMENTO
            txtbElencoRighe.Text = "Elenco Righe Selezionate";

            for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
            {
                txtbElencoRighe.Text += string.Format("\r\n - {0} {1}", radGridViewCOA.SelectedRows[i].Cells["ARTICOLO"].Value.ToString(), radGridViewCOA.SelectedRows[i].Cells["DESCRIZIONE"].Value.ToString());
            }

            panelRigheSelezionate.Visible = true;
            panelRigheSelezionate.BringToFront();

        }

        private void btnPosizionamentiMultipli_Click(object sender, EventArgs e)
        {
            int idtesta = 0;
            int idriga = 0;
            int chiudi = 0;

            Cursor c = Cursors.WaitCursor;

            try
            {
                toolStripProgressBarLOG.Minimum = 0;
                toolStripProgressBarLOG.Maximum = radGridViewCOA.SelectedRows.Count;
                toolStripProgressBarLOG.Visible = true;

                for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
                {
                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

                    if (checkInteramente.Checked)
                        chiudi = 1;

                    toolStripProgressBarLOG.Increment(1);
                    updateDatePosizionameto(idtesta, idriga, chiudi, 0, CurrentUser, Environment.MachineName);
                }

                updateDatePosizionameto(0, 0, 0, 1, CurrentUser, Environment.MachineName);
            }
            catch (Exception ex)
            {

            }

            finally
            {
                btnCercaCOA_Click(null, null);

                panelRigheSelezionate.Visible = false;

                toolStripProgressBarLOG.Visible = false;

                c = Cursors.Default;
            }

        }

        private void updateDatePosizionameto(int idt, int idr, int chiudi, int giorno, string user, string machinename)
        {

            if (CurrentUser != "")
            {
                user = CurrentUser;
            }


            if (idtesta == 0)
            {
                MessageBox.Show("Selezionare una riga", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }
            
            string strSQL = string.Format("ITA_SP_UPDATE_DATEPOSIZIONAMENTO");

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("idtesta", idt);
                    cmd.Parameters.Add("idriga", idr);
                    cmd.Parameters.Add("dataposizionamento", dTPPosizionamentoM.Value);
                    cmd.Parameters.Add("codsped", comboBoxSpedizionieriPM.SelectedValue);
                    cmd.Parameters.Add("chiudi", chiudi);
                    cmd.Parameters.Add("giorno", giorno);
                    

                    if (chkUPDCarico.Checked)
                    {
                        cmd.Parameters.Add("datacarico", dTPCaricoM.Value);
                    }
                    else
                    {
                        cmd.Parameters.Add("datacarico", null);
                    }

                    if (chkUPDCarico.Checked)
                    {
                        cmd.Parameters.Add("dataconsegna", dTPConsegnaM.Value);
                    }
                    else
                    {
                        cmd.Parameters.Add("dataconsegna", null);
                    }

                    cmd.Parameters.Add("user", user);

                    if (chkPrevisionale.Checked)
                    {
                        cmd.Parameters.Add("magstatom05", 1);
                    }
                    else
                    {
                        cmd.Parameters.Add("magstatom05", 2);
                    }


                    cmd.ExecuteNonQuery();
                }

            }
        }

        private void updateRaggruppamento(int idt, int idr, string user, string raggruppamento)
        {



            if (idtesta == 0)
            {
                MessageBox.Show("Selezionare una riga", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }

            string strSQL = string.Format("ITA_SP_UPDATE_RAGGRUPPAMENTO");

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("idtesta", idt);
                    cmd.Parameters.Add("idriga", idr);
                    cmd.Parameters.Add("user", user);
                    cmd.Parameters.Add("raggruppamento", raggruppamento);

                    cmd.ExecuteNonQuery();
                }

                }

            
        }

        private void btnAnnullaPosizionamenti_Click(object sender, EventArgs e)
        {
            panelRigheSelezionate.Visible = false;
        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void btnExport2Excel_Click(object sender, EventArgs e)
        {

            //ExportToExcelML exporter = new ExportToExcelML(radGridViewCOA);

            try
            {

                ExportModuloTrasportatore("M05", true, false);

                //string fileNameTMP = (Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".xls"));

                //string fileName = (Path.Combine(Application.StartupPath, "ExportedDataM05.xls"));

                //if (File.Exists(fileName))
                //{
                //    File.Delete(fileName);
                //}
                ////string fileName = "C:\\temp\\ExportedDataM05.xlsx";

                
                //exporter.HiddenColumnOption = HiddenOption.ExportAsHidden;
                //exporter.ExportVisualSettings = false;
                //exporter.SheetName = "M05";

                //exporter.RunExport(fileName);
                //radGridViewCOA.CurrentRowChanging -= radGridViewCOA_CurrentRowChanging;
                //radGridViewCOA.CurrentRowChanging += radGridViewCOA_CurrentRowChanging;


                //// We will open Filename with wordpad.exe.
                //ProcessStartInfo start_info =
                //    new ProcessStartInfo("excel.exe", fileName);
                //start_info.WindowStyle = ProcessWindowStyle.Maximized;

                //// Open wordpad.
                //Process proc = new Process();
                //proc.StartInfo = start_info;
                //proc.Start();

                //// Wait for wordpad to finish.
                //proc.WaitForExit();





            }
            catch (Exception ex)
            {
                //exporter = null;

                MessageBox.Show(ex.Message);
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void cmbLottiODP_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void cmbLottiODP_SelectionChangeCommitted(object sender, EventArgs e)
        {
            txtbLotto.Text = cmbLottiODP.SelectedValue.ToString();
        }

        private void rGVRichiestaProduzione_Click(object sender, EventArgs e)
        {

        }

        private void dTP_POSIZDA_ValueChanged(object sender, EventArgs e)
        {
            dTP_POSIZA.Value = dTP_POSIZDA.Value;
            chkFP.Checked = true;
        }

        private void dTP_POSIZA_ValueChanged(object sender, EventArgs e)
        {
            chkFP.Checked = true;
        }

        private void chkFP_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFP.Checked)
            {
                chkFP.BackColor = Color.Red;
            }
            else
            {
                chkFP.BackColor = SystemColors.Control;
            }
        }

        private void chkUPDCarico_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUPDCarico.Checked)
            {
                chkUPDCarico.BackColor = Color.Red;
                dTPCaricoM.Enabled = true;
            }
            else
            {
                chkUPDCarico.BackColor = SystemColors.Control;
                dTPCaricoM.Enabled = false;
            }
        }

        private void chkUPDConsegna_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUPDConsegna.Checked)
            {
                chkUPDConsegna.BackColor = Color.Red;
                dTPConsegnaM.Enabled = true;
            }
            else
            {
                chkUPDConsegna.BackColor = SystemColors.Control;
                dTPConsegnaM.Enabled = false;
            }
        }

        private void txtbDISPONIBILITA_TextChanged(object sender, EventArgs e)
        {
            double disp = 0;
            double disp_imb = 0;

            if (txtbDISPONIBILITA.Focused == true)
            {
                double.TryParse(txtbDISPONIBILITA.Text, out disp);

                //txtbDISPONIBILITA_IMB.Text = "0";

                if (nrpezziimballo > 0)
                {
                    txtbDISPONIBILITA_IMB.Text = (disp / nrpezziimballo).ToString();
                }
            }


        }

        private void txtbDISPONIBILITA_IMB_TextChanged(object sender, EventArgs e)
        {
            

            double disp = 0;
            double disp_imb = 0;

            if (txtbDISPONIBILITA_IMB.Focused == true)
            {
                double.TryParse(txtbDISPONIBILITA_IMB.Text, out disp_imb);

                txtbDISPONIBILITA.Text = (disp_imb * nrpezziimballo).ToString();
            }
        }
        

        private void txtbDISPONIBILITA_Leave(object sender, EventArgs e)
        {
            updateDisponibilitaRiga(idtesta, idriga, CurrentUser, Environment.MachineName);
        }


        private void updateDisponibilitaRiga(int idt, int idr, string user, string machine)
        {
            if (CurrentUser != "")
            {
                user = CurrentUser;
            }

            if (idtesta == 0)
            {
                MessageBox.Show("Selezionare una riga", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }


            string dt = System.DateTime.Today.ToShortDateString();

            double disponibilita = 0;
            double.TryParse(txtbDISPONIBILITA.Text, out disponibilita);


            string strSQL = string.Format("ITA_SP_UPDATE_DISPONIBILITA");

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("idtesta", idtesta);
                    cmd.Parameters.Add("idriga", idriga);
                    cmd.Parameters.Add("disponibilita", disponibilita);
                    cmd.Parameters.Add("posizionamento", txtbPOSIZIONAMENTO.Text);

                    if (chkMAGSTATOM05.SelectedIndex < 1)
                    {
                        chkMAGSTATOM05.SelectedIndex = 1;

                        chkMAGSTATOM05.SelectedItem = 1;
                    }
                    

                    try
                    {

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("{0}", ex.Message), "Errore");
                    }
                }

            }
        }

        private void txtbDISPONIBILITA_IMB_Leave(object sender, EventArgs e)
        {
            updateDisponibilitaRiga(idtesta, idriga, CurrentUser, Environment.MachineName);
        }

        private void toolStripStatusLabelLOG_TextChanged(object sender, EventArgs e)
        {
            txtbLOGOperazioni.Text += string.Format("{0}\r\n", toolStripStatusLabelLOG.Text);
        }

        private void txtbDISPONIBILITA_DoubleClick(object sender, EventArgs e)
        {
            txtbDISPONIBILITA.Focus();
            txtbDISPONIBILITA.Text = txtbQTAGESTRES.Text;
            txtbPOSIZIONAMENTO.Text = "X";
            txtbPOSIZIONAMENTO.Focus();
            updateDisponibilitaRiga(idtesta, idriga, CurrentUser, Environment.MachineName);

        }

        private void txtbPOSIZIONAMENTO_Leave(object sender, EventArgs e)
        {
            updateDisponibilitaRiga(idtesta, idriga, CurrentUser, Environment.MachineName);
        }

        private void bindingSource2_CurrentChanged_1(object sender, EventArgs e)
        {

        }

        private void btnExportExcelM04_Click(object sender, EventArgs e)
        {
            if (cmbTIPIORDINI.SelectedIndex == 2)
            {
                ExportModuloTrasportatoreITA("M04");
            }
            else
            {
                MessageBox.Show("Imposta il tipo ordini ESTERO");
                cmbTIPIORDINI.Focus();
                return;
            }
        }

        private void frmMagazzino_FormClosing(object sender, FormClosingEventArgs e)
        {
            btnSalvaImpostazioni_Click(null, null);
        }

        
        private void frmMagazzino_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                // help
                string help = Path.Combine(Application.StartupPath, "ZSI-GESTIONE DELLE SPEDIZIONI.pdf");

                if (File.Exists(help))
                {
                    try
                    {
                        Process.Start(help);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show(string.Format("Manca il file {0}", help));
                }

            }

            if (e.KeyCode == Keys.F5)
            {
                btnCercaCOA_Click(null, null);

            }
        }

        private void itaCalendarObject2_Load(object sender, EventArgs e)
        {

        }

        private void radGridViewCOA_GroupSummaryEvaluate(object sender, GroupSummaryEvaluationEventArgs e)
        {
            if (e.SummaryItem.Name == "DATAPOSIZIONAMENTO")
            {
                e.FormatString = String.Format("Posizionamento : {0}", e.Value);
            }
        }

        private void itaCalendarObject1_Load(object sender, EventArgs e)
        {

        }

        private void btnSalvaRaggruppamenti_Click(object sender, EventArgs e)
        {

            int idtesta = 0;
            int idriga = 0;
            int chiudi = 0;

            Cursor c = Cursors.WaitCursor;

            try
            {
                toolStripProgressBarLOG.Minimum = 0;
                toolStripProgressBarLOG.Maximum = radGridViewCOA.SelectedRows.Count;
                toolStripProgressBarLOG.Visible = true;

                for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
                {
                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

                    if (checkInteramente.Checked)
                        chiudi = 1;

                    toolStripProgressBarLOG.Increment(1);
                    updateRaggruppamento(idtesta, idriga, CurrentUser, txtbRaggruppamentoM.Text);
                }

                
            }
            catch (Exception ex)
            {

            }

            finally
            {
                btnCercaCOA_Click(null, null);

                panelRaggrupamento.Visible = false;

                toolStripProgressBarLOG.Visible = false;

                c = Cursors.Default;
            }
        }

        private void btnRaggruppamenti_Click(object sender, EventArgs e)
        {
            txtbElencoRigheRaggruppamento.Text = "Elenco Righe Selezionate";

            for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
            {
                txtbElencoRigheRaggruppamento.Text += string.Format("\r\n - {0} {1}", radGridViewCOA.SelectedRows[i].Cells["ARTICOLO"].Value.ToString(), radGridViewCOA.SelectedRows[i].Cells["DESCRIZIONE"].Value.ToString());
            }

            panelRaggrupamento.Visible = true;
            panelRaggrupamento.BringToFront();
        }

        private void txtbRAGGRUPPAMENTO_Leave(object sender, EventArgs e)
        {
            updateRaggruppamento(idtesta, idriga, CurrentUser, txtbRAGGRUPPAMENTO.Text);
        }

        private void btnAnnullaRaggruppamenti_Click(object sender, EventArgs e)
        {
            txtbRaggruppamentoM.Text = txtbElencoRigheRaggruppamento.Text = "";
            panelRaggrupamento.Visible = false;
        }

        private void lblDOCUMENTO_Click(object sender, EventArgs e)
        {
            string url = "";
            frmWebBrowser frm = new frmWebBrowser();

            if (TipoCurrentUser(CurrentUser) == (int)TipoUtente.BackOffice)
            {
                if (MessageBox.Show("Apro il modulo del documento?", "Apertura allegato Knos", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    url = "{0}/KnoS_Catalog/0/{1}/{2}/{3}";
                    url = string.Format(url, txtKnosUrl.Text, radGridViewCOA.SelectedRows[0].Cells["IDOBJECT_DOC"].Value.ToString(), radGridViewCOA.SelectedRows[0].Cells["IDDOC"].Value.ToString(), radGridViewCOA.SelectedRows[0].Cells["FILENAME"].Value.ToString().Substring(0, 15) + ".PDF");
                    url = url.Replace("#", "_");

                    
                    frm.Text = url;
                    frm.url = url;
                    frm.ShowDialog();
                }

            }

            if (radGridViewCOA.SelectedRows[0].Cells["IDDOC_FDS"].Value.ToString() != "")
            {
                url = "{0}/KnoS_Catalog/0/{1}/{2}/{3}";
                url = string.Format(url, txtKnosUrl.Text, radGridViewCOA.SelectedRows[0].Cells["IDOBJECT_DOC"].Value.ToString(), radGridViewCOA.SelectedRows[0].Cells["IDDOC_FDS"].Value.ToString(), radGridViewCOA.SelectedRows[0].Cells["FILENAME_FDS"].Value.ToString().Substring(0, 19) + ".PDF");
                url = url.Replace("#", "_");


                frm.Text = url;
                frm.url = url;
                frm.ShowDialog();
            }

        }

        private void radGridViewCOA_ViewCellFormatting(object sender, Telerik.WinControls.UI.CellFormattingEventArgs e)
        {
            Font x = new Font(radGridViewCOA.Font.FontFamily, radGridViewCOA.Font.Size + 1, FontStyle.Bold);

            if (e.CellElement is GridSummaryCellElement)
            {
                e.CellElement.TextAlignment = ContentAlignment.MiddleCenter;
                e.CellElement.Font = new Font(e.CellElement.Font, FontStyle.Bold);
                e.Row.Height = 20;
            }

            var groupContentCellElement = sender as GridGroupContentCellElement;
            if (groupContentCellElement != null)
            {
                groupContentCellElement.DrawFill = true;
                groupContentCellElement.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                groupContentCellElement.Font = x;

                if (groupContentCellElement.RowIndex % 2 == 0)
                {
                    groupContentCellElement.BackColor = Color.LightGoldenrodYellow;

                }
                else
                {
                    groupContentCellElement.BackColor = Color.LightBlue;
                }
            }
        }

        private void txtbLotto_Leave(object sender, EventArgs e)
        {
            updateLottoRiga(idtesta, idriga, CurrentUser, Environment.MachineName);
        }

        private void updateLottoRiga(int idt, int idr, string user, string machine)
        {
            if (CurrentUser != "")
            {
                user = CurrentUser;
            }

            if (idtesta == 0)
            {
                MessageBox.Show("Selezionare una riga", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }


            string dt = System.DateTime.Today.ToShortDateString();

            string nrlotto = txtbLotto.Text;

            string strSQL = string.Format("ITA_SP_UPDATE_LOTTORIGA");

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("idtesta", idtesta);
                    cmd.Parameters.Add("idriga", idriga);
                    cmd.Parameters.Add("nrlotto", nrlotto);
                    cmd.Parameters.Add("user", user);
                    cmd.Parameters.Add("machine", machine);


                    try
                    {

                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("{0}", ex.Message), "Errore");
                    }
                }

            }
        }

        private int TipoCurrentUser(string user)
        {
            int TipoUtente = 3;

            if (CurrentUser != "")
            {
                user = CurrentUser;
            }

            string strSQL = string.Format("SELECT TOP 1 GRUPPOM05 FROM TABUTENTI WHERE USERID = '{0}'", user);

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.Text;

                    try
                    {

                        int.TryParse(cmd.ExecuteScalar().ToString(), out TipoUtente);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("{0}", ex.Message), "Errore");
                    }
                }

            }

            return TipoUtente;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (btnViewPanel.Text == "Nascondi")
            {
                splitContainer2.Panel2Collapsed= true;
                btnViewPanel.Text = "Visualizza";
            }
            else
            {
                splitContainer2.Panel2Collapsed = false;
                btnViewPanel.Text = "Nascondi";
            }

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // disponibilità totale

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void radGridViewCOA_ContextMenuOpening(object sender, ContextMenuOpeningEventArgs e)
        {
            contextMenu = new RadContextMenu();
            RadMenuItem menuDispComp = new RadMenuItem("Disponibilità Completa (X)");
            menuDispComp.ForeColor = Color.DarkGreen;
            menuDispComp.Click += new EventHandler(menuDispComp_Click);
            RadMenuItem menuDispParz = new RadMenuItem("Disponibilità Parziale");
            menuDispParz.Click += new EventHandler(menuDispParz_Click);

            RadMenuHeaderItem AltezzaRighe = new RadMenuHeaderItem("Altezza righe");
            RadMenuComboItem menuAltezzaRighe = new RadMenuComboItem();
            menuAltezzaRighe.Items.Add("Standard");
            menuAltezzaRighe.Items.Add("Medio");
            menuAltezzaRighe.Items.Add("Grande");

            menuAltezzaRighe.Click += new EventHandler(menuAltezzaRighe_Click);

            RadMenuSeparatorItem separator = new RadMenuSeparatorItem();
            e.ContextMenu.Items.Add(separator);
            e.ContextMenu.Items.Add(menuDispComp);
            e.ContextMenu.Items.Add(menuDispParz);
            e.ContextMenu.Items.Add(separator);
            e.ContextMenu.Items.Add(AltezzaRighe);
            e.ContextMenu.Items.Add(menuAltezzaRighe);
            //e.ContextMenu = contextMenu.DropDown;
        }


        private void menuDispComp_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
            {
                int.TryParse(radGridViewCOA.SelectedRows[0].Cells["IDTESTA"].Value.ToString(), out idtesta);
                int.TryParse(radGridViewCOA.SelectedRows[0].Cells["IDRIGA"].Value.ToString(), out idriga);


                txtbDISPONIBILITA_DoubleClick(null, null);

            }


            return;

        }

        private void menuDispParz_Click(object sender, EventArgs e)
        {

            return;

        }


        private void menuAltezzaRighe_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
            {
                int.TryParse(radGridViewCOA.SelectedRows[0].Cells["IDTESTA"].Value.ToString(), out idtesta);
                int.TryParse(radGridViewCOA.SelectedRows[0].Cells["IDRIGA"].Value.ToString(), out idriga);


                txtbDISPONIBILITA_DoubleClick(null, null);

            }


            return;

        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            Application.DoEvents();

            DataSet xx = getM10();

            radGridViewM10.DataSource = xx.Tables[0];
        }

        DataSet getM10()
        {


            
            DataSet dsOut = new DataSet();

            string strSQL = Properties.Settings.Default.MAGMetodoCommandM10 + " {0}";

            string strWHERE = " WHERE 1=1";

            strWHERE += String.Format(" AND DATAPOSIZIONAMENTO = '{0}' ", monthCalendar1.SelectionStart.ToString("yyyyMMdd"));




            //if (cmbTIPIORDINI.SelectedIndex == 0)
            //{
            //    if ((Properties.Settings.Default.MAGTipiDocItalia != "") && (Properties.Settings.Default.MAGTipiDocEstero != ""))
            //    {
            //        strWHERE += string.Format(" AND TIPODOC IN ({0}, {1})", Properties.Settings.Default.MAGTipiDocItalia, Properties.Settings.Default.MAGTipiDocEstero);
            //    }
            //}

            //if (cmbTIPIORDINI.SelectedIndex == 1)
            //{
            //    if (Properties.Settings.Default.MAGTipiDocItalia != "")
            //    {
            //        strWHERE += string.Format(" AND TIPODOC IN ({0})", Properties.Settings.Default.MAGTipiDocItalia);
            //    }
            //}

            //if (cmbTIPIORDINI.SelectedIndex == 2)
            //{
            //    if (Properties.Settings.Default.MAGTipiDocEstero != "")
            //    {
            //        strWHERE += string.Format(" AND TIPODOC IN ({0})", Properties.Settings.Default.MAGTipiDocEstero);
            //    }
            //}

            //if (cmbIMBALLI.SelectedIndex <= 0)
            //{
            //}

            //if (cmbIMBALLI.SelectedIndex == 2)
            //{
            //    if (Properties.Settings.Default.MAGTipiImballiEsclusi != "")
            //    {
            //        strWHERE += string.Format(" AND CODIMBALLO IN ({0})", Properties.Settings.Default.MAGTipiImballiEsclusi);
            //    }

            //}

            //if (cmbIMBALLI.SelectedIndex == 1)
            //{
            //    if (Properties.Settings.Default.MAGTipiImballiEsclusi != "")
            //    {
            //        strWHERE += string.Format(" AND CODIMBALLO NOT IN ({0})", Properties.Settings.Default.MAGTipiImballiEsclusi);
            //    }
            //}


            //if (cmbSTATOM05.SelectedIndex <= 0)
            //{
            //}

            strSQL = string.Format(strSQL, strWHERE); // + " ORDER BY DATADOC DESC, RAGIONESOCIALE";






            DataTable x = new DataTable();

            using (SqlConnection cn = new SqlConnection(sqlConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);

                }

            }

            dsOut.Tables.Add(x);

            return dsOut;
        }

        private void radGridViewM10_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {


        }

        private void btnCalendarPrint_Click(object sender, EventArgs e)
        {

            if (radGridViewCOA.SelectedRows.Count == 0)
            {
                MessageBox.Show("Attenzione! devi selezionare una riga per poter stampare il calendario spedizioni del cliente selezionato!", "Stampa Calendario Spedizioni", MessageBoxButtons.OK);
            }
            else
            {
                Cursor.Current = Cursors.WaitCursor;

                using (SqlConnection cn = new SqlConnection(sqlConnectionString))
                {
                    crPrint.frmPrint f = new crPrint.frmPrint();
                    f.cn = cn;
                    f.user = "trm" + "1";
                    f.pwd = "terminale";
                    f.reportFile = Application.StartupPath + "\\Reports\\CalSpedizioni.rpt";
                    f.anteprima = true;
                    f.stampante = "";
                    f.CODCONTO = radGridViewCOA.SelectedRows[0].Cells["CODCLIFOR"].Value.ToString();
                    f.nrGiorniOrizzonte = 1;
                    f.ShowDialog();
                }

                Cursor.Current = Cursors.Default;

            }
        }


        private void toolStripDropDownButton1_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.ToString().ToUpper().Contains("MOSTRA"))
            {
                bolShowTotali = true;
            }
            else
            {
                bolShowTotali = false;
            }

            btnCercaCOA_Click(null, null);
        }

        private void toolStripDropDownButton2_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem.ToString().ToUpper().Contains("STANDARD"))
            {
                hRows = 40;
            }

            if (e.ClickedItem.ToString().ToUpper().Contains("MEDIO"))
            {
                hRows = 80;
            }

            if (e.ClickedItem.ToString().ToUpper().Contains("GRANDE"))
            {
                hRows = 120;
            }

            btnCercaCOA_Click(null, null);
        }

        private void radGridViewCOA_CellClick(object sender, GridViewCellEventArgs e)
        {
            int idtesta = 0;
            int idriga = 0;

            if (radGridViewCOA.Rows.Count > 0)
            {
                if (e.RowIndex >= 0)
                {
                    //datiriga(e.RowIndex);
                    //lognotificheriga(e.RowIndex);

                    if (e.Column.HeaderText == "Ordine")
                    {
                        int.TryParse(radGridViewCOA.Rows[e.RowIndex].Cells["IDTESTA"].Value.ToString(), out idtesta);
                        //int.TryParse(radGridViewCOA.Rows[e.RowIndex].Cells["PROGRIGADOC"].Value.ToString(), out idriga);

                        if ((idtesta > 0))
                        {
                            EmbyonAction(actionRip, string.Format(actionRipKey, idtesta, 0));


                        }
                    }


                }



            }
        }


        private void EmbyonAction(string a, string k)
        {
            try
            {
                log.LogSomething("inizializza");
                MetodoApp.EmbyonConnector.InizializzaAtMetodo();
                log.LogSomething("ditta");
                MetodoApp.EmbyonConnector.ditta = Global.Ditta;
                log.LogSomething("utente");
                MetodoApp.EmbyonConnector.utente = Global.UtenteMetodo;
                log.LogSomething("azione");
                if (MetodoApp.EmbyonConnector.ExecMetodoAction(a, k) == false)
                {
                    string m = string.Format("Tentativo di connessione a Embyon\n-utente:{0}\n-ditta:{1}\n-action:{2}\n-key:{3}", Global.UtenteMetodo, Global.Ditta, a, k);

                    MessageBox.Show("Attenzione!!" + m + "\r\nverificare che Embyon sia aperto sulla ditta corretta e con l'utente corretto!", "Connessione a Embyon", MessageBoxButtons.OK, MessageBoxIcon.Error);

                };
            }
            catch (Exception ex)
            {
                string m = string.Format("Tentativo di connessione a Embyon\n-utente:{0}\n-ditta:{1}\n-action:{2}\n-key:{3}", Global.UtenteMetodo, Global.Ditta, a, k);

                MessageBox.Show("Attenzione!!" + m, "ERRORE Connessione a Embyon", MessageBoxButtons.OK, MessageBoxIcon.Error);

                log.LogSomething(m);
                log.LogSomething(ex.Message);

            }
        }

    }


}

