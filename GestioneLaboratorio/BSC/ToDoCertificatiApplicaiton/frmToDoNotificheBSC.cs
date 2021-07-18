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






namespace ToDoNotificheBSC
{

    public partial class frmToDoNotificheBSC : Form
    {
        Logger log;

        Cursor _cursor;

        int CurrentIdObject = 0;
        int CurrentIdDoc = 0;
        int CurrentIdObjectCertificato = 0;
        int CurrentIdDocCertificato = 0;
        string CurrentFileDescr = "";
        string CurrentFileName = "";
        int CurrentIdAction = 0;
        string CurrentAttrNameData = "";

        public static int CurrentIDStatusPDL = 0;
        public static string CurrentStatusNamePDL = "";
        public static string CurrentPDFPDLUrl = "";



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


        public bool notifyPopUp = true;

        string s = Path.Combine(Application.StartupPath, "GridLayout");




        public class KnoSWrapper
        {


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
                                MessageBox.Show("Utente non loggato da Internet Explorer");
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


            public bool GetPDL(int _idObject, ListView lvAttr, DataGridView dgCertificati, ListView lvFirme, StatusStrip s)
            {
                bool retvalue = false;
                string fileName = "";
                string fileUrl = "";
                string fileLocalPath = "";
                int fileIdDoc = 0;
                string fileDescr = "";

                nrCertificati1F = nrCertificati2F = nrCertificatiTot = nrCertificatiUtente = nrCertificatiUtente1F = nrCertificatiUtente2F = nrCertificatiUtente1FDaFirmare = nrCertificatiUtente2FDaFirmare = 0;

                string UtenteTecnico = "";
                string UtenteResponsabileTecnico = "";
                string UtenteCapoCommessa = "";
                string DataPrimaFirma = "";
                string DataSecondaFirma = "";
                string UtenteResponsabileTecnicoSost = "";
                string UtenteCapoCommessaSost = "";

                lvFirme.Clear();
                lvFirme.Columns.Clear();
                lvFirme.Columns.Add("Utente");
                lvFirme.Columns.Add("PathFileFirma");


                lvAttr.Clear();
                lvAttr.Columns.Clear();
                lvAttr.Columns.Add("Nome Attributo");
                lvAttr.Columns[0].Width = 150;
                lvAttr.Columns.Add("Valore Attributo");
                lvAttr.Columns[1].Width = lvAttr.Width - 150;
                lvAttr.Columns.Add("Campo Attributo");




                dgCertificati.DataSource = null;
                dgCertificati.Refresh();

                DataTable dtCertificati = new DataTable();
                dtCertificati.Columns.Add("IdObject");
                dtCertificati.Columns.Add("IdStatus");
                dtCertificati.Columns.Add("Status");
                dtCertificati.Columns.Add("Tecnico");
                dtCertificati.Columns.Add("DataPrimaFirma");
                //dtCertificati.Columns.Add("ResponsabileTecnico");
                //dtCertificati.Columns.Add("DataSecondaFirma");
                dtCertificati.Columns.Add("CapoCommessa");
                dtCertificati.Columns.Add("File");
                dtCertificati.Columns.Add("Url");
                dtCertificati.Columns.Add("LocalFile");
                dtCertificati.Columns.Add("IdDoc");
                dtCertificati.Columns.Add("FileDescr");
                //dtCertificati.Columns.Add("ResponsabileTecnicoSost");
                dtCertificati.Columns.Add("CapoCommessaSost");


                //KnosInstance.Client.Login(CurrentUser, "sash17ne", out cIdSubject);

                knosObject = KnosInstance.Client.CreateKnosObject();
                knosObjectCertificato = KnosInstance.Client.CreateKnosObject();
                knosObjectCliente = KnosInstance.Client.CreateKnosObject();

                knosObject.GetAllObjectData(0);
                knosObjectCertificato.GetAllObjectData(0);
                knosObjectCliente.GetAllObjectData(0);


                //IKnosResult ikr = knosObject.GetAllObjectData(_idObject);
                IKnosResult ikr = knosObject.GetObjectAttributes(_idObject);

                if (ikr.HasErrors == false)
                {
                    // gestione dello stato del PDL
                    // e link ai PDF
                    CurrentIDStatusPDL = knosObject.IdStatus;
                    CurrentStatusNamePDL = knosObject.StatusName;

                    knosObject.GetObjectDocuments();

                    if (knosObject.DocumentList.ItemCount > 0)
                    {
                        CurrentPDFPDLUrl = knosObject.DocumentList.GetItem(0).GetUrl();
                    }



                    for (int i = 0; i < knosObject.AttrValueList.ItemCount; i++)
                    {
                        lvAttr.Items.Add(knosObject.AttrValueList.GetItem(i).AttrName);
                        lvAttr.Items[i].SubItems.Add(knosObject.AttrValueList.GetItem(i).ToString());
                        lvAttr.Items[i].SubItems.Add(knosObject.AttrValueList.GetItem(i).ColumnName);

                        if ((knosObject.AttrValueList.GetItem(i).DataType == EnumKnosDataType.ShortTextType) && (knosObject.AttrValueList.GetItem(i).AttrName.ToLower().Contains("certificato")))
                        {
                            // attributo numero di certificato
                            strFilePDFPDL = knosObject.AttrValueList.GetItem(i).ToString();
                        }


                        if ((knosObject.AttrValueList.GetItem(i).DataType == EnumKnosDataType.ObjectListType) && (knosObject.AttrValueList.GetItem(i).AttrName.ToLower().Contains("cliente/fornitore")))
                        {


                            IKnosObjectViewList knosObjectViewListCliente = knosObject.AttrValueList.GetItemByColumnName((knosObject.AttrValueList.GetItem(i).ColumnName)).ToKnosObjectViewList();

                            for (int j = 0; j < knosObjectViewListCliente.ItemCount; j++)
                            {

                                // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                                // ci si è loggati a KnoS
                                //	lvAttr.Items[lvAttr.Items.Count-1].SubItems[1].Text	= knosObjectViewListCliente.GetItem(0).AttrValueList.GetItemByColumnName("varchar_04").ToString();

                                knosObjectCliente.GetObjectAttributes(knosObjectViewListCliente.GetItem(j).IdObject);
                                lvAttr.Items[lvAttr.Items.Count - 1].SubItems[1].Text = string.Format("{0} - {1}", knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_04").ToString(), knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_05").ToString());
                            }
                        }



                        if ((knosObject.AttrValueList.GetItem(i).DataType == EnumKnosDataType.ObjectListType) && (knosObject.AttrValueList.GetItem(i).AttrName.ToLower().Contains("certificati")))
                        {
                            IKnosObjectViewList knosObjectViewList = knosObject.AttrValueList.GetItem(i).ToKnosObjectViewList();

                            nrCertificatiTot = knosObjectViewList.ItemCount;

                            s.Text = "Caricamento certificati in corso.....";


                            for (int j = 0; j < knosObjectViewList.ItemCount; j++)
                            {
                                s.Text = string.Format("Caricamento certificati in corso..... ({0}/{1})", j.ToString(), nrCertificatiTot.ToString());
                                // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                                // ci si è loggati a KnoS

                                knosObjectCertificato.GetObjectAttributes(knosObjectViewList.GetItem(j).IdObject);
                                UtenteTecnico = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_51").ToString();
                                DataPrimaFirma = knosObjectCertificato.AttrValueList.GetItemByColumnName("datetime_08").ToString();
                                UtenteResponsabileTecnico = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_52").ToString();
                                DataSecondaFirma = knosObjectCertificato.AttrValueList.GetItemByColumnName("datetime_09").ToString();
                                UtenteCapoCommessa = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_53").ToString();
                                //UtenteResponsabileTecnicoSost = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_54").ToString();
                                UtenteCapoCommessaSost = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_55").ToString();

                                CurrentCapoCommessa = UtenteCapoCommessa;

                                if (UtenteCapoCommessaSost != "")
                                {
                                    CurrentCapoCommessa = UtenteCapoCommessaSost;
                                }

                                ikr = knosObjectViewList.GetItem(j).GetObjectDocuments();
                                if (ikr.HasErrors == false)
                                {

                                    fileName = fileUrl = fileDescr = fileLocalPath = "";
                                    fileIdDoc = 0;
                                    if (knosObjectViewList.GetItem(j).DocumentList.ItemCount == 1)
                                    {
                                        fileName = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).FileName;
                                        fileUrl = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).GetUrl();
                                        fileIdDoc = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).IdDoc;
                                        fileDescr = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).FileDescr;

                                        // pulizia dei file locali
                                        fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);
                                        File.Delete(Path.Combine(fileLocalPath));

                                        ////download local del file
                                        //File.Delete(Path.Combine(Path.GetTempPath(), fileName));
                                        ////ikr = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).DownloadFile(Path.GetTempPath(), fileName);
                                        ////if (ikr.HasErrors == false)
                                        ////{
                                        ////}
                                    }


                                    if (knosObjectViewList.GetItem(j).IdStatus == SignFiles.KnoS_Certificato_IdStatusIniziale)
                                    {
                                        if (UtenteTecnico == CurrentUser)
                                        {
                                            nrCertificatiUtente1FDaFirmare += 1;
                                        }
                                    }

                                    //if (knosObjectViewList.GetItem(j).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                                    //{
                                    //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                                    //    {
                                    //        nrCertificatiUtente2FDaFirmare += 1;
                                    //    }

                                    //}

                                    if (knosObjectViewList.GetItem(j).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                                    {
                                        nrCertificati1F += 1;

                                        if (UtenteTecnico == CurrentUser)
                                        {
                                            nrCertificatiUtente1F += 1;
                                        }
                                    }


                                    //if (knosObjectViewList.GetItem(j).IdStatus == SignFiles.KnoS_Certificato_IdStatus2F)
                                    //{
                                    //    nrCertificati2F += 1;

                                    //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                                    //    {
                                    //        nrCertificatiUtente2F += 1;
                                    //    }

                                    //}



                                    dtCertificati.Rows.Add(
                                        knosObjectViewList.GetItem(j).IdObject,
                                        knosObjectViewList.GetItem(j).IdStatus,
                                        knosObjectViewList.GetItem(j).StatusName,
                                        UtenteTecnico,
                                        DataPrimaFirma,
                                        //UtenteResponsabileTecnico, 
                                        //DataSecondaFirma, 
                                        UtenteCapoCommessa,
                                        fileName,
                                        fileUrl,
                                        fileLocalPath,
                                        fileIdDoc,
                                        fileDescr,
                                        //UtenteResponsabileTecnicoSost,
                                        UtenteCapoCommessaSost
                                        );


                                }


                                s.Text = string.Format("Caricamento certificati completato!", j.ToString(), nrCertificatiTot.ToString());

                                s.Text = string.Format("Caricamento utenti sostitutivi", j.ToString(), nrCertificatiTot.ToString());

                                ListViewItem li;
                                ListViewItem lx = null;

                                // recupero firme
                                ikr = knosObjectViewList.GetItem(j).GetObjectLinks();
                                if (ikr.HasErrors == false)
                                {

                                    for (int xLink = 0; xLink < knosObjectViewList.GetItem(j).LinkList.ItemCount; xLink++)
                                    {
                                        lx = new ListViewItem();
                                        li = new ListViewItem(knosObjectViewList.GetItem(j).LinkList.GetItem(xLink).LinkDescr);

                                        if (lvFirme.Items.Count > 0)
                                        {
                                            lx = lvFirme.FindItemWithText(li.Text, false, 0);
                                        }
                                        else
                                        {
                                            lx = null;
                                        }

                                        if ((lx == null))
                                        {
                                            lvFirme.Items.Add(li);

                                            if (knosObjectViewList.GetItem(j).LinkList.GetItem(xLink).Url.StartsWith("file:"))
                                            {
                                                lvFirme.Items[lvFirme.Items.Count - 1].SubItems.Add(knosObjectViewList.GetItem(j).LinkList.GetItem(xLink).Url.ToString().Replace(@"file://", ""));

                                            }
                                            else
                                            {
                                                lvFirme.Items[lvFirme.Items.Count - 1].SubItems.Add(knosObjectViewList.GetItem(j).LinkList.GetItem(xLink).Url.ToString());
                                            }

                                        }
                                    }

                                }


                            }

                            dgCertificati.DataSource = dtCertificati;


                            if ((nrCertificati2F == nrCertificatiTot) && (CurrentCapoCommessa == CurrentUser))
                            {
                                SignFiles.tipofirma = 2;
                            }



                        }

                    }

                    retvalue = true;
                }
                else
                {
                    MessageBox.Show(ikr.GetError(1).Description);
                }

                return retvalue;


            }


            public bool GetPDLSelector(int _idObject, ListView lvAttr, DataGridView dgCertificati, ListView lvFirme, StatusStrip s)
            {
                bool retvalue = false;
                string fileName = "";
                string fileUrl = "";
                string fileLocalPath = "";
                int fileIdDoc = 0;
                string fileDescr = "";

                nrCertificati1F = nrCertificati2F = nrCertificatiTot = nrCertificatiUtente = nrCertificatiUtente1F = nrCertificatiUtente2F = nrCertificatiUtente1FDaFirmare = nrCertificatiUtente2FDaFirmare = 0;

                string UtenteTecnico = "";
                //string UtenteResponsabileTecnico = "";
                string UtenteCapoCommessa = "";
                string DataPrimaFirma = "";
                //string DataSecondaFirma = "";
                //string UtenteResponsabileTecnicoSost = "";
                string UtenteCapoCommessaSost = "";

                lvFirme.Clear();
                lvFirme.Columns.Clear();
                lvFirme.Columns.Add("Utente");
                lvFirme.Columns.Add("PathFileFirma");


                lvAttr.Clear();
                lvAttr.Columns.Clear();
                lvAttr.Columns.Add("Nome Attributo");
                lvAttr.Columns[0].Width = 150;
                lvAttr.Columns.Add("Valore Attributo");
                lvAttr.Columns[1].Width = lvAttr.Width - 150;
                lvAttr.Columns.Add("Campo Attributo");




                dgCertificati.DataSource = null;
                dgCertificati.Refresh();

                DataTable dtCertificati = new DataTable();
                dtCertificati.Columns.Add("IdObject");
                dtCertificati.Columns.Add("IdStatus");
                dtCertificati.Columns.Add("Status");
                dtCertificati.Columns.Add("Tecnico");
                dtCertificati.Columns.Add("DataPrimaFirma");
                //dtCertificati.Columns.Add("ResponsabileTecnico");
                //dtCertificati.Columns.Add("DataSecondaFirma");
                dtCertificati.Columns.Add("CapoCommessa");
                dtCertificati.Columns.Add("File");
                dtCertificati.Columns.Add("Url");
                dtCertificati.Columns.Add("LocalFile");
                dtCertificati.Columns.Add("IdDoc");
                dtCertificati.Columns.Add("FileDescr");
                //dtCertificati.Columns.Add("ResponsabileTecnicoSost");
                dtCertificati.Columns.Add("CapoCommessaSost");


                //KnosInstance.Client.Login(CurrentUser, "sash17ne", out cIdSubject);

                knosObject = KnosInstance.Client.CreateKnosObject();
                knosObjectCertificato = KnosInstance.Client.CreateKnosObject();
                knosObjectCliente = KnosInstance.Client.CreateKnosObject();

                knosObject.GetAllObjectData(0);
                knosObjectCertificato.GetAllObjectData(0);
                knosObjectCliente.GetAllObjectData(0);

                //IKnosObjectSelector knosObjectSelectorPDL = KnosInstance.Client.CreateKnosObjectSelector();
                //knosObjectSelectorPDL.PageSize = 1;
                //knosObjectSelectorPDL.SelectIdView = 127;
                //knosObjectSelectorPDL.SearchExpression = string.Format("IdObject = {0}", _idObject);

                try
                {




                    //IKnosResult ikr = knosObjectSelectorPDL.GetPage(1);
                    IKnosResult ikr = knosObject.GetObjectAttributes(_idObject);

                    if (ikr.HasErrors == false)
                    {
                        // gestione dello stato del PDL
                        // e link ai PDF
                        CurrentIDStatusPDL = knosObject.IdStatus;
                        CurrentStatusNamePDL = knosObject.StatusName;

                        knosObject.GetObjectDocuments();

                        if (knosObject.DocumentList.ItemCount > 0)
                        {
                            CurrentPDFPDLUrl = knosObject.DocumentList.GetItem(0).GetUrl();
                        }



                        for (int i = 0; i < knosObject.AttrValueList.ItemCount; i++)
                        {
                            lvAttr.Items.Add(knosObject.AttrValueList.GetItem(i).AttrName);
                            lvAttr.Items[i].SubItems.Add(knosObject.AttrValueList.GetItem(i).ToString());
                            lvAttr.Items[i].SubItems.Add(knosObject.AttrValueList.GetItem(i).ColumnName);

                            if ((knosObject.AttrValueList.GetItem(i).DataType == EnumKnosDataType.ShortTextType) && (knosObject.AttrValueList.GetItem(i).AttrName.ToLower().Contains("certificato")))
                            {
                                // attributo numero di certificato
                                strFilePDFPDL = knosObject.AttrValueList.GetItem(i).ToString();
                            }


                            if ((knosObject.AttrValueList.GetItem(i).DataType == EnumKnosDataType.ObjectListType) && (knosObject.AttrValueList.GetItem(i).AttrName.ToLower().Contains("cliente/fornitore")))
                            {


                                IKnosObjectViewList knosObjectViewListCliente = knosObject.AttrValueList.GetItemByColumnName((knosObject.AttrValueList.GetItem(i).ColumnName)).ToKnosObjectViewList();

                                for (int j = 0; j < knosObjectViewListCliente.ItemCount; j++)
                                {

                                    // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                                    // ci si è loggati a KnoS
                                    //	lvAttr.Items[lvAttr.Items.Count-1].SubItems[1].Text	= knosObjectViewListCliente.GetItem(0).AttrValueList.GetItemByColumnName("varchar_04").ToString();

                                    knosObjectCliente.GetObjectAttributes(knosObjectViewListCliente.GetItem(j).IdObject);
                                    lvAttr.Items[lvAttr.Items.Count - 1].SubItems[1].Text = string.Format("{0} - {1}", knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_04").ToString(), knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_05").ToString());
                                }
                            }



                            //if (ikr.HasErrors == false)
                            //{
                            //    // gestione dello stato del PDL
                            //    // e link ai PDF
                            //    CurrentIDStatusPDL = knosObjectSelectorPDL.GetItem(0).IdStatus;
                            //    CurrentStatusNamePDL = knosObjectSelectorPDL.GetItem(0).StatusName;

                            //    knosObjectSelectorPDL.GetItem(0).GetObjectDocuments();

                            //    if (knosObjectSelectorPDL.GetItem(0).DocumentList.ItemCount > 0)
                            //    {
                            //        CurrentPDFPDLUrl = knosObjectSelectorPDL.GetItem(0).DocumentList.GetItem(0).GetUrl();
                            //    }



                            //    for (int i = 0; i < knosObjectSelectorPDL.GetItem(0).AttrValueList.ItemCount; i++)
                            //    {
                            //        lvAttr.Items.Add(knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).AttrName);
                            //        lvAttr.Items[i].SubItems.Add(knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).ToString());
                            //        lvAttr.Items[i].SubItems.Add(knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).ColumnName);

                            //        if ((knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).DataType == EnumKnosDataType.ShortTextType) && (knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).AttrName.ToLower().Contains("certificato")))
                            //        {
                            //            // attributo numero di certificato
                            //            strFilePDFPDL = knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).ToString();
                            //        }


                            //        if ((knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).DataType == EnumKnosDataType.ObjectListType) && (knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).AttrName.ToLower().Contains("cliente/fornitore")))
                            //        {


                            //            IKnosObjectViewList knosObjectViewListCliente = knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItemByColumnName((knosObjectSelectorPDL.GetItem(0).AttrValueList.GetItem(i).ColumnName)).ToKnosObjectViewList();

                            //            for (int j = 0; j < knosObjectViewListCliente.ItemCount; j++)
                            //            {

                            //                // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                            //                // ci si è loggati a KnoS
                            //                //	lvAttr.Items[lvAttr.Items.Count-1].SubItems[1].Text	= knosObjectViewListCliente.GetItem(0).AttrValueList.GetItemByColumnName("varchar_04").ToString();

                            //                knosObjectCliente.GetObjectAttributes(knosObjectViewListCliente.GetItem(j).IdObject);
                            //                lvAttr.Items[lvAttr.Items.Count - 1].SubItems[1].Text = string.Format("{0} - {1}", knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_04").ToString(), knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_05").ToString());
                            //            }
                            //        }

                        }

                        IKnosObjectSelector knosObjectSelectorCertificati = KnosInstance.Client.CreateKnosObjectSelector();

                        knosObjectSelectorCertificati.SearchExpression = string.Format("IDClass = 47 AND IdObject in (SELECT IdChild FROM Object_Linkage WHERE idparent = {0} AND IdAttr = 120)", _idObject);
                        knosObjectSelectorCertificati.PageSize = 50;
                        knosObjectSelectorCertificati.SelectIdView = 125;
                        knosObjectSelectorCertificati.GetPage(1);

                        //MessageBox.Show("inizio caricamento certificati");
                        nrCertificatiTot = knosObjectSelectorCertificati.RecordCount;
                        for (int i = 0; i < nrCertificatiTot; i++)
                        {
                            //MessageBox.Show(string.Format("Caricamento certificati in corso..... ({0}/{1})", i.ToString(), nrCertificatiTot.ToString()));
                            s.Text = string.Format("Caricamento certificati in corso..... ({0}/{1})", (i + 1).ToString(), nrCertificatiTot.ToString());
                            // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                            // ci si è loggati a KnoS

                            UtenteTecnico = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_51").ToString();
                            try
                            {
                                DataPrimaFirma = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("datetime_08").ToString();
                            }
                            catch (Exception ex)
                            {
                                DataPrimaFirma = "";
                            }
                            //UtenteResponsabileTecnico = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_52").ToString();
                            //try
                            //{
                            //    DataSecondaFirma = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("datetime_09").ToString();
                            //}
                            //catch (Exception ex)
                            //{
                            //    DataSecondaFirma = "";
                            //}
                            //UtenteCapoCommessa = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_53").ToString();
                            //try
                            //{
                            //    UtenteResponsabileTecnicoSost = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_54").ToString();
                            //}
                            //catch (Exception ex)
                            //{
                            //    UtenteResponsabileTecnicoSost = "";
                            //}

                            try
                            {
                                UtenteCapoCommessaSost = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_55").ToString();
                            }
                            catch (Exception ex)
                            {
                                UtenteCapoCommessaSost = "";
                            }

                            CurrentCapoCommessa = UtenteCapoCommessa;

                            if (UtenteCapoCommessaSost != "")
                            {
                                CurrentCapoCommessa = UtenteCapoCommessaSost;
                            }

                            //if (UtenteResponsabileTecnicoSost != "")
                            //{
                            //    CurrentResponsabileTecnico = UtenteResponsabileTecnicoSost;
                            //    UtenteResponsabileTecnico = UtenteResponsabileTecnicoSost;
                            //}


                            ikr = knosObjectSelectorCertificati.GetItem(i).GetObjectDocuments();
                            if (ikr.HasErrors == false)
                            {

                                fileName = fileUrl = fileDescr = fileLocalPath = "";
                                fileIdDoc = 0;
                                if (knosObjectSelectorCertificati.GetItem(i).DocumentList.ItemCount == 1)
                                {
                                    fileName = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).FileName;
                                    fileUrl = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).GetUrl();
                                    fileIdDoc = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).IdDoc;
                                    fileDescr = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).FileDescr;

                                    // pulizia dei file locali
                                    fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);
                                    File.Delete(Path.Combine(fileLocalPath));

                                    ////download local del file
                                    //File.Delete(Path.Combine(Path.GetTempPath(), fileName));
                                    ////ikr = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).DownloadFile(Path.GetTempPath(), fileName);
                                    ////if (ikr.HasErrors == false)
                                    ////{
                                    ////}
                                }


                                if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatusIniziale)
                                {
                                    if (UtenteTecnico == CurrentUser)
                                    {
                                        nrCertificatiUtente1FDaFirmare += 1;
                                    }
                                }

                                //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                                //{
                                //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                                //    {
                                //        nrCertificatiUtente2FDaFirmare += 1;
                                //    }

                                //}

                                if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                                {
                                    nrCertificati1F += 1;

                                    if (UtenteTecnico == CurrentUser)
                                    {
                                        nrCertificatiUtente1F += 1;
                                    }
                                }


                                //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus2F)
                                //{
                                //    nrCertificati2F += 1;

                                //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                                //    {
                                //        nrCertificatiUtente2F += 1;
                                //    }

                                //}



                                dtCertificati.Rows.Add(
                                    knosObjectSelectorCertificati.GetItem(i).IdObject,
                                    knosObjectSelectorCertificati.GetItem(i).IdStatus,
                                    knosObjectSelectorCertificati.GetItem(i).StatusName,
                                    UtenteTecnico,
                                    DataPrimaFirma,
                                    //UtenteResponsabileTecnico,
                                    //DataSecondaFirma,
                                    UtenteCapoCommessa,
                                    fileName,
                                    fileUrl,
                                    fileLocalPath,
                                    fileIdDoc,
                                    fileDescr,
                                    //UtenteResponsabileTecnicoSost,
                                    UtenteCapoCommessaSost
                                    );


                            }


                            s.Text = string.Format("Caricamento certificati completato!", i.ToString(), nrCertificatiTot.ToString());

                            s.Text = string.Format("Caricamento utenti sostitutivi", i.ToString(), nrCertificatiTot.ToString());

                            ListViewItem li;
                            ListViewItem lx = null;

                            // recupero firme
                            ikr = knosObjectSelectorCertificati.GetItem(i).GetObjectLinks();
                            if (ikr.HasErrors == false)
                            {

                                for (int xLink = 0; xLink < knosObjectSelectorCertificati.GetItem(i).LinkList.ItemCount; xLink++)
                                {
                                    lx = new ListViewItem();
                                    li = new ListViewItem(knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).LinkDescr);

                                    if (lvFirme.Items.Count > 0)
                                    {
                                        lx = lvFirme.FindItemWithText(li.Text, false, 0);
                                    }
                                    else
                                    {
                                        lx = null;
                                    }

                                    if ((lx == null))
                                    {
                                        lvFirme.Items.Add(li);

                                        if (knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).Url.StartsWith("file:"))
                                        {
                                            lvFirme.Items[lvFirme.Items.Count - 1].SubItems.Add(knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).Url.ToString().Replace(@"file://", ""));

                                        }
                                        else
                                        {
                                            lvFirme.Items[lvFirme.Items.Count - 1].SubItems.Add(knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).Url.ToString());
                                        }

                                    }
                                }

                            }


                            dgCertificati.DataSource = dtCertificati;


                            if ((nrCertificati2F == nrCertificatiTot) && (CurrentCapoCommessa == CurrentUser))
                            {
                                SignFiles.tipofirma = 2;
                            }



                        }



                        retvalue = true;
                    }


                    else
                    {
                        MessageBox.Show(ikr.GetError(1).Description);
                    }

                    return retvalue;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore \r\n{0}", ex.Message));
                    retvalue = false;
                    return retvalue;

                }
            }


            public bool GetCertificatoPDLSelector(int _idObject, ListView lvAttr, DataGridView dgCertificati, ListView lvFirme, StatusStrip s)
            {
                bool retvalue = false;
                string fileName = "";
                string fileUrl = "";
                string fileLocalPath = "";
                int fileIdDoc = 0;
                string fileDescr = "";

                nrCertificati1F = nrCertificati2F = nrCertificatiTot = nrCertificatiUtente = nrCertificatiUtente1F = nrCertificatiUtente2F = nrCertificatiUtente1FDaFirmare = nrCertificatiUtente2FDaFirmare = 0;

                string UtenteTecnico = "";
                // string UtenteResponsabileTecnico = "";
                string UtenteCapoCommessa = "";
                string DataPrimaFirma = "";
                //string DataSecondaFirma = "";
                //string UtenteResponsabileTecnicoSost = "";
                string UtenteCapoCommessaSost = "";

                lvFirme.Clear();
                lvFirme.Columns.Clear();
                lvFirme.Columns.Add("Utente");
                lvFirme.Columns.Add("PathFileFirma");


                dgCertificati.DataSource = null;
                dgCertificati.Refresh();

                DataTable dtCertificati = new DataTable();
                dtCertificati.Columns.Add("IdObject");
                dtCertificati.Columns.Add("IdStatus");
                dtCertificati.Columns.Add("Status");
                dtCertificati.Columns.Add("Tecnico");
                dtCertificati.Columns.Add("DataPrimaFirma");
                //dtCertificati.Columns.Add("ResponsabileTecnico");
                //dtCertificati.Columns.Add("DataSecondaFirma");
                dtCertificati.Columns.Add("CapoCommessa");
                dtCertificati.Columns.Add("File");
                dtCertificati.Columns.Add("Url");
                dtCertificati.Columns.Add("LocalFile");
                dtCertificati.Columns.Add("IdDoc");
                dtCertificati.Columns.Add("FileDescr");
                //dtCertificati.Columns.Add("ResponsabileTecnicoSost");
                dtCertificati.Columns.Add("CapoCommessaSost");


                //KnosInstance.Client.Login(CurrentUser, "sash17ne", out cIdSubject);

                knosObject = KnosInstance.Client.CreateKnosObject();
                knosObjectCertificato = KnosInstance.Client.CreateKnosObject();
                knosObjectCliente = KnosInstance.Client.CreateKnosObject();

                try
                {

                    IKnosResult ikr;
                    IKnosObjectSelector knosObjectSelectorCertificati = KnosInstance.Client.CreateKnosObjectSelector();

                    knosObjectSelectorCertificati.SearchExpression = string.Format("IDClass = 47 and IdObject = {0}", SignFiles.startXML_idobject_certificato);
                    knosObjectSelectorCertificati.PageSize = 50;
                    knosObjectSelectorCertificati.SelectIdView = 125;
                    knosObjectSelectorCertificati.GetPage(1);


                    nrCertificatiTot = knosObjectSelectorCertificati.RecordCount;
                    for (int i = 0; i < nrCertificatiTot; i++)
                    {

                        s.Text = string.Format("Caricamento certificati in corso..... ({0}/{1})", i.ToString(), nrCertificatiTot.ToString());
                        // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                        // ci si è loggati a KnoS



                        UtenteTecnico = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_51").ToString();
                        try
                        {
                            DataPrimaFirma = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("datetime_08").ToString();
                        }
                        catch (Exception ex)
                        {
                            DataPrimaFirma = "";
                        }
                        //UtenteResponsabileTecnico = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_52").ToString();
                        //try
                        //{
                        //    DataSecondaFirma = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("datetime_09").ToString();
                        //}
                        //catch (Exception ex)
                        //{
                        //    DataSecondaFirma = "";
                        //}
                        UtenteCapoCommessa = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_53").ToString();
                        //try
                        //{
                        //    UtenteResponsabileTecnicoSost = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_54").ToString();
                        //}
                        //catch (Exception ex)
                        //{
                        //    UtenteResponsabileTecnicoSost = "";
                        //}
                        try
                        {
                            UtenteCapoCommessaSost = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_55").ToString();
                        }
                        catch (Exception ex)
                        {
                            UtenteCapoCommessaSost = "";
                        }

                        CurrentCapoCommessa = UtenteCapoCommessa;

                        if (UtenteCapoCommessaSost != "")
                        {
                            CurrentCapoCommessa = UtenteCapoCommessaSost;
                        }

                        ikr = knosObjectSelectorCertificati.GetItem(i).GetObjectDocuments();
                        if (ikr.HasErrors == false)
                        {

                            fileName = fileUrl = fileDescr = fileLocalPath = "";
                            fileIdDoc = 0;
                            if (knosObjectSelectorCertificati.GetItem(i).DocumentList.ItemCount == 1)
                            {
                                fileName = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).FileName;
                                fileUrl = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).GetUrl();
                                fileIdDoc = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).IdDoc;
                                fileDescr = knosObjectSelectorCertificati.GetItem(i).DocumentList.GetItem(0).FileDescr;

                                // pulizia dei file locali
                                fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);
                                File.Delete(Path.Combine(fileLocalPath));
                            }


                            if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatusIniziale)
                            {
                                if (UtenteTecnico == CurrentUser)
                                {
                                    nrCertificatiUtente1FDaFirmare += 1;
                                }
                            }

                            //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                            //{
                            //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                            //    {
                            //        nrCertificatiUtente2FDaFirmare += 1;
                            //    }

                            //}

                            if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                            {
                                nrCertificati1F += 1;

                                if (UtenteTecnico == CurrentUser)
                                {
                                    nrCertificatiUtente1F += 1;
                                }
                            }


                            //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus2F)
                            //{
                            //    nrCertificati2F += 1;

                            //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                            //    {
                            //        nrCertificatiUtente2F += 1;
                            //    }

                            //}



                            dtCertificati.Rows.Add(
                                knosObjectSelectorCertificati.GetItem(i).IdObject,
                                knosObjectSelectorCertificati.GetItem(i).IdStatus,
                                knosObjectSelectorCertificati.GetItem(i).StatusName,
                                UtenteTecnico,
                                DataPrimaFirma,
                                //UtenteResponsabileTecnico,
                                //DataSecondaFirma,
                                UtenteCapoCommessa,
                                fileName,
                                fileUrl,
                                fileLocalPath,
                                fileIdDoc,
                                fileDescr,
                                //UtenteResponsabileTecnicoSost,
                                UtenteCapoCommessaSost
                                );


                        }



                        ListViewItem li;
                        ListViewItem lx = null;

                        // recupero firme
                        ikr = knosObjectSelectorCertificati.GetItem(i).GetObjectLinks();
                        if (ikr.HasErrors == false)
                        {

                            for (int xLink = 0; xLink < knosObjectSelectorCertificati.GetItem(i).LinkList.ItemCount; xLink++)
                            {
                                lx = new ListViewItem();
                                li = new ListViewItem(knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).LinkDescr);

                                if (lvFirme.Items.Count > 0)
                                {
                                    lx = lvFirme.FindItemWithText(li.Text, false, 0);
                                }
                                else
                                {
                                    lx = null;
                                }

                                if ((lx == null))
                                {
                                    lvFirme.Items.Add(li);

                                    if (knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).Url.StartsWith("file:"))
                                    {
                                        lvFirme.Items[lvFirme.Items.Count - 1].SubItems.Add(knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).Url.ToString().Replace(@"file://", ""));

                                    }
                                    else
                                    {
                                        lvFirme.Items[lvFirme.Items.Count - 1].SubItems.Add(knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(xLink).Url.ToString());
                                    }

                                }
                            }

                        }




                        dgCertificati.DataSource = dtCertificati;


                        if ((nrCertificati2F == nrCertificatiTot) && (CurrentCapoCommessa == CurrentUser))
                        {
                            SignFiles.tipofirma = 2;
                        }

                        retvalue = true;
                    }

                    return retvalue;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore \r\n{0}", ex.Message));
                    retvalue = false;
                    return retvalue;

                }
            }


            public List<Allegato> GetAllegatiPubblicazione(int _idObject)
            {
                bool retvalue = false;
                string fileName = "";
                string fileUrl = "";
                string fileLocalPath = "";
                int fileIdDoc = 0;
                string fileDescr = "";

                knosObject = KnosInstance.Client.CreateKnosObject();

                List<Allegato> _allegati = new List<Allegato>();

                try
                {
                    IKnosResult ikr = knosObject.GetObjectLinks(_idObject);

                    if (ikr.HasErrors == false)
                    {

                        for (int i = 0; i < knosObject.LinkList.ItemCount; i++)
                        {
                            fileDescr = knosObject.LinkList.GetItem(i).LinkDescr;
                            fileUrl = knosObject.LinkList.GetItem(i).Url;
                            fileName = knosObject.LinkList.GetItem(i).Url;

                            Allegato a = new Allegato(fileDescr, fileDescr, fileUrl.Replace("file://", ""));

                            _allegati.Add(a);
                        }

                    }

                }

                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore \r\n{0}", ex.Message));

                }

                return _allegati;
            }


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


                    ikr = knosObjectMaker.UploadFiles(_idObject, out kui);

                    if (ikr.HasErrors == false)
                    {
                        retvalue = true;
                    }
                    else
                    {
                        c = Cursors.Default;
                        MessageBox.Show(ikr.GetError(0).Description);
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


            public bool DeleteFiles(int _idObject,
                int _idDoc
                )
            {
                bool retvalue = false;
                IKnosUploadInfo kui;

                // delete file
                IKnosObject knosObject = KnosInstance.Client.CreateKnosObject();
                IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


                Cursor c;

                c = Cursors.WaitCursor;

                IKnosResult ikr = knosObject.GetObjectDocuments(_idObject);

                if (ikr.HasErrors == false)
                {
                    if (_idDoc >= 0)
                    {
                        // cancello il documento attualmente presente
                        ikr = knosObjectMaker.DeleteDoc(_idObject, _idDoc);
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

                    }
                    else
                    {
                        for (int i = 0; i < knosObject.DocumentList.ItemCount; i++)
                        {
                            int iddoc = 0;
                            int.TryParse(knosObject.DocumentList.GetItem(i).IdDoc.ToString(), out iddoc);

                            // cancello il documento attualmente presente
                            ikr = knosObjectMaker.DeleteDoc(_idObject, iddoc);
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
                        }

                    }
                }

                return retvalue;


            }

            public bool EseguiAzione(int _idObject,
                    int _actionWF,
                    string _attrNameDate
            )
            {
                bool retvalue = false;
                IKnosUploadInfo kui;

                // upload file
                IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


                Cursor c;

                c = Cursors.WaitCursor;

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

                    if (_attrNameDate != "")
                    {

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
                }

                return retvalue;

            }



            public bool EseguiAzioneWS(int idObject,
                    int _actionWF,
                    string _attrNameDate)
            {


                bool bOK = false;
                IKnosResult result;


                //Inizializza(KnosInstance.Client.KnosBaseUrl);

                IKnosRequest request = KnosInstance.Client.CreateKnosRequest();



                request.SetParameter("IdObject", idObject.ToString());
                request.SetParameter("IdAction", _actionWF.ToString());
                request.SetParameter("IgnoreError", "1");
                request.SetParameter("SkipAction", "0");
                request.SetParameter("SkipNotify", "2");
                request.SetParameter("CheckObjectUnlock", "0");
                IKnosResponse response;

                result = KnosInstance.Client.ParseResponse(string.Format("{0}/knos/system/webservices/object_changestatusbyaction.asp", KnosInstance.Client.KnosBaseUrl), ref request, out response);

                if (result.NoErrors == true)
                {
                    bOK = true;// Pubblicazione bloccata, si può elaborare
                }
                //else
                //{
                //    ;// Pubblicazione non bloccata, si deve saltare 
                //}

                return bOK;

            }

            public bool EliminaAllegato(int _idObject,
                                int _idDoc,
                                string _attrNameDate
                        )
            {
                bool retvalue = false;
                IKnosUploadInfo kui;

                // upload file
                IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();


                Cursor c;

                c = Cursors.WaitCursor;

                if ((_idObject > 0) && (_idDoc > 0))
                {
                    retvalue = true;

                    // altrimenti faccio la transizione di stato della pubblicazione certificato

                    IKnosResult knosResult = knosObjectMaker.DeleteDoc(_idObject, _idDoc);

                    // se qualcosa è andato storto esco
                    if (knosResult.HasWarningsErrors)
                    {
                        return retvalue;
                    }

                }

                return retvalue;

            }




            public bool downloadDoc(int _idCertificato, int _idDoc = 1, string filePath = "")
            {
                IKnosResult ikr;
                bool bOK = false;

                IKnosObject knosObject = KnosInstance.Client.CreateKnosObject();

                ikr = knosObject.GetObjectDocuments(_idCertificato);
                if (ikr.HasErrors == false)
                {
                    //download local del file
                    for (int i = 0; i < knosObject.DocumentList.ItemCount; i++)
                    {
                        if (_idDoc == knosObject.DocumentList.GetItem(i).IdDoc)
                        {
                            ikr = knosObject.DocumentList.GetItem(i).DownloadFile(Path.GetTempPath(), filePath);
                            break;
                        }
                    }
                }

                if (ikr.HasErrors == false)
                {
                    return true;
                }
                else
                {
                    MessageBox.Show(string.Format("{0}\\{1} \r\n {2}", Path.GetTempPath(), filePath, ikr.GetError(0).Description), "Errore in Download allegato");
                    return false;
                }
            }


            public bool downloadDoc(int _idCertificato, int _idDoc = 1, string filePath = "", string filename = "")
            {
                IKnosResult ikr;
                bool bOK = false;

                IKnosObject knosObject = KnosInstance.Client.CreateKnosObject();

                ikr = knosObject.GetObjectDocuments(_idCertificato);
                if (ikr.HasErrors == false)
                {
                    //download local del file
                    for (int i = 0; i < knosObject.DocumentList.ItemCount; i++)
                    {
                        if (_idDoc == knosObject.DocumentList.GetItem(i).IdDoc)
                        {
                            ikr = knosObject.DocumentList.GetItem(i).DownloadFile(filePath, filename);
                            break;
                        }
                    }
                }

                if (ikr.HasErrors == false)
                {
                    return true;
                }
                else
                {
                    MessageBox.Show(string.Format("{0}\\{1} \r\n {2}", filePath, filename, ikr.GetError(0).Description), "Errore in Download allegato");
                    return false;
                }
            }


            public bool SetSostituto(int _idObject, string _column_name, string _utente)
            {
                IKnosResult ikr;
                bool bOK = false;

                IKnosObjectMaker ko = KnosInstance.Client.CreateKnosObjectMaker();



                ikr = ko.SetAttrValue(_column_name, _utente);
                if (ikr.HasErrors == false)
                {
                    //aggiorno la pubblicazione certificato
                    ikr.ClearAll();
                    ikr = ko.UpdateObject(_idObject);
                }

                if (ikr.HasErrors == false)
                {
                    return true;
                }
                else
                {
                    MessageBox.Show(ikr.GetError(0).ToString(), "Errore aggiornamento certificato");
                    return false;
                }
            }


            public bool AddDestinatarioCapoCommessa(int _idObject)
            {
                bool bAddDestinatario = false;
                int _idSubject = 0;
                bool retvalue = false;

                _idSubject = GetIdSubjectByName(ToDoNotificheBSC.frmToDoNotificheBSC.CurrentCapoCommessa);

                if (_idSubject > 0)
                {
                    IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();
                    IKnosRecipientEditor kRE = knosObjectMaker.RecipientEditor;

                    IKnosResult ikr = knosObject.GetAllObjectData(_idObject);

                    if (ikr.HasErrors == false)
                    {

                        IKnosRecipientList knosRecipientList = knosObject.RecipientList;

                        for (int i = 0; i < knosRecipientList.ItemCount; i++)
                        {
                            if (knosRecipientList.GetItem(i).IdSubject == _idSubject)
                            {
                                bAddDestinatario = false;
                                break;
                            }

                        }
                    }

                    if (bAddDestinatario)
                    {
                        IKnosRecipient knosRecipient = KnosInstance.Client.CreateKnosRecipient();
                        knosRecipient.IdSubject = _idSubject;

                        kRE.AddValue(knosRecipient);

                        IKnosResult knosResult = knosObjectMaker.UpdateObject(_idObject);

                        if (knosResult.HasWarningsErrors)
                        {
                            MessageBox.Show(knosResult.ToString());
                        }
                        else
                        {
                            retvalue = true;
                        }

                    }

                }

                return retvalue;


            }


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


            public bool GetSignImage(int _idObject, ListView lvFirme, string _signer)
            {
                bool retvalue = false;

                foreach (ListViewItem li in lvFirme.Items)
                {
                    if (li.Text == _signer)
                    {
                        SignFiles.filePNG = (li.SubItems[1].Text);
                        retvalue = true;
                        break;
                    }

                }


                //lvFirme.Clear();
                //lvFirme.Columns.Clear();
                //lvFirme.Columns.Add("Utente");
                //lvFirme.Columns.Add("PathFileFirma");


                //knosObjectCertificato = KnosInstance.Client.CreateKnosObject();

                //IKnosResult ikr = knosObjectCertificato.GetObjectLinks(_idObject);

                //if (ikr.HasErrors == false)
                //{

                //    for (int i = 0; i < knosObjectCertificato.LinkList.ItemCount; i++)
                //    {
                //        lvFirme.Items.Add(knosObjectCertificato.LinkList.GetItem(i).LinkDescr);


                //        if (knosObjectCertificato.LinkList.GetItem(i).Url.StartsWith("file:"))
                //        { 
                //            lvFirme.Items[i].SubItems.Add(knosObjectCertificato.LinkList.GetItem(i).Url.ToString().Replace(@"file://", ""));

                //        }


                //        if (knosObjectCertificato.LinkList.GetItem(i).LinkDescr == _signer)
                //        {
                //            SignFiles.filePNG = (knosObjectCertificato.LinkList.GetItem(i).Url.ToString().Replace(@"file://", ""));
                //            retvalue = true;
                //        }

                //    }


                //}
                //else
                //{
                //    MessageBox.Show(ikr.GetError(0).Description);
                //}

                return retvalue;


            }


            public DataTable GetMyCertificates(string whereAgg)
            {
                string Tecnico = "";
                string DataPrimaFirma = "";
                //string ResponsabileTecnico = "";
                //string DataSecondaFirma = "";
                string CapoCommessa = "";
                string File = "";
                string Url = "";
                string LocalFile = "";
                string IdDoc = "";
                string FileDescr = "";
                //string ResponsabileTecnicoSost = "";
                string CapoCommessaSost = "";
                string IdPDL = "";
                string DatiPDL = "";
                string ClientePDL = "";

                string fileName = "";
                string fileUrl = "";
                string fileLocalPath = "";
                int fileIdDoc = 0;
                string fileDescr = "";

                string DataChiusuraPDL = "";

                DataTable dtCertificati = new DataTable();
                dtCertificati.Columns.Add("IdObject");

                Boolean bFile = false;


                IKnosObjectSelector knosObjectSelector = KnosInstance.Client.CreateKnosObjectSelector();

                knosObjectSelector.SearchExpression = string.Format("IdClass = 5012 AND exists(SELECT 1 FROM OBJECT_DOC WHERE FILENAME LIKE " + whereAgg);
                knosObjectSelector.SelectIdView = 125;
                knosObjectSelector.PageSize = 100;

                knosObjectSelector.GetPage(1);

                int nRec = knosObjectSelector.PageSize;

                if (knosObjectSelector.RecordCount < nRec)
                {
                    nRec = knosObjectSelector.RecordCount;
                }

                for (int i = 0; i < nRec; i++)
                {
                    DataPrimaFirma = CapoCommessaSost = ""; //DataSecondaFirma = ResponsabileTecnicoSost = 

                    try
                    {
                        Tecnico = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("varchar_51").ToString();
                    }
                    catch
                    {
                    }

                    //try
                    //{
                    //    ResponsabileTecnico = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("varchar_52").ToString();
                    //}
                    //catch
                    //{
                    //}

                    try
                    {
                        CapoCommessa = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("varchar_53").ToString();
                    }
                    catch
                    {
                    }

                    try
                    {
                        DataPrimaFirma = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("datetime_08").ToString();
                    }
                    catch
                    {
                    }

                    //try
                    //{
                    //    DataSecondaFirma = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("datetime_09").ToString();
                    //}
                    //catch
                    //{
                    //}

                    //try
                    //{
                    //    ResponsabileTecnicoSost = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("varchar_54").ToString();
                    //}
                    //catch
                    //{
                    //}

                    try
                    {
                        CapoCommessaSost = knosObjectSelector.GetItem(i).AttrValueList.GetItemByColumnName("varchar_55").ToString();
                    }
                    catch
                    {
                    }

                    //IdPDL
                    try
                    {
                        IdPDL = knosObjectSelector.GetItem(i).AttrValueList.GetItemByName("IdPDL").ToString();
                    }
                    catch
                    {
                    }

                    //Dati PDL
                    try
                    {
                        DatiPDL = knosObjectSelector.GetItem(i).AttrValueList.GetItemByName("DatiPDL").ToString();
                    }
                    catch
                    {
                    }

                    //Cliente PDL
                    try
                    {
                        ClientePDL = knosObjectSelector.GetItem(i).AttrValueList.GetItemByName("ClientePDL").ToString();
                    }
                    catch
                    {
                    }

                    //IKnosResult ikr = knosObjectSelector.GetItem(i).GetObjectDocuments();
                    //if (ikr.HasErrors == false)
                    //{

                    //    fileName = fileUrl = fileDescr = fileLocalPath = "";
                    //    fileIdDoc = 0;
                    //    if (knosObjectSelector.GetItem(i).DocumentList.ItemCount == 1)
                    //    {
                    //        fileName = knosObjectSelector.GetItem(i).DocumentList.GetItem(0).FileName;
                    //        fileUrl = knosObjectSelector.GetItem(i).DocumentList.GetItem(0).GetUrl();
                    //        fileIdDoc = knosObjectSelector.GetItem(i).DocumentList.GetItem(0).IdDoc;
                    //        fileDescr = knosObjectSelector.GetItem(i).DocumentList.GetItem(0).FileDescr;

                    //        // pulizia dei file locali
                    //        fileLocalPath = "";//Path.Combine(Path.GetTempPath(), fileName);
                    //        //File.Delete(Path.Combine(fileLocalPath));
                    //    }
                    //}

                    try
                    {
                        DataChiusuraPDL = knosObjectSelector.GetItem(i).AttrValueList.GetItemByName("DataChiusuraPDL").ToString();

                        DateTime x = new DateTime();
                        DateTime.TryParse(DataChiusuraPDL, out x);

                        if (DataChiusuraPDL.StartsWith("1900"))
                        {
                            DataChiusuraPDL = "";
                        }
                        else
                        {
                            DataChiusuraPDL = x.ToShortDateString();
                        }
                    }
                    catch
                    {
                    }

                    bFile = false;

                    IKnosResult ikr = knosObjectSelector.GetItem(i).GetObjectLinks();
                    if (ikr.HasErrors == false)
                    {
                        fileName = fileUrl = fileDescr = fileLocalPath = "";
                        fileIdDoc = 0;
                        if (knosObjectSelector.GetItem(i).LinkList.ItemCount == 1)
                        {
                            bFile = true;

                            fileName = knosObjectSelector.GetItem(i).LinkList.GetItem(0).Url;
                            fileUrl = knosObjectSelector.GetItem(i).LinkList.GetItem(0).Url;
                            fileIdDoc = knosObjectSelector.GetItem(i).LinkList.GetItem(0).IdLink;
                            fileDescr = knosObjectSelector.GetItem(i).LinkList.GetItem(0).LinkDescr;

                            // pulizia dei file locali
                            fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);

                        }

                    }

                    int IdAction = 0;

                    switch (knosObjectSelector.GetItem(i).IdStatus.ToString())
                    {
                        case "159":
                            {
                                bFile = true;
                                break;
                            }

                        case "160":
                            {
                                IdAction = 240;
                                break;
                            }

                        case "161":
                            {
                                IdAction = 239;
                                break;
                            }
                    }


                    if (bFile)
                    {
                        dtCertificati.Rows.Add(
                            IdAction.ToString(),
                            knosObjectSelector.GetItem(i).IdObject.ToString(),
                            knosObjectSelector.GetItem(i).IdStatus.ToString(),
                            knosObjectSelector.GetItem(i).StatusName.ToString(),
                            Tecnico,
                            DataPrimaFirma,
                            //ResponsabileTecnico,
                            //DataSecondaFirma,
                            CapoCommessa,
                            fileName,
                            fileUrl,
                            //fileLocalPath,
                            fileIdDoc,
                            fileDescr,
                            //ResponsabileTecnicoSost,
                            CapoCommessaSost,
                            IdPDL,
                            DatiPDL,
                            ClientePDL,
                            DataChiusuraPDL);
                    }
                }



                return dtCertificati;


            }



            public int StoreEmailSent(int idClass
                , string Tipo
                , string IdObjectCli
                , DateTime Data
                , string Da
                , string A
                , string CC
                , string CCN
                , string Title
                , string Testo
                , string Note
                , List<int> links
                , List<string> files
                , bool SS = false
                , bool COA = false
                , bool DOCMETODO = false
                )
            {
                int idobjectMail = 0;
                int idObjectCLI = 0;

                IKnosObjectMaker kom = KnosInstance.Client.CreateKnosObjectMaker();

                kom.Reset();

                kom.IdClass = idClass;
                kom.SetAttrValue("enum_27", Tipo);

                IKnosMultivalueEditor kme = KnosInstance.Client.CreateKnosMultivalueEditor();
                int.TryParse(IdObjectCli, out idObjectCLI);
                kme.AddValue(idObjectCLI);

                if (links != null)
                {
                    foreach (int l in links)
                    {
                        IKnosLink ikl = KnosInstance.Client.CreateKnosLink();
                        ikl.IdObjectTo = l;
                        kom.LinkEditor.AddValue(ikl);
                    }
                }

                kom.SetAttrValue("object_19", kme);
                kom.SetAttrValue("datetime_01", Data);
                kom.SetAttrValue("text_03", Da);
                kom.SetAttrValue("text_04", A);
                kom.SetAttrValue("text_05", CC);
                kom.SetAttrValue("text_06", CCN);
                kom.SetAttrValue("title", Title);
                kom.SetAttrValue("text_01", Testo);
                kom.SetAttrValue("text_02", Note);
                kom.SetAttrValue("smallint_01", SS);
                kom.SetAttrValue("smallint_02", COA);
                kom.SetAttrValue("smallint_04", DOCMETODO);


                IKnosResult kr = kom.CreateObject(out idobjectMail);

                if (idobjectMail > 0)
                {
                    if (files != null)
                    {
                        if (files.Count > 0)
                        {
                            foreach (string a in files)
                            {
                                //IKnosUploadItem kui = KnosInstance.Client.CreateKnosUploadItem();
                                //kui.FilePath = a;
                                //kui.FileName = Path.GetFileName(a);

                                //kom.AddUploadItem(kui);

                                UploadFileCertificato(idobjectMail, 0, a, Path.GetFileName(a), Path.GetFileName(a), 0, "", "");


                            }

                            //kom.UpdateObject(idobjectMail);

                            //kr = kom.CreateObject(out idobjectMail);
                        }
                    }

                }

                return idobjectMail;


            }

        }

        string nomeRTF = "";
        string nomePDF = "";
        bool opened = false;
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



        public frmToDoNotificheBSC()
        {
            InitializeComponent();
        }




        private void frmToDoNotificheBSC_Load(object sender, EventArgs e)
        {
            chkFilesDaDB.Visible = false;
            if (Properties.Settings.Default.SqlSchede != "")
            {
                chkFilesDaDB.Visible = true;
            }

            string pathCheckMessage = "";


            log = new Logger();
            log.Setup();
            log.LogSomething("Start servizio");

            clsRadGridSettings.log = log;
            clsRadGridSettings.GetColumnsSettings(radGridView1, s, "SchedeSicurezza", cmbImpostazioni);

            this.Text = string.Format("ZSI - Gestione notifiche documenti ({0})", Application.ProductVersion);

            txtUltimoAggiornamento.Text = Properties.Settings.Default.DataAggiornamento.ToShortDateString();
            txtUltimoInvio.Text = Properties.Settings.Default.DataInvioSchede.ToShortDateString();

            try
            {

                //// start path per i PDF da importare
                //if (Directory.Exists(Properties.Settings.Default.PathPDF))
                //    lblPathEpy.Text = Properties.Settings.Default.PathPDF;

                //if (Directory.Exists(Properties.Settings.Default.PathPDF_H))
                //    lblPathH.Text = Properties.Settings.Default.PathPDF_H;

                ////if (Directory.Exists(Properties.Settings.Default.PathRMI))
                ////    lblPathRMI.Text = Properties.Settings.Default.PathRMI;

                foreach (string p in Properties.Settings.Default.listPathPDF)
                {
                    if (Directory.Exists(p))
                    {

                    }
                    else
                    {
                        pathCheckMessage += "\r\n percorso File Epy non trovato: " + p;
                    }

                    cmbPathEpy.Items.Add(p);

                }

                if (cmbPathEpy.Items.Count > 0)
                    cmbPathEpy.SelectedIndex = 0;


                foreach (string p in Properties.Settings.Default.listPathPDF_H)
                {
                    if (Directory.Exists(p))
                    {

                    }
                    else
                    {
                        pathCheckMessage += "\r\n percorso File Epy (da file system) non trovato: " + p;
                    }

                    cmbPathH.Items.Add(p);

                }

                if (cmbPathH.Items.Count > 0)
                    cmbPathH.SelectedIndex = 0;











                foreach (SettingsProperty currentProperty in Properties.Settings.Default.Properties)
                {
                    textBoxSettings.Text += string.Format("\r\n {0} - {1}", currentProperty.Name, currentProperty.DefaultValue);
                }

                opened = false;

                bool.TryParse(Properties.Settings.Default.sendMailPopUp, out notifyPopUp);



            }
            catch (Exception ex)
            {
                lblPNGFirma.Text = "";
                //return;
            }


            // caricamento automatico schede da epy
            if (Properties.Settings.Default.RicercaAutomaticaSchede)
            {
                txtUltimoAggiornamento.Text = System.DateTime.Today.ToShortDateString();
                btnOldArchivio_Click(null, null);
            }


            textBoxLOG.Text += string.Format("\r\nApplication.LocalUserAppDataPath {0}", Application.LocalUserAppDataPath);
            textBoxLOG.Text += string.Format("\r\nApplication.UserAppDataPath {0}", Application.UserAppDataPath);
            textBoxLOG.Text += string.Format("\r\nApplication.UserAppDataRegistry {0}", Application.UserAppDataRegistry);

            dTP_BOLLEDA.Value = System.DateTime.Today;
            dTP_COADA.Value = System.DateTime.Today.AddMonths(-6);


            //Knos
            refreshKnosLogin();

            Application.DoEvents();

            this.WindowState = FormWindowState.Maximized;

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

            }


            if (tabControl1.SelectedIndex == 1)
            {

            }
        }

        private void txtIdPDL_Leave(object sender, EventArgs e)
        {

        }

        private void LoadPDL()
        {
            //int _intRes = 0;

            ////webBrowser1.Navigate("about:blank");
            ////axAcroPDF1.Dispose();



            //if (int.TryParse(txtIdPDL.Text, out _intRes) == true)
            //{
            //    _cursor = Cursors.WaitCursor;


            //    CurrentIdObject = _intRes;
            //    toolStripStatusLabel1.Text = "Caricamento dati e certificati del PDL.....";

            //    //if (SignFiles.startXML_idobject_certificato > 0)
            //    //{
            //    //    kw.GetCertificatoPDLSelector(SignFiles.startXML_idobject_certificato, listViewAttr, dataGridViewCertificati, lvFileFirma, statusStrip1);
            //    //}
            //    //else
            //    //{
            //        kw.GetPDLSelector(_intRes, listViewAttr, dataGridViewCertificati, lvFileFirma, statusStrip1);
            //    //}
            //    // bloccaggio ordinamento colonne
            //    if (dataGridViewCertificati.Rows.Count > 0)
            //    {
            //        foreach (DataGridViewColumn dvc in dataGridViewCertificati.Columns)
            //        {
            //            dvc.SortMode = DataGridViewColumnSortMode.NotSortable;
            //        }
            //        dataGridViewCertificati.Sort(dataGridViewCertificati.Columns["IdObject"], ListSortDirection.Ascending);
            //    }
            //    // stato PDL
            //    btnPDLStatus.Text = CurrentStatusNamePDL;
            //    btnPDLStatus.Tag = CurrentPDFPDLUrl;

            //    btnFirmaCapoCommessa.Enabled = (kw.CurrentUser == CurrentCapoCommessa);

            //    //GetAzioneCertificati();

            //    if (SignFiles.tipofirma > 0)
            //    {
            //        tabControl1.SelectedIndex = 1;
            //    }
            //}
        }


        private void LoadCertificatoPDL()
        {
            //int _intRes = 0;

            //_cursor = Cursors.WaitCursor;

            //CurrentIdObject = _intRes;
            //toolStripStatusLabel1.Text = "Caricamento dati e certificati del PDL.....";
            //kw.GetCertificatoPDLSelector(_intRes, listViewAttr, dataGridViewCertificati, lvFileFirma, statusStrip1);

            //// bloccaggio ordinamento colonne
            //if (dataGridViewCertificati.Rows.Count > 0)
            //{
            //    foreach (DataGridViewColumn dvc in dataGridViewCertificati.Columns)
            //    {
            //        dvc.SortMode = DataGridViewColumnSortMode.NotSortable;
            //    }
            //    dataGridViewCertificati.Sort(dataGridViewCertificati.Columns["IdObject"], ListSortDirection.Ascending);
            //}

            //// stato PDL
            //btnPDLStatus.Text = CurrentStatusNamePDL;
            //btnPDLStatus.Tag = CurrentPDFPDLUrl;

            //btnFirmaCapoCommessa.Enabled = (kw.CurrentUser == CurrentCapoCommessa);

            ////GetAzioneCertificati();

            //if (SignFiles.tipofirma > 0)
            //{
            //    tabControl1.SelectedIndex = 1;
            //}

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


        private void btnMyCertificates_Click(object sender, EventArgs e)
        {
            string strW = "";

            //// check nuovi clienti
            //int nrNC = nrNuoviClienti();

            //if (nrNC > 0)
            //{
            //    if (MessageBox.Show(string.Format("Ci sono {0} articoli venduti a clienti che hanno effettuato il primo acquisto, vuoi visionarli?", nrNC), "Controllo Nuovi Clienti", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
            //    {
            //        chkNC.Checked = false;
            //    }
            //    else
            //    {
            //        chkNC.Checked = true;
            //    }

            //}


            aggiornaRegistroPA();



            string commandtext = Properties.Settings.Default.MetodoCommand;

            //commandtext += " WHERE DATADOC >= @DATADOC";

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {

                try
                {
                    cn.Open();

                    commandtext += " WHERE 1=1";

                    if ((checkBoxBOL.Checked) && (checkBoxSCH.Checked))
                    {
                        commandtext += string.Format(" AND (IDOBJECT_BOL > 0 OR IDOBJECT_SCH > 0)");
                    }
                    else
                    {
                        if (checkBoxBOL.Checked)
                        {
                            commandtext += string.Format(" AND IDOBJECT_BOL > 0 ");
                        }
                        if (checkBoxSCH.Checked)
                        {
                            commandtext += string.Format(" AND IDOBJECT_SCH > 0");
                        }
                    }


                    if (checkBoxNONSpedibile.Checked)
                    {
                        commandtext += string.Format(" AND EMAIL_CLIENTE = '' ");
                    }
                    else
                    {
                        commandtext += string.Format(" AND EMAIL_CLIENTE <> '' ");
                    }

                    if ((checkBoxSpediti.Checked) && (checkBoxDaSpedire.Checked))
                    { 
                    
                    }
                    else
                    {
                        if (checkBoxSpediti.Checked)
                        {
                            commandtext += string.Format(" AND NOT DATAULTIMOINVIO_SCH IS NULL ");
                        }

                        if (checkBoxDaSpedire.Checked)
                        {
                            commandtext += string.Format(" AND DATAULTIMOINVIO_SCH IS NULL ");
                        }


                    }


                    //if (checkBoxSpediti.Checked)
                    //{
                    //    //commandtext += string.Format(" AND NOT DATAULTIMOINVIO_SCH IS NULL ");
                    //}
                    //else
                    //{
                    //    commandtext += string.Format(" AND DATAULTIMOINVIO_SCH IS NULL ");
                    //}

                    if (chkStorico.Checked == false)
                    {
                        commandtext += string.Format(" AND ANNOULTIMADDT > 0 ");
                    }

                    if (chkNC.Checked)
                    {
                        commandtext += string.Format(" AND (NUOVOACQUISTO = 1)");
                    }

                    toolStripStatusLabel1.Text = string.Format("caricamento dati in corso..........");

                    radGridView1.EnableFiltering = false;
                    radGridView1.ShowFilteringRow = false;

                    using (SqlCommand cmd = new SqlCommand(commandtext, cn))
                    {
                        log.LogSomething(commandtext);

                        //cmd.Parameters.AddWithValue("DATADOC", dateTimePickerDa.Value);
                        //cmd.ExecuteNonQuery();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        radGridView1.DataSource = dt;

                        for (int i = 0; i < radGridView1.Columns.Count; i++)
                        {
                            if (radGridView1.Columns[i].FieldName.StartsWith("IDOBJECT"))
                            {
                                radGridView1.Columns[i].IsVisible = false;
                            }
                            else
                            {
                                radGridView1.Columns[i].BestFit();
                            }
                        }

                        //toolStripStatusLabel1.Text = string.Format("raggruppamento per cliente..........");

                        //groupclienti();

                        //toolStripStatusLabel1.Text = string.Format("raggruppamento per cliente completato");

                        radGridView1.AutoScroll = true;
                        radGridView1.Refresh();

                        toolStripStatusLabel1.Text = string.Format("Caricamento completato");

                    }


                    radGridView1.EnableFiltering = true;
                    radGridView1.ShowFilteringRow = true;
                    radGridView1.EnableAlternatingRowColor = true;
                    radGridView1.MultiSelect = true;
                    

                    if (radGridView1.Rows.Count() > 0)
                    {
                        getImpostazioniGriglia(radGridView1);
                    }

                    radGridView1.AutoSizeColumnsMode = Telerik.WinControls.UI.GridViewAutoSizeColumnsMode.None;
                    radGridView1.MasterTemplate.BestFitColumns();                    
                    
                    //Telerik.WinControls.Data.FilterDescriptor filter = new Telerik.WinControls.Data.FilterDescriptor();
                    //filter.PropertyName = "RAGIONESOCIALE";
                    //filter.Operator = Telerik.WinControls.Data.FilterOperator.Contains;
                    //filter.Value = "*";
                    //filter.IsFilterEditor = true;
                    //radGridView1.Columns["RAGIONESOCIALE"].FilterDescriptor = filter;
                    //                    radGridView1.FilterDescriptors.Add(filter);


                }

                catch (SqlException ex)
                {
                    MessageBox.Show(string.Format("Errore SQL SERVER: {0} - {1}", Properties.Settings.Default.MetodoConnectionString, ex.Message));

                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore : {0}", ex.Message));
                }


            }


            Cursor.Current = Cursors.WaitCursor;

            //LoadGridSettings(radGridView1);

            //SaveGridSettings(dataGridViewMyCertificates);

            //txtSearch_Search.Text = strW;

            //dataGridViewMyCertificates.DataSource = kw.GetMyCertificates(strW);

            //LoadGridSettings(dataGridViewMyCertificates);

            //if (dataGridViewMyCertificates.Rows.Count > 0)
            //{
            //    //((DataGridViewButtonColumn)dataGridViewMyCertificates.Columns[dataGridViewMyCertificates.Columns["dataGridViewButtonColumn1"].Index]).DefaultCellStyle.ForeColor = Color.Silver;
            //    //((DataGridViewButtonColumn)dataGridViewMyCertificates.Columns[dataGridViewMyCertificates.Columns["dataGridViewButtonColumn1"].Index]).DefaultCellStyle.BackColor = Color.Silver;
            //}

            Cursor.Current = Cursors.Default;

            lblPNGFirma.Text = string.Format("Nr articoli trovati: {0}", radGridView1.Rows.Count.ToString());
        }

        //private int nrNuoviClienti()
        //{
        //    string strW = "";
        //    int n = 0;

        //    string commandtext = Properties.Settings.Default.MetodoCommand;

        //    //commandtext += " WHERE DATADOC >= @DATADOC";

        //    using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
        //    {
        //        commandtext += " WHERE NUOVOCLIENTE = 1";

        //        using (SqlCommand cmd = new SqlCommand(commandtext, cn))
        //        {
                    
        //            //cmd.Parameters.AddWithValue("DATADOC", dateTimePickerDa.Value);
        //            //cmd.ExecuteNonQuery();
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            DataTable dt = new DataTable();
        //            da.Fill(dt);

        //            n = dt.Rows.Count;
        //        }
        //    }

        //    return n;
        //}
        private void btnSchedaPDL_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
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

        private void checkBoxALLUsers_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxALLUsers.Checked)
            {
                checkBoxALLUsers.BackColor = Color.Yellow;
            }
            else
            {
                checkBoxALLUsers.BackColor = Control.DefaultBackColor;
            }
        }

        private void SaveGridSettings(DataGridView dg)
        {
            // salva le impostazioni della gridview in un file XML per utente
            string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            DataTable dt = new DataTable("table");

            var query = from DataGridViewColumn col in dg.Columns
                        orderby col.DisplayIndex
                        select col;

            foreach (DataGridViewColumn col in query)
            {
                dt.Columns.Add(col.Name);
            }

            dt.WriteXmlSchema(pathGridSettings);
        }


        //private void LoadGridSettings(DataGridView dg)
        //{
        //    // salva le impostazioni della gridview in un file XML per utente
        //    string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

        //    DataTable dt = new DataTable();
        //    dt.ReadXmlSchema(pathGridSettings);

        //    int i = 0;
        //    foreach (DataColumn col in dt.Columns)
        //    {
        //        dataGridViewMyCertificates.Columns[col.ColumnName].DisplayIndex = i;
        //        i++;
        //    }        
        
        //}


        private void SaveGridSettings(Telerik.WinControls.UI.RadGridView dg)
        {
            // salva le impostazioni della gridview in un file XML per utente
            string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            dg.SaveLayout(pathGridSettings);
        }


        //private void LoadGridSettings(Telerik.WinControls.UI.RadGridView dg)
        //{
        //    // salva le impostazioni della gridview in un file XML per utente
        //    string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

        //    try
        //    {
        //        DataTable dt = new DataTable();
        //        dt.ReadXmlSchema(pathGridSettings);

        //        int i = 0;
        //        foreach (DataColumn col in dt.Columns)
        //        {
        //            dg.Columns[col.ColumnName].DisplayIndex = i;
        //            i++;
        //        }
        //    }
        //    catch { }
        //}




        void getImpostazioniGriglia(Telerik.WinControls.UI.RadGridView dg)
        {

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            string[] fi = Directory.GetFiles(s);

            string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            try
            {
                if (File.Exists(pathGridSettings))
                { 
                    dg.LoadLayout(pathGridSettings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delle righe ordine {0}", ex.Message));
            }

        }

        private void tabControl1_DoubleClick(object sender, EventArgs e)
        {
            txtSearch_Search.Visible = !(txtSearch_Search.Visible);
        }

        private void groupclienti()
        {
            radGridView1.EnableGrouping = true;
            radGridView1.GroupDescriptors.Clear();
            Telerik.WinControls.Data.GroupDescriptor descriptor = new Telerik.WinControls.Data.GroupDescriptor();
            descriptor.GroupNames.Add("RAGIONESOCIALE", ListSortDirection.Ascending);

            radGridView1.GroupDescriptors.Add(descriptor);

        }

        private void sendgroupclienti()
        { 
            
        }

        private void btnSendMail_Click(object sender, EventArgs e)
        {
            string address = Properties.Settings.Default.sendMailBCCSimulazione; //;kavanzi@italcom.biz";
            string addressCC = "";  //Properties.Settings.Default.sendMailBCCSimulazione; //"alfredo.deangelo@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string addressBCC = Properties.Settings.Default.sendMailBCC; // "knosmail@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string body = ""; 
            string subject = "";
            string dettaglioBOL = "";
            string dettaglioSCH = "";
            string dettaglioSCHAllegati = "";

            bool bOKDownload = false;

            string codclifor = "";
            string codart = "";

            
            int IdObjectBOL = 0;
            int IdDocBOL = 0;
            int IdObjectSCH = 0;
            int IdDocSCH = 0;
            string localfilenameBOL = "";
            string localfilenameSCH = "";
            string localfileBOL = "";
            string localfileSCH = "";
            string subjectBOL= "";
            string subjectSCH = "";
            int IdObjectSentMail = 0;
            string msg = "";
            
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = radGridView1.SelectedRows.Count+1;
            toolStripProgressBar1.Value = 1;
            toolStripProgressBar1.Step = 1;
            toolStripProgressBar1.Visible = true;

            //Knos
            if (refreshKnosLogin() == false)
                return;

            SaveGridSettings(radGridView1);

            Application.DoEvents();


            log.LogSomething(string.Format("Nr mail da inviare: {0}", radGridView1.SelectedRows.Count));

            string tempPathDownload = Path.Combine(Application.StartupPath, "TEMP");
            if (!Directory.Exists(tempPathDownload))
            {
                Directory.CreateDirectory(tempPathDownload);
            }

            cleanTempFolder(tempPathDownload);

            checkBoxInterrompiInvio.Enabled = true;

            if (radGridView1.SelectedRows.Count > 0)
            {
                msg = string.Format("Procedo con l'invio delle notifiche {0}", radGridView1.SelectedRows.Count);

                if (MessageBox.Show(msg, "Invio Schede e Bollettini", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {

                    for (int i = 0; i < radGridView1.SelectedRows.Count; i++)
                    {
                        toolStripProgressBar1.Value += 1;
                        toolStripProgressBar1.Text = string.Format("Record {0}/{1}", i+1, radGridView1.SelectedRows.Count);
                        log.LogSomething(string.Format("Record {0}/{1}", i + 1, radGridView1.SelectedRows.Count));

                        var attachments = new List<string>();

                        localfileBOL = "";
                        localfileSCH = "";
                        IdObjectBOL = IdDocBOL = IdObjectSCH = IdDocSCH = 0;
                        dettaglioBOL = dettaglioSCH = "";

                        codclifor = radGridView1.SelectedRows[i].Cells["CODCLIFOR"].Value.ToString();
                        codart = radGridView1.SelectedRows[i].Cells["CODART"].Value.ToString();

                        int.TryParse(radGridView1.SelectedRows[i].Cells["IDOBJECT_BOL"].Value.ToString(), out IdObjectBOL);
                        int.TryParse(radGridView1.SelectedRows[i].Cells["IDDOC_BOL"].Value.ToString(), out IdDocBOL);
                        int.TryParse(radGridView1.SelectedRows[i].Cells["IDOBJECT_SCH"].Value.ToString(), out IdObjectSCH);
                        int.TryParse(radGridView1.SelectedRows[i].Cells["IDDOC_SCH"].Value.ToString(), out IdDocSCH);

                        if (chkSimulazione.Checked == false)
                        {
                            // destinatari reali
                            address = radGridView1.SelectedRows[i].Cells["EMAIL_CLIENTE"].Value.ToString();
                            addressCC = radGridView1.SelectedRows[i].Cells["EMAIL_AGENTE"].Value.ToString();

                            if (address.Contains(addressCC) && addressCC.Length > 0)
                                address = address.Replace(addressCC, "");



                        }

                        log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));

                        subjectSCH = string.Format(Properties.Settings.Default.sendMailSchedeTecnicheSubject, radGridView1.SelectedRows[i].Cells["RAGIONESOCIALE"].Value.ToString());

                        if ((IdDocBOL > 0) && (checkBoxBOL.Checked == true))
                        {
                            //subjectBOL = string.Format(Properties.Settings.Default.sendMailBollettiniSubject, radGridView1.SelectedRows[i].Cells["RAGIONESOCIALE"].Value.ToString());
                            localfilenameBOL = "[B]-" + radGridView1.SelectedRows[i].Cells["FILENAME_BOL"].Value.ToString().Substring(16).Replace("#", "");
                            localfileBOL = Path.Combine(tempPathDownload, localfilenameBOL);//"[B]-" + radGridView1.SelectedRows[i].Cells["FILENAME_BOL"].Value.ToString().Substring(16));

                            dettaglioBOL = string.Format("\r\nRMI - ARTICOLO/ITEM: {0} {1} - {2}", radGridView1.SelectedRows[i].Cells["ARTICOLOBSC"].Value.ToString(), radGridView1.SelectedRows[i].Cells["DESCRIZIONEARTICOLO"].Value.ToString(), radGridView1.SelectedRows[i].Cells["CODICEDOCUMENTO_BOL"].Value.ToString());
                            toolStripStatusLabel1.Text = dettaglioBOL;

                            Application.DoEvents();


                            bOKDownload = kw.downloadDoc(IdObjectBOL, IdDocBOL, tempPathDownload, localfilenameBOL);

                            if (bOKDownload == false)
                            {
                                txtLog.Text += string.Format("\r\nNON sono riuscito a scaricare il file della pubblicazione IdObject {0} IdDoc {1} nella cartella {2} con nome {3}", IdObjectBOL, IdDocBOL, tempPathDownload, localfilenameBOL);
                            }

                            attachments.Add(localfileBOL);


                        }


                        //if (IdDocSCH > 0)
                        if ((IdDocSCH > 0) && (checkBoxSCH.Checked == true))
                        {
                            //subjectSCH = string.Format(Properties.Settings.Default.sendMailSchedeTecnicheSubject, radGridView1.SelectedRows[i].Cells["RAGIONESOCIALE"].Value.ToString());
                            localfilenameSCH = "[S]-" + radGridView1.SelectedRows[i].Cells["FILENAME_SCH"].Value.ToString().Substring(16).Replace("#", "");
                            localfileSCH = Path.Combine(tempPathDownload, localfilenameSCH);// "[S]-" + radGridView1.SelectedRows[i].Cells["FILENAME_SCH"].Value.ToString().Substring(16));

                            dettaglioSCH = string.Format("\r\nSCHEDA DI SICUREZZA/SAFETY DATA SHEET - ARTICOLO/ITEM: {0} {1} - {2}", radGridView1.SelectedRows[i].Cells["ARTICOLOBSC"].Value.ToString(), radGridView1.SelectedRows[i].Cells["DESCRIZIONEARTICOLO"].Value.ToString(), radGridView1.SelectedRows[i].Cells["CODICEDOCUMENTO_BOL"].Value.ToString());
                            toolStripStatusLabel1.Text = dettaglioSCH;

                            Application.DoEvents();

                            bOKDownload = kw.downloadDoc(IdObjectSCH, IdDocSCH, tempPathDownload, localfilenameSCH);

                            if (bOKDownload == false)
                            {
                                log.LogSomething(string.Format("NON sono riuscito a scaricare il file della pubblicazione IdObject {0} IdDoc {1} nella cartella {2} con nome {3}", IdObjectSCH, IdDocSCH, tempPathDownload, localfilenameSCH));

                                txtLog.Text += string.Format("\r\nNON sono riuscito a scaricare il file della pubblicazione IdObject {0} IdDoc {1} nella cartella {2} con nome {3}", IdObjectSCH, IdDocSCH, tempPathDownload, localfilenameSCH);
                            }

                            attachments.Add(localfileSCH);


                            // cerca allegati
                            if (chkAllegati.Checked == true)
                            {
                                List<Allegato> listA = new List<Allegato>();

                                foreach (DataGridViewRow r in dataGridView1.Rows)
                                {
                                    if (r.Cells[0].Value != null)
                                    {
                                        Allegato a = new Allegato(r.Cells[0].Value.ToString(), r.Cells[0].Value.ToString(), r.Cells[2].Value.ToString());

                                        listA.Add(a);
                                    }

                                }

                                if (listA.Count > 0)
                                {
                                    dettaglioSCHAllegati = "<br /> ALLEGATI / ATTACHEMENTS:";




                                    foreach (Allegato a in listA)
                                    {
                                        dettaglioSCHAllegati += string.Format("<br /> - {0}", a.FileName);
                                        localfileSCH = a.Path.Replace("file://", "");
                                        attachments.Add(localfileSCH);
                                
                                    }
                                
                                }
                            }




                        }

                        // invio singolo
                        Application.DoEvents();

                        if (bOKDownload)
                        {

                            body = string.Format(Properties.Settings.Default.sendMailSchede, dettaglioBOL, dettaglioSCH, dettaglioSCHAllegati);

                            if (radGridView1.SelectedRows[i].Cells["CODLINGUA"].Value.ToString() != "IT_")
                            {
                                // inglese da definire
                                subject = string.Format("SEND BOLLETTINI-SCHEDE - CUSTOMER: {0}", radGridView1.SelectedRows[i].Cells["RAGIONESOCIALE"].Value.ToString());
                                body = string.Format(Properties.Settings.Default.sendMailBollettini, dettaglioBOL, dettaglioSCH);
                            }

                            if (checkBoxInterrompiInvio.Checked == true)
                            {
                                MessageBox.Show("Invio email interrotto dall'utente!", "Invio notifiche", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                checkBoxInterrompiInvio.Enabled = false;
                                checkBoxInterrompiInvio.Checked = false;

                                return;
                            }

                            if (IdDocSCH > 0 || IdDocBOL > 0)

                                if (Properties.Settings.Default.UseVbs)
                                {
                                    log.LogSomething("Invio tramite UseVBS - SendNotifyVBSLotus");

                                    subject = string.Format("{0}  {1}", subjectBOL, subjectSCH);
                                    dettaglioSCH = string.Format("SCHEDA DI SICUREZZA/SAFETY DATA SHEET - ARTICOLO/ITEM: {0} {1} - {2}", radGridView1.SelectedRows[i].Cells["ARTICOLOBSC"].Value.ToString(), radGridView1.SelectedRows[i].Cells["DESCRIZIONEARTICOLO"].Value.ToString(), radGridView1.SelectedRows[i].Cells["CODICEDOCUMENTO_BOL"].Value.ToString());
                                    body = string.Format(Properties.Settings.Default.sendMailBollettiniVbsLotus, "", dettaglioSCH);
                                    //if (Notifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                    if (Notifica.SendNotifyVBSLotus(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                    {
                                        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                                        if (chkSimulazione.Checked == false)
                                        {
                                            if (chkAllegati.Checked) //IdDocBOL > 0)
                                            {
                                                // aggirono registro
                                                updateInvioSchedeBollettini(codart, codclifor, "DATAULTIMOINVIO_BOL");
                                            }

                                            if (IdDocSCH > 0)
                                            {
                                                // aggirono registro
                                                updateInvioSchedeBollettini(codart, codclifor, "DATAULTIMOINVIO_SCH");
                                            }
                                        }
                                        textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, null, attachments[0], addressCC, addressBCC);

                                        // store della mail inviata
                                        try
                                        {
                                            toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                            IdObjectSentMail = 0;
                                            Application.DoEvents();
                                            IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridView1.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString(), System.DateTime.Now, "msds", address, addressCC, addressBCC, subject, body, "", null, attachments, true);
                                            log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                            textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.Message);
                                        }

                                        radGridView1.ShowRowHeaderColumn = true;

                                        radGridView1.SelectedRows[i].Cells[1].Style.DrawFill = true;
                                        radGridView1.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                        radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                                        radGridView1.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                                        radGridView1.SelectedRows[i].Cells[2].Style.DrawFill = true;
                                        radGridView1.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                        radGridView1.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                                        Application.DoEvents();
                                    }
                                    else
                                    {
                                        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                                        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                                        radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                    }
                                }
                                else
                                {


                                    //if (Properties.Settings.Default.UseLotus)
                                    //{
                                    //    log.LogSomething("Invio tramite UseLotus - SendNotifyMAPILotus");

                                    //    //                                        if (Notifica.SendNotifyMAPI(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                                    //    if (Notifica.SendNotifyMAPILotus(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                                    //    {
                                    //        radGridView1.ShowRowHeaderColumn = true;

                                    //        radGridView1.SelectedRows[i].Cells[1].Style.DrawFill = true;
                                    //        radGridView1.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                    //        radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                                    //        radGridView1.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                                    //        radGridView1.SelectedRows[i].Cells[2].Style.DrawFill = true;
                                    //        radGridView1.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                    //        radGridView1.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                                    //        Application.DoEvents();
                                    //        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                                    //        textBoxLOG.Text += string.Format("\r\n - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC);




                                    //    }
                                    //    else
                                    //    {
                                    //        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                                    //        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                                    //        radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                    //    }
                                    //}
                                    //else
                                    //{
                                        log.LogSomething("Invio tramite UseCdo - SendNotifyCdo");
                                        if (Properties.Settings.Default.UseCdo)
                                        {
                                            subject = string.Format("{0}  {1}", subjectBOL, subjectSCH);

                                            Notifica cNotifica = new ToDoNotificheBSC.Notifica();

                                            if (cNotifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                            //if (Notifica.SendNotifyVBS(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                            {
                                                log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                                                if (chkSimulazione.Checked == false)
                                                {
                                                    if (chkAllegati.Checked) //IdDocBOL > 0)
                                                    {
                                                        // aggirono registro
                                                        updateInvioSchedeBollettini(codart, codclifor, "DATAULTIMOINVIO_BOL");
                                                    }

                                                    if (IdDocSCH > 0)
                                                    {
                                                        // aggiorno registro
                                                        updateInvioSchedeBollettini(codart, codclifor, "DATAULTIMOINVIO_SCH");
                                                    }
                                                }
                                                textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, null, attachments[0], addressCC, addressBCC);

                                                // store della mail inviata

                                                toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                                IdObjectSentMail = 0;
                                                Application.DoEvents();
                                                IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridView1.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString(), System.DateTime.Now, "msds", address, addressCC, addressBCC, subject, body, "", null, attachments, true, false);
                                                log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                                textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);


                                                radGridView1.ShowRowHeaderColumn = true;

                                                radGridView1.SelectedRows[i].Cells[1].Style.DrawFill = true;
                                                radGridView1.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                                radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                                                radGridView1.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                                                radGridView1.SelectedRows[i].Cells[2].Style.DrawFill = true;
                                                radGridView1.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                                radGridView1.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                                                Application.DoEvents();
                                            }
                                            else
                                            {
                                                log.LogSomething(string.Format("ERRORE - Invio mail NON riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                                                textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                                                radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                            }
                                        }
                                    //}

                                }


                        }

                    }

                    toolStripProgressBar1.Visible = false;
                    checkBoxInterrompiInvio.Enabled = false;

                    MessageBox.Show("Invio completato!");
                }
            }
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

        private void radGridView1_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            int IdObjectBOL = 0;
            int IdDocBOL = 0;
            int IdObjectSCH = 0;
            int IdDocSCH = 0;


            if (refreshKnosLogin() == false)
                return;

            //http://vsrv2k8bsn2:8780/KnoS_Catalog/0/0000035964/0001/1426089157015/Zetesol%20MGS.doc
            string url = "{0}/KnoS_Catalog/0/{1}/{2}/{3}";

            if (e.ColumnIndex > 0 && e.ColumnIndex > 0)
            {
                if (radGridView1.Rows[e.RowIndex].Cells["FILENAME_BOL"].ColumnInfo.Index == e.ColumnIndex)
                {
                    url = string.Format(url, txtKnosUrl.Text, radGridView1.Rows[e.RowIndex].Cells["IDOBJECT_BOL"].Value.ToString(), radGridView1.Rows[e.RowIndex].Cells["IDDOC_BOL"].Value.ToString(), radGridView1.Rows[e.RowIndex].Cells["FILENAME_BOL"].Value.ToString().Substring(0, 15) + ".PDF");
                    url = url.Replace("#", "_");

                    webBrowser2.Navigate(url);
                }

                if (radGridView1.Rows[e.RowIndex].Cells["FILENAME_SCH"].ColumnInfo.Index == e.ColumnIndex)
                {
                    url = string.Format(url, txtKnosUrl.Text, radGridView1.Rows[e.RowIndex].Cells["IDOBJECT_SCH"].Value.ToString(), radGridView1.Rows[e.RowIndex].Cells["IDDOC_SCH"].Value.ToString(), radGridView1.Rows[e.RowIndex].Cells["FILENAME_SCH"].Value.ToString().Substring(0, 15) + ".PDF");
                    url = url.Replace("#", "_");

                    webBrowser2.Navigate(url);
                }

            }
            else
            { 
                // selezione celle
                Debug.Print(string.Format("ci: {0}  - ri: {1}", e.ColumnIndex, e.RowIndex));

                if ((e.RowIndex == -1) && (e.ColumnIndex == -1))
                {
                    radGridView1.SelectAll();
                }

                if ((e.RowIndex > -1) && (e.ColumnIndex == -1))
                {

                    if (e.Row.Group != null)
                    {
                        for (int x = 0; x < e.Row.Parent.ChildRows.Count; x++)
                        {

                            e.Row.Parent.ChildRows[x].IsSelected = true;
                            Application.DoEvents();

                        }


                    }

                }


            
            
            }
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



            kw.GetMyCertificates("");



        }






        void getSchede(string path)
        {
            string filename = "";
            string codart = "";
            string revisione = "";
            string lingua = "";
            DateTime datamodifica = System.DateTime.Now;
            string codartMetodo = "";
            string descrizioneMetodo = "";
            int IDOBJECT = 0;
            string subpath = "";

            DataTable dt = new DataTable();
            dt.Columns.Add("filename");
            dt.Columns.Add("codart");
            dt.Columns.Add("revisione");
            dt.Columns.Add("lingua");
            dt.Columns.Add("datamodifica");
            dt.Columns.Add("codartMetodo");
            dt.Columns.Add("descrizioneMetodo");
            dt.Columns.Add("IDOBJECT");
            dt.Columns.Add("subpath");

            int i = 0;


            if (Properties.Settings.Default.SqlSchede != "")
            {
                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(string.Format(Properties.Settings.Default.SqlSchede, "")))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        SqlDataAdapter da = new SqlDataAdapter(cmd);

                        da.Fill(dt);


                    }
                }
            }
            else
            {


                if (Directory.Exists(path))
                {
                    string[] files = Directory.GetFiles(path);

                    foreach (string f in files)
                    {
                        i += 1;

                        FileInfo fi = new FileInfo(f);
                        filename = Path.GetFileNameWithoutExtension(fi.FullName);
                        string[] fname = filename.Split('_');
                        filename = fi.Name;

                        lingua = fname[0];
                        codart = fname[1];
                        revisione = fname[2];

                        datamodifica = fi.LastWriteTime;
                        subpath = Path.GetFullPath(fi.FullName);
                        // articolo Metodo
                        List<string> artmetodo = getArticoloMetodo(codart);

                        codartMetodo = "";
                        descrizioneMetodo = "";
                        IDOBJECT = 0;

                        if (artmetodo.Count == 1)
                        {
                            string[] a = artmetodo[0].Split('_');

                            codartMetodo = a[0];
                            descrizioneMetodo = a[1];
                            int.TryParse(a[2], out IDOBJECT);

                        }

                        statusStrip1.Text = string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length);
                        Application.DoEvents();

                        object[] orow = { filename, codart, revisione, lingua, datamodifica, codartMetodo, descrizioneMetodo, IDOBJECT, subpath };
                        dt.Rows.Add(orow);



                    }

                }

                radGridViewEpy.DataSource = dt;

                radGridViewEpy.EnableSorting = true;
                radGridViewEpy.SortDescriptors.Add("codart", ListSortDirection.Ascending);
                radGridViewEpy.SortDescriptors.Add("lingua", ListSortDirection.Ascending);
                radGridViewEpy.SortDescriptors.Add("revisione", ListSortDirection.Ascending);

            }


        }


        private List<string> getArticoloMetodo(string codart)
        { 
            List<string> lOut = new List<string>();

            using(SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
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
        
        private void btnProcessEpy_Click(object sender, EventArgs e)
        {
            string filename;
            string descrizioneMetodo;
            string codlingua;
            string revisione;
            string subpath = "";

            int IdObjectCliAll = 38750;
            int IdObjectArticolo = 0;
            int IdObjectScheda = 0;
            IKnosObjectSelector os = KnosInstance.Client.CreateKnosObjectSelector();

            List<int> schedeaggiornate = new List<int>();


            //Knos
            if (refreshKnosLogin() == false)
                return;


            string searchExpression = " IdClass = 5012 and exists(select 1 from object_linkage x where x.idparent = _idobject_ and idchild = {0}) and exists(select 1 from object_linkage x where x.idparent = _idobject_ and idchild = {1})";

            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.Minimum = 1;

            for (int i = 0; i < radGridViewEpy.SelectedRows.Count;  i++)
            {
                toolStripProgressBar1.Maximum = radGridViewEpy.RowCount;
                toolStripProgressBar1.Step = 1;
                toolStripProgressBar1.PerformStep();

                IdObjectArticolo = 0;
                IdObjectScheda = 0;
                filename = descrizioneMetodo = codlingua = "";

                os.Reset(true);


                Application.DoEvents();
                
                int.TryParse(radGridViewEpy.SelectedRows[i].Cells["IDOBJECT"].Value.ToString(), out IdObjectArticolo);
                filename = radGridViewEpy.SelectedRows[i].Cells["filename"].Value.ToString().Replace("#", "");
                string codartMetodo = radGridViewEpy.SelectedRows[i].Cells["codartMetodo"].Value.ToString();
                descrizioneMetodo = radGridViewEpy.SelectedRows[i].Cells["descrizioneMetodo"].Value.ToString();
                codlingua = radGridViewEpy.SelectedRows[i].Cells["lingua"].Value.ToString();
                revisione = radGridViewEpy.SelectedRows[i].Cells["revisione"].Value.ToString();
                subpath = radGridViewEpy.SelectedRows[i].Cells["subpath"].Value.ToString();

                string title = string.Format("{0} [{1}]", descrizioneMetodo, codartMetodo);

                Application.DoEvents();
                toolStripStatusLabel1.Text = string.Format("Allego file: {0}, {1}, {2} ({3}/{4})", filename, descrizioneMetodo, codlingua, i, radGridViewEpy.RowCount);

                if (IdObjectArticolo > 0)
                {
                    os.SearchExpression = string.Format(searchExpression, IdObjectCliAll, IdObjectArticolo);
                    os.GetPage(1);

                    if (os.RecordCount == 1)
                    {
                        // scheda esistente
                        IdObjectScheda = os.GetItem(0).IdObject;

                        if (schedeaggiornate != null)
                        {
                            if (schedeaggiornate.Contains(IdObjectScheda) == false)
                            {
                                kw.DeleteFiles(IdObjectScheda, -1);
                                schedeaggiornate.Add(IdObjectScheda);

                                if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateSchede, "") == false)
                                {
                                    if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateSchedeRedazione, "") == false)
                                    {
                                        log.LogSomething(string.Format("Errore nella transizione di stato della scheda! {0}", IdObjectScheda));
                                    }
                                }

                            }

                        }
                    }
                    else
                    { 
                        // la creo
                        IKnosMultivalueEditor meC = KnosInstance.Client.CreateKnosMultivalueEditor();
                        meC.AddValue(IdObjectCliAll);
                        IKnosMultivalueEditor meA = KnosInstance.Client.CreateKnosMultivalueEditor();
                        meA.AddValue(IdObjectArticolo);
                        
                        IKnosObjectMaker om = KnosInstance.Client.CreateKnosObjectMaker();
                        om.IdClass = 5012;
                        om.SetAttrValue("object_19", meC, EnumKnosDataType.ObjectListType);
                        om.SetAttrValue("object_5036", meA, EnumKnosDataType.ObjectListType);
                        om.SetAttrValue("title", title);
                        om.CreateObject(out IdObjectScheda);
                    
                    
                    }

                    if (IdObjectScheda == 0)
                    {
                        MessageBox.Show("Nessuna scheda trovata o creata!");
                    }
                    else
                    {

                        //upload file
                        if (kw.UploadFileCertificato(IdObjectScheda, 0, subpath, string.Format("{0}-{1}", codlingua, descrizioneMetodo), filename, 0, "", revisione))
                        {
                            radGridViewEpy.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                            Application.DoEvents();
                        }
                    }




                }



            
            }

            // aggiornamento registro schede
            aggiornaRegistro();


            DateTime dtUltimoAggiornamento;
            DateTime.TryParse(txtUltimoAggiornamento.Text, out dtUltimoAggiornamento);
            Properties.Settings.Default.DataAggiornamento = dtUltimoAggiornamento;
            Properties.Settings.Default.Save();

            MessageBox.Show("Aggiornamento completato!", "Importazione schede da Epy", MessageBoxButtons.OK);
            
        }


        void aggiornaRegistro()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(Properties.Settings.Default.SqlUpdateBSC.ToString()))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            log.LogSomething(string.Format("Errore nell'aggiornamento del registro! {0}", ex.Message));

                        }
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



        void aggiornaRegistroPA()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;

                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(Properties.Settings.Default.SqlUpdateBSC_PA.ToString()))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            log.LogSomething(string.Format("Errore nell'aggiornamento del registro Primo Acquisto! {0}", ex.Message));

                        }
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

        //private void dataGridViewMyCertificates_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        //{
        //    if ((e.RowIndex >= 0) && (e.ColumnIndex >= 0))
        //    {

        //        if (e.ColumnIndex == dataGridViewMyCertificates.Columns["dataGridViewButtonColumn1"].Index)
        //        {
        //            ((DataGridViewButtonColumn)dataGridViewMyCertificates.Columns[dataGridViewMyCertificates.Columns["dataGridViewButtonColumn1"].Index]).DefaultCellStyle.ForeColor = Color.Red;
        //        }
        //    }
        //}

        private void btnOldArchivio_Click(object sender, EventArgs e)
        {


            string filename = "";
            string codart = "";
            string revisione = "";
            string lingua = "";
            DateTime datamodifica = System.DateTime.Now;
            string codartMetodo = "";
            string descrizioneMetodo = "";
            int IDOBJECT = 0;
            string subpath = "";
            string nomeEPY = "";

            DataTable dt = new DataTable();
            dt.Columns.Add("filename");
            dt.Columns.Add("codart");
            dt.Columns.Add("revisione");
            dt.Columns.Add("lingua");
            dt.Columns.Add("datamodifica");
            dt.Columns.Add("codartMetodo");
            dt.Columns.Add("descrizioneMetodo");
            dt.Columns.Add("IDOBJECT");
            dt.Columns.Add("subpath");
            dt.Columns.Add("nomeEPY");

            Cursor.Current = Cursors.WaitCursor;

            System.DateTime dtUltimoAggiornamento;

            if (DateTime.TryParse(txtUltimoAggiornamento.Text, out dtUltimoAggiornamento))
            {
                Properties.Settings.Default.DataAggiornamento = dtUltimoAggiornamento;
                Properties.Settings.Default.Save();

            }
            else
            {
                DateTime.TryParse("01/01/2015", out dtUltimoAggiornamento);
            }

            //MessageBox.Show(string.Format("Leggo gli articoli di cui cercare schede aggiornate dopo {0}", dtUltimoAggiornamento.ToString()));
            Application.DoEvents();


            if (chkFilesDaDB.Checked == true)
            {
                toolStripStatusLabel1.Text = "Caricamento da tabella ....";

                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(string.Format(Properties.Settings.Default.SqlSchede, dtUltimoAggiornamento.ToShortDateString())))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        SqlDataAdapter da = new SqlDataAdapter(cmd);

                        da.Fill(dt);


                    }
                }

                toolStripStatusLabel1.Text = "Caricamento articoli schede COMPLETATO (tabella)";
            }
            else
            {

                log.LogSomething("Caricamento articoli schede");
                toolStripStatusLabel1.Text = "Caricamento articoli schede ....";

                DataTable xx = getArticoliSchede();

                log.LogSomething("Caricamento articoli schede COMPLETATO");
                toolStripStatusLabel1.Text = "Caricamento articoli schede COMPLETATO";
                string path = lblPathEpy.Text;







                int i = 0;

                Application.DoEvents();

                if (Directory.Exists(path))
                {
                    log.LogSomething("Inizio ricerca file EPY");

                    toolStripStatusLabel1.Text = string.Format("ciclo sugli articoli che ho trovato {0}", xx.Rows.Count.ToString());
                    Application.DoEvents();

                    // ciclo sugli articoli che ho trovato
                    for (int xi = 0; xi < xx.Rows.Count; xi++)
                    {
                        string xi_filename = xx.Rows[xi]["NOMEEPY"].ToString();

                        if (xi_filename != "")
                        {
                            log.LogSomething("Inizio ricerca file EPY  * " + xi_filename + " * ");

                            string[] files = Directory.GetFiles(path, "*" + xi_filename + "*", SearchOption.AllDirectories);

                            foreach (string f in files)
                            {
                                i += 1;



                                FileInfo fi = new FileInfo(f);
                                string.Format("file: {0}", fi.FullName);

                                if (fi.LastWriteTime > dtUltimoAggiornamento)
                                {
                                    filename = Path.GetFileNameWithoutExtension(fi.FullName);
                                    string[] fname = filename.Split('_');
                                    filename = fi.Name;

                                    lingua = fname[0];
                                    codart = xx.Rows[xi]["CODICE"].ToString();
                                    revisione = fname[2];
                                    nomeEPY = "";

                                    if (fname.Count() == 4)
                                        nomeEPY = fname[3];

                                    datamodifica = fi.LastWriteTime;
                                    subpath = Path.GetFullPath(fi.FullName);

                                    // articolo Metodo
                                    //List<string> artmetodo = getArticoloMetodo(codart);

                                    codartMetodo = codart;
                                    descrizioneMetodo = xx.Rows[xi]["descrizioneMetodo"].ToString();
                                    int.TryParse(xx.Rows[xi]["IDOBJECT"].ToString(), out IDOBJECT);

                                    toolStripStatusLabel1.Text = string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length);

                                    log.LogSomething(string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length));

                                    Application.DoEvents();

                                    object[] orow = { filename, codart, revisione, lingua, datamodifica, codartMetodo, descrizioneMetodo, IDOBJECT, subpath, nomeEPY };
                                    dt.Rows.Add(orow);
                                }

                            }

                        }

                    }
                }
            }

            radGridViewEpy.DataSource = null;
            radGridViewEpy.Rows.Clear();
            radGridViewEpy.DataSource = dt;

            radGridViewEpy.EnableFiltering = true;
            radGridViewEpy.EnableSorting = true;
            radGridViewEpy.SortDescriptors.Add("codart", ListSortDirection.Ascending);
            radGridViewEpy.SortDescriptors.Add("lingua", ListSortDirection.Ascending);
            radGridViewEpy.SortDescriptors.Add("revisione", ListSortDirection.Ascending);


            this.radGridViewEpy.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.DisplayedDataCells);

            lblNrFiles.Text = "Nr Files: " + radGridViewEpy.Rows.Count.ToString();

            MessageBox.Show("Caricamento effettuato!");
            tabControl1.SelectedIndex = 3;

            if (this.Visible == false)
                this.Visible = true;
            

        }



        DataTable getArticoliSchede()
        {
            DataTable x = new DataTable();
            string strSQL = string.Format("SELECT * FROM VISTA_RELAZIONIARTICOLIBSC"); ;


            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
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

        bool updateInvioSchedeBollettini(string art, string cli, string fieldname)
        {

            bool bOK = false;

            try
            {
                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(string.Format("UPDATE dbo.ITA_TABREGISTRONOTIFICHEBSC SET {2} = getdate() WHERE CODART = '{0}' AND CODCLIFOR = '{1}' ", art, cli, fieldname)))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        cmd.ExecuteNonQuery();
                        bOK = true;
                    }

                }
            }
            catch (Exception ex)
            {
                log.LogSomething(string.Format("ERRORE: {0} - SQL: {1}", ex.Message, string.Format("UPDATE dbo.ITA_TABREGISTRONOTIFICHEBSC SET {2} = getdate() WHERE CODART = '{0}' AND CODCLIFOR = '{1}' ", art, cli, fieldname)));
                checkBoxInterrompiInvio.Checked = true;
            }

            return bOK;

        }

        private void labeH_Click(object sender, EventArgs e)
        {
            
        }

        private void btnLeggiArchivioH_Click(object sender, EventArgs e)
        {
            DataTable xx = getArticoliSchedeH();

            string path = lblPathH.Text;

            string filename = "";
            string codart = "";
            string revisione = "";
            string lingua = "";
            DateTime datamodifica = System.DateTime.Now;
            string codartMetodo = "";
            string descrizioneMetodo = "";
            int IDOBJECT = 0;
            string subpath = "";
            string nomeEPY = "";

            DataTable dt = new DataTable();
            dt.Columns.Add("filename");
            dt.Columns.Add("codart");
            dt.Columns.Add("revisione");
            dt.Columns.Add("lingua");
            dt.Columns.Add("datamodifica");
            dt.Columns.Add("codartMetodo");
            dt.Columns.Add("descrizioneMetodo");
            dt.Columns.Add("IDOBJECT");
            dt.Columns.Add("subpath");
            dt.Columns.Add("nomeEPY");

            int i = 0;

            Application.DoEvents();

            if (Directory.Exists(path))
            {

                toolStripStatusLabel1.Text = string.Format("ciclo sugli articoli che ho trovato {0}", xx.Rows.Count.ToString());
                Application.DoEvents();

                // ciclo sugli articoli che ho trovato
                for (int xi = 0; xi < xx.Rows.Count; xi++)
                {
                    string xi_filename = xx.Rows[xi]["FILEPREFIX"].ToString();



                    if (xi_filename != "")
                    {

                        xi_filename = xi_filename.TrimEnd(' ');


                        string[] files = Directory.GetFiles(path, xi_filename + "*", SearchOption.AllDirectories);

                        foreach (string f in files)
                        {
                            i += 1;



                            FileInfo fi = new FileInfo(f);



                            switch (fi.Directory.Name.ToUpper())
                            {
                                case "FRANCESE":
                                    lingua = "FR";
                                    break;

                                case "BULGARO":
                                    lingua = "BL";
                                    break;

                                case "GRECO":
                                    lingua = "EL";
                                    break;

                                case "INGLESE":
                                    lingua = "EN";
                                    break;

                                case "ITALIANO":
                                    lingua = "IT";
                                    break;

                                case "OLANDESE":
                                    lingua = "NL";
                                    break;

                                case "POLACCO":
                                    lingua = "PL";
                                    break;

                                case "PORTOGHESE":
                                    lingua = "PT";
                                    break;

                                case "SLOVENO":
                                    lingua = "SL";
                                    break;

                                case "SPAGNOLO":
                                    lingua = "SP";
                                    break;

                                case "SVEDESE":
                                    lingua = "SV";
                                    break;

                                case "TEDESCO":
                                    lingua = "DE";
                                    break;



                                default:
                                    lingua = "";
                                    break;
                            }


                            //if (fi.LastWriteTime > dtUltimoAggiornamento)
                            //{

                            codart = xx.Rows[xi]["CODICE"].ToString();
                            revisione = "X";
                            filename = string.Format("{0}_{1}_{2}_{3}", lingua, codart, revisione, xi_filename + fi.Extension);


                            datamodifica = fi.LastWriteTime;
                            subpath = Path.GetFullPath(fi.FullName);

                            // articolo Metodo
                            //List<string> artmetodo = getArticoloMetodo(codart);

                            codartMetodo = codart;
                            descrizioneMetodo = xx.Rows[xi]["descrizioneMetodo"].ToString();
                            int.TryParse(xx.Rows[xi]["IDOBJECT"].ToString(), out IDOBJECT);

                            toolStripStatusLabel1.Text = string.Format("file: {0}  nr {1} - {2}", filename , i, files.Length);
                            Application.DoEvents();

                            object[] orow = { filename, codart, revisione, lingua, datamodifica, codartMetodo, descrizioneMetodo, IDOBJECT, subpath, nomeEPY };
                            dt.Rows.Add(orow);
                            //}

                        }

                    }

                }

            }

            //ricerca in cartella precedente
            path = Path.Combine(Directory.GetParent(path).FullName, "Prodotti ECOCERT");
            if (Directory.Exists(path))
            {

                toolStripStatusLabel1.Text = string.Format("ciclo sugli articoli che ho trovato {0}", xx.Rows.Count.ToString());
                Application.DoEvents();

                // ciclo sugli articoli che ho trovato
                for (int xi = 0; xi < xx.Rows.Count; xi++)
                {
                    string xi_filename = xx.Rows[xi]["FILEPREFIX"].ToString();



                    if (xi_filename != "")
                    {

                        xi_filename = xi_filename.TrimEnd(' ');


                        string[] files = Directory.GetFiles(path, xi_filename + "*.PDF", SearchOption.AllDirectories);

                        foreach (string f in files)
                        {
                            i += 1;



                            FileInfo fi = new FileInfo(f);


                            toolStripStatusLabel1.Text = string.Format("file: {0}  nr {1} - {2}", fi.Name.ToUpper(), i, files.Length);

                            if ((fi.Name.ToUpper().Contains("BOL-")))
                            { }
                            else
                            {
                                lingua = "";

                                if (fi.Name.ToUpper().Contains("FRANCESE"))
                                        lingua = "FR";

                                if (fi.Name.ToUpper().Contains("BULGARO"))
                                        lingua = "BL";

                                if (fi.Name.ToUpper().Contains("GRECO"))
                                        lingua = "EL";

                                if (fi.Name.ToUpper().Contains("INGL"))
                                        lingua = "EN";

                                if (fi.Name.ToUpper().Contains("ITALIANO"))
                                        lingua = "IT";

                                if (fi.Name.ToUpper().Contains("OLANDESE"))
                                        lingua = "NL";

                                if (fi.Name.ToUpper().Contains("POLACCO"))
                                        lingua = "PL";

                                if (fi.Name.ToUpper().Contains("PORTOGHESE"))
                                        lingua = "PT";

                                if (fi.Name.ToUpper().Contains("SLOVENO"))
                                        lingua = "SL";

                                if (fi.Name.ToUpper().Contains("SPAGNOLO"))
                                        lingua = "SP";

                                if (fi.Name.ToUpper().Contains("SVEDESE"))
                                        lingua = "SV";

                                if (fi.Name.ToUpper().Contains("TEDESCO"))
                                        lingua = "DE";



                                //if (fi.LastWriteTime > dtUltimoAggiornamento)
                                //{

                                codart = xx.Rows[xi]["CODICE"].ToString();
                                revisione = "X";
                                filename = string.Format("{0}_{1}_{2}_{3}", lingua, codart, revisione, xi_filename + fi.Extension);


                                datamodifica = fi.LastWriteTime;
                                subpath = Path.GetFullPath(fi.FullName);

                                // articolo Metodo
                                //List<string> artmetodo = getArticoloMetodo(codart);

                                codartMetodo = codart;
                                descrizioneMetodo = xx.Rows[xi]["descrizioneMetodo"].ToString();
                                int.TryParse(xx.Rows[xi]["IDOBJECT"].ToString(), out IDOBJECT);

                                toolStripStatusLabel1.Text = string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length);
                                Application.DoEvents();

                                object[] orow = { filename, codart, revisione, lingua, datamodifica, codartMetodo, descrizioneMetodo, IDOBJECT, subpath, nomeEPY };
                                dt.Rows.Add(orow);
                                //}
                            }

                        }

                    }

                }





                radGridViewH.DataSource = null;
                radGridViewH.Rows.Clear();
                radGridViewH.DataSource = dt;

                radGridViewH.EnableSorting = true;
                radGridViewH.SortDescriptors.Add("codart", ListSortDirection.Ascending);
                radGridViewH.SortDescriptors.Add("lingua", ListSortDirection.Ascending);
                radGridViewH.SortDescriptors.Add("revisione", ListSortDirection.Ascending);


                this.radGridViewH.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.DisplayedDataCells);

                lblNrFilesH.Text = "Nr Files: " + radGridViewH.Rows.Count.ToString();

                MessageBox.Show("Caricamento effettuato!");
            }
        }




        DataTable getArticoliSchedeH()
        {
            DataTable x = new DataTable();

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(string.Format("SELECT * FROM VISTA_RELAZIONIARTICOLIBSC_H WHERE CODARTMETODO LIKE '%{0}%'", textBox2.Text)))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);

                }

            }
            return x;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string filename;
            string descrizioneMetodo;
            string codlingua;
            string revisione;
            string subpath = "";

            int IdObjectCliAll = 38750;
            int IdObjectArticolo = 0;
            int IdObjectScheda = 0;
            IKnosObjectSelector os = KnosInstance.Client.CreateKnosObjectSelector();

            List<int> schedeaggiornate = new List<int>();

            //Knos
            if (refreshKnosLogin() == false)
                return;

            string searchExpression = " IdClass = 5012 and exists(select 1 from object_linkage x where x.idparent = _idobject_ and idchild = {0}) and exists(select 1 from object_linkage x where x.idparent = _idobject_ and idchild = {1})";

            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.Minimum = 1;

            for (int i = 0; i < radGridViewH.SelectedRows.Count; i++)
            {
                toolStripProgressBar1.Maximum = radGridViewH.RowCount;
                toolStripProgressBar1.Step = 1;
                toolStripProgressBar1.PerformStep();

                IdObjectArticolo = 0;
                IdObjectScheda = 0;
                filename = descrizioneMetodo = codlingua = "";

                os.Reset(true);


                Application.DoEvents();

                int.TryParse(radGridViewH.Rows[i].Cells["IDOBJECT"].Value.ToString(), out IdObjectArticolo);
                filename = radGridViewH.Rows[i].Cells["filename"].Value.ToString().Replace("#", "");
                string codartMetodo = radGridViewH.Rows[i].Cells["codartMetodo"].Value.ToString();
                descrizioneMetodo = radGridViewH.Rows[i].Cells["descrizioneMetodo"].Value.ToString();
                codlingua = radGridViewH.Rows[i].Cells["lingua"].Value.ToString();
                revisione = radGridViewH.Rows[i].Cells["revisione"].Value.ToString();
                subpath = radGridViewH.Rows[i].Cells["subpath"].Value.ToString();

                string title = string.Format("{0} [{1}]", descrizioneMetodo, codartMetodo);

                Application.DoEvents();
                toolStripStatusLabel1.Text = string.Format("Allego file: {0}, {1}, {2} ({3}/{4})", filename, descrizioneMetodo, codlingua, i, radGridViewH.RowCount);

                if (IdObjectArticolo > 0)
                {
                    os.SearchExpression = string.Format(searchExpression, IdObjectCliAll, IdObjectArticolo);
                    os.GetPage(1);

                    if (os.RecordCount == 1)
                    {
                        // scheda esistente
                        IdObjectScheda = os.GetItem(0).IdObject;

                        if (schedeaggiornate != null)
                        {
                            if (schedeaggiornate.Contains(IdObjectScheda) == false)
                            {
                                kw.DeleteFiles(IdObjectScheda, -1);
                                schedeaggiornate.Add(IdObjectScheda);

                                if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateSchede, "") == false)
                                {
                                    if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateSchedeRedazione, "") == false)
                                    {
                                        log.LogSomething(string.Format("Errore nella transizione di stato della scheda! {0}", IdObjectScheda));
                                    }
                                }

                            }

                        }
                    }
                    else
                    {
                        // la creo
                        IKnosMultivalueEditor meC = KnosInstance.Client.CreateKnosMultivalueEditor();
                        meC.AddValue(IdObjectCliAll);
                        IKnosMultivalueEditor meA = KnosInstance.Client.CreateKnosMultivalueEditor();
                        meA.AddValue(IdObjectArticolo);

                        IKnosObjectMaker om = KnosInstance.Client.CreateKnosObjectMaker();
                        om.IdClass = 5012;
                        om.SetAttrValue("object_19", meC, EnumKnosDataType.ObjectListType);
                        om.SetAttrValue("object_5036", meA, EnumKnosDataType.ObjectListType);
                        om.SetAttrValue("title", title);
                        om.CreateObject(out IdObjectScheda);


                    }

                    if (IdObjectScheda == 0)
                    {
                        MessageBox.Show("Nessuna scheda trovata o creata!");
                    }
                    else
                    {

                        //upload file
                        if (kw.UploadFileCertificato(IdObjectScheda, 0, subpath, string.Format("{0}-{1}", codlingua, descrizioneMetodo), filename, 0, "", revisione))
                        {
                            radGridViewH.Rows[i].Cells[1].Style.BackColor = Color.Lime;
                            Application.DoEvents();
                        }
                    }


                }




            }

            // aggiornamento registro schede
            aggiornaRegistro();


            DateTime dtUltimoAggiornamento;
            DateTime.TryParse(txtUltimoAggiornamento.Text, out dtUltimoAggiornamento);
            Properties.Settings.Default.DataAggiornamento = dtUltimoAggiornamento;
            DateTime dtUltimoInvioSchede;
            DateTime.TryParse(txtUltimoInvio.Text, out dtUltimoInvioSchede);
            Properties.Settings.Default.DataInvioSchede = dtUltimoInvioSchede; 
            Properties.Settings.Default.Save();

            MessageBox.Show("Aggiornamento completato!", "Importazione schede da Epy", MessageBoxButtons.OK);
            
        }

        private void radGridView1_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            if (radGridView1.Rows.Count > 0)
            {
                if (e.RowIndex >= 0)
                    datiriga(e.RowIndex);
            }
        }

        private void datiriga(int r)
        {
            string tooltip = "Cliente: {0} | " + 
                "Articolo: {1} \r\n" + 
                "Indirizzo email: {2} | " + 
                "Indirizzo email agenti: {3} \r\n" + 
                "Scheda Sicurezza: {4} | " + 
                "Ultimo invio: {5} \r\n" +
                "RMI: {6} | " +
                "Ultimo invio: {7} \r\n";

            string cliente = string.Format("{0} - {1}", radGridView1.Rows[r].Cells["CODCLIFOR"].Value.ToString(), radGridView1.Rows[r].Cells["RAGIONESOCIALE"].Value.ToString());
            string articolo = radGridView1.Rows[r].Cells["CODART"].Value.ToString();

            string address = radGridView1.Rows[r].Cells["EMAIL_CLIENTE"].Value.ToString();
            string addressCC = radGridView1.Rows[r].Cells["EMAIL_AGENTE"].Value.ToString();

            string fileSCH = "";
            if (radGridView1.Rows[r].Cells["FILENAME_SCH"].Value.ToString().Length > 0)
            {
                fileSCH = radGridView1.Rows[r].Cells["FILENAME_SCH"].Value.ToString().Substring(16) + ".PDF";
            }
            
            string ultimoinvioSCH = "";
            if (radGridView1.Rows[r].Cells["DATAULTIMOINVIO_SCH"].Value != null)
                ultimoinvioSCH = radGridView1.Rows[r].Cells["DATAULTIMOINVIO_SCH"].Value.ToString();


            string fileRMI = "";
            if (radGridView1.Rows[r].Cells["FILENAME_BOL"].Value.ToString().Length > 0)
            {
                fileRMI = radGridView1.Rows[r].Cells["FILENAME_BOL"].Value.ToString().Substring(16) + ".PDF";
            }

            string ultimoinvioRMI = "";
            if (radGridView1.Rows[r].Cells["DATAULTIMOINVIO_BOL"].Value != null)
                ultimoinvioRMI = radGridView1.Rows[r].Cells["DATAULTIMOINVIO_BOL"].Value.ToString();



            textBoxToolTip.Text = string.Format(tooltip, cliente, articolo, address, addressCC, fileSCH, ultimoinvioSCH, fileRMI, ultimoinvioRMI);
        }

        private void radGridView1_CurrentRowChanging(object sender, Telerik.WinControls.UI.CurrentRowChangingEventArgs e)
        {
            if (radGridView1.Rows.Count > 0)
            {
                if (e.NewRow.Index >= 0)
                    datiriga(e.NewRow.Index);
            }
        }


        DataTable getArticoliSchedeCOA()
        {
            DataTable x = new DataTable();
            
            string strSQL = "SELECT * FROM VISTA_NOTIFICHEBSCCOA {0}";

            string strWHERE = " WHERE 1=1";

            strWHERE += String.Format(" AND DATADOC BETWEEN '{0}' AND '{1}'", dTP_BOLLEDA.Value.Date.ToString("yyyyMMdd"), dTP_BOLLEA.Value.Date.ToString("yyyyMMdd"));



            if (chkChiusi.Checked == false)
            {
                strWHERE += " AND QTAGESTRES > 0";
            }
            else
            {
                strWHERE += " AND QTAGESTRES >= 0";
            }

            if (chkInviati.Checked == false)
            {
                strWHERE += " AND ISNULL(DATAINVIOCOA, 1) = 1";
            }

            if (txtLotto.Text != "")
            {
                strWHERE += string.Format(" AND LOTTO LIKE '%{0}%'", txtLotto.Text);
            }
            
            if (txtClienteCOA.Text != "")
            {
                strWHERE += string.Format(" AND RAGIONESOCIALE LIKE '%{0}%'", txtClienteCOA.Text);
            }

            strSQL = string.Format(strSQL, strWHERE);

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
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

        private void btnCercaCOA_Click(object sender, EventArgs e)
        {
            DataTable xx = getArticoliSchedeCOA();

            string path = Properties.Settings.Default.PathCOA;// lblPathH.Text;

            string filename = "";
            string codart = "";
            string cliente = "";
            string documento = "";
            string email = "";
            string lotto = "";
            DateTime datamodifica = System.DateTime.Now;
            string codartMetodo = "";
            string descrizioneMetodo = "";
            int IDOBJECT_DOC = 0;
            int IDOBJECT_CLI = 0;
            DateTime datainvioCOA;
            int IDTESTA = 0;
            int IDRIGA = 0;
            int IDDOC_DDT = 0;
            string FILENAME_DDT = "";

            DataTable dt = new DataTable();
            dt.Columns.Add("filename");
            dt.Columns.Add("cliente");
            dt.Columns.Add("codart");
            dt.Columns.Add("documento");
            dt.Columns.Add("EMAIL_CLIENTE");
            dt.Columns.Add("LOTTO");
            dt.Columns.Add("IDOBJECT_DOC");
            dt.Columns.Add("filenamecoacalc1");
            dt.Columns.Add("filenamecoacalc2");
            dt.Columns.Add("datainvioCOA");
            dt.Columns.Add("IDTESTA");
            dt.Columns.Add("IDRIGA");
            dt.Columns.Add("IDOBJECT_CLI");
            dt.Columns.Add("IDDOC_DDT");
            dt.Columns.Add("FILENAME_DDT");


            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = xx.Rows.Count + 1;
            toolStripProgressBar1.Value = 1;
            toolStripProgressBar1.Step = 1;
            toolStripProgressBar1.Visible = true;


            int i = 0;

            Application.DoEvents();

            if (Directory.Exists(path))
            {

                toolStripStatusLabel1.Text = string.Format("ciclo sugli articoli che ho trovato {0}", xx.Rows.Count.ToString());
                Application.DoEvents();

                // ciclo sugli articoli che ho trovato
                for (int xi = 0; xi < xx.Rows.Count; xi++)
                {
                    // cerco per cliente
                    string xi_filename = xx.Rows[xi]["filenamecoacalc2"].ToString().Replace("/", "").Replace("#", "_");

                    log.LogSomething(string.Format("Ricerca COA {0}/{1} path: {2} cerco file {3}", xi + 1, xx.Rows.Count, path, xi_filename + "*.PDF"));

                    if (xi_filename != "")
                    {

                        Cursor.Current = Cursors.WaitCursor;
                        Application.DoEvents();

                        string[] files = Directory.GetFiles(path, xi_filename + "*.PDF", SearchOption.AllDirectories);

                        if (files.Count() == 0)
                        {
                            xi_filename = xx.Rows[xi]["filenamecoacalc1"].ToString().Replace("/", "");
                        }

                        log.LogSomething(string.Format("Ricerca COA {0}/{1} path: {2} cerco file {3}", xi + 1, xx.Rows.Count, path, xi_filename + "*.PDF"));

                        files = Directory.GetFiles(path, xi_filename + "*.PDF", SearchOption.AllDirectories);

                        if (files.Count() == 0)
                        {
                            xi_filename = xx.Rows[xi]["filenamecoacalc1"].ToString().Replace("/", "").Replace("#", "_");
                        }

                        log.LogSomething(string.Format("Ricerca COA {0}/{1} path: {2} cerco file {3}", xi + 1, xx.Rows.Count, path, xi_filename + "*.PDF"));

                        files = Directory.GetFiles(path, xi_filename + "*.PDF", SearchOption.AllDirectories);

                        toolStripStatusLabel1.Text = string.Format("Ricerca COA {0}/{1} path: {2} cerco file {3}", xi + 1, xx.Rows.Count,  path, xi_filename + "*.PDF");

                               
                        Application.DoEvents();
                        foreach (string f in files)
                        {
                            FileInfo fi = new FileInfo(f);

                            if (!fi.Name.Contains("ENG."))
                            {
                                log.LogSomething(string.Format("Ricerca COA file {0} creato il {1} da data {2} a data {3}", fi.FullName, fi.CreationTime, path, xi_filename + "*.PDF"));
                                if ((fi.CreationTime <= dTP_COAA.Value.Date.AddHours(23)) && (fi.CreationTime >= dTP_COADA.Value.Date))
                                {
                                    codart = xx.Rows[xi]["ARTICOLOBSC"].ToString();
                                    codartMetodo = codart;
                                    descrizioneMetodo = xx.Rows[xi]["DESCRIZIONEARTICOLO"].ToString();
                                    int.TryParse(xx.Rows[xi]["IDOBJECT_DOC"].ToString(), out IDOBJECT_DOC);
                                    int.TryParse(xx.Rows[xi]["IDOBJECT_CLI"].ToString(), out IDOBJECT_CLI);
                                    cliente = xx.Rows[xi]["CODCLIFOR"].ToString() + " " + xx.Rows[xi]["RAGIONESOCIALE"].ToString();
                                    documento = xx.Rows[xi]["DOCUMENTO"].ToString();
                                    email = xx.Rows[xi]["EMAIL_CLIENTE"].ToString();
                                    lotto = xx.Rows[xi]["LOTTO"].ToString();
                                    DateTime.TryParse(xx.Rows[xi]["DATAINVIOCOA"].ToString(), out datainvioCOA);
                                    int.TryParse(xx.Rows[xi]["IDTESTA"].ToString(), out IDTESTA);
                                    int.TryParse(xx.Rows[xi]["IDRIGA"].ToString(), out IDRIGA);
                                    int.TryParse(xx.Rows[xi]["IDDOC_DDT"].ToString(), out IDDOC_DDT);
                                    FILENAME_DDT = xx.Rows[xi]["FILENAME_DDT"].ToString();

                                    toolStripStatusLabel1.Text = string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length);

                                    Application.DoEvents();

                                    object[] orow = { fi.FullName,
                                                    cliente,
                                                    codart,
                                                    documento,
                                                    email,
                                                    lotto,
                                                    IDOBJECT_DOC,
                                                    xx.Rows[xi]["filenamecoacalc1"].ToString(),
                                                    xx.Rows[xi]["filenamecoacalc2"].ToString(),
                                                    datainvioCOA,
                                                    IDTESTA,
                                                    IDRIGA,
                                                    IDOBJECT_CLI,
                                                    IDDOC_DDT,
                                                    FILENAME_DDT
                                                

                                                };
                                    dt.Rows.Add(orow);
                                }
                            }

                            Cursor.Current = Cursors.Default;
                            Application.DoEvents();

                        }

                    }


                    toolStripProgressBar1.Value++;

                }

            }

            toolStripProgressBar1.Visible = false;

            radGridViewCOA.DataSource = null;
            radGridViewCOA.Rows.Clear();
            radGridViewCOA.DataSource = dt;

            radGridViewCOA.EnableSorting = true;
            radGridViewCOA.SortDescriptors.Add("ARTICOLOBSC", ListSortDirection.Ascending);
            radGridViewCOA.SortDescriptors.Add("CLIENTE", ListSortDirection.Ascending);
            radGridViewCOA.SortDescriptors.Add("DOCUMENTO", ListSortDirection.Ascending);

            Cursor.Current = Cursors.Default;
            Application.DoEvents();

            this.radGridViewCOA.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.DisplayedDataCells);

            lblNrFilesCOA.Text = "Nr Files: " + radGridViewCOA.Rows.Count.ToString();

            MessageBox.Show("Caricamento effettuato!");
            
        }


        private void btnSendMailCOA_Click(object sender, EventArgs e)
        {
            string address = Properties.Settings.Default.sendMailBCCSimulazioneCOA; //;kavanzi@italcom.biz";
            string addressCC = "";  //Properties.Settings.Default.sendMailBCCSimulazione; //"alfredo.deangelo@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string addressBCC = Properties.Settings.Default.sendMailBCCCOA; // "knosmail@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string body = "";
            string subject = "";
            string dettaglioCOA = "";

            bool bOKUpload = false;

            string codclifor = "";
            string codart = "";


            int IdObjectDOC = 0;

            string localfilenameCOA = "";
            string localfileCOA = "";
            string subjectCOA = "";

            // modifiche per invio DDT insieme al COA
            string localfilenameDDT = "";
            string tempPathDownload = Path.Combine(Application.StartupPath, "TEMP");
            if (!Directory.Exists(tempPathDownload))
            {
                Directory.CreateDirectory(tempPathDownload);
            }

            cleanTempFolder(tempPathDownload);

            string msg = "";

            int IdObjectSentMail = 0;


            //Knos
            if (refreshKnosLogin() == false)
                return;

            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = radGridViewCOA.SelectedRows.Count + 1;
            toolStripProgressBar1.Value = 1;
            toolStripProgressBar1.Step = 1;
            toolStripProgressBar1.Visible = true;

            log.LogSomething(string.Format("Nr mail da inviare: {0}", radGridViewCOA.SelectedRows.Count));

            //checkBoxInterrompiInvio.Enabled = true;

            if (radGridViewCOA.SelectedRows.Count > 0)
            {
                msg = string.Format("Procedo con l'invio delle notifiche {0}", radGridViewCOA.SelectedRows.Count);

                if (MessageBox.Show(msg, "Invio Certificati di Analisi", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {

                    for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
                    {
                        toolStripProgressBar1.Value += 1;
                        toolStripProgressBar1.Text = string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridViewCOA.SelectedRows.Count);
                        log.LogSomething(string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridViewCOA.SelectedRows.Count));

                        var attachments = new List<string>();

                        bool allegatoDDT = false;

                        localfileCOA = "";
                        IdObjectDOC = 0;
                        dettaglioCOA = "";

                        codclifor = radGridViewCOA.SelectedRows[i].Cells["CLIENTE"].Value.ToString();
                        codart = radGridViewCOA.SelectedRows[i].Cells["CODART"].Value.ToString();
                        localfileCOA = radGridViewCOA.SelectedRows[i].Cells["FILENAME"].Value.ToString();
                        localfilenameCOA = "[C.O.A.]-" + radGridViewCOA.SelectedRows[i].Cells["FILENAME"].Value.ToString().Substring(radGridViewCOA.SelectedRows[i].Cells["FILENAME"].Value.ToString().LastIndexOf("\\") + 1).Replace("#", "");

                        int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_DOC"].Value.ToString(), out IdObjectDOC);

                        if (chkSimulazioneCOA.Checked == false)
                        {
                            // destinatari reali
                            address = radGridViewCOA.SelectedRows[i].Cells["EMAIL_CLIENTE"].Value.ToString();
                        }

                        log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));


                        if (IdObjectDOC > 0)
                        {
                            subjectCOA = string.Format(Properties.Settings.Default.sendMailCOASubject,
                                radGridViewCOA.SelectedRows[i].Cells["CLIENTE"].Value.ToString());

                            dettaglioCOA = string.Format("\r\n CERTIFICATO DI ANALISI/C.O.A. - ARTICOLO/ITEM: {0} {1} ",
                                radGridViewCOA.SelectedRows[i].Cells["CODART"].Value.ToString(),
                                radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());
                            toolStripStatusLabel1.Text = dettaglioCOA;

                            Application.DoEvents();


                            bOKUpload = kw.UploadFileCertificato(IdObjectDOC, 0, localfileCOA, "COA", localfilenameCOA, 0, "", "");

                            if (bOKUpload == false)
                            {
                                txtLog.Text += string.Format("\r\nNON sono riuscito a ALLEGARE il file della pubblicazione IdObject {0} IdDoc {1} nella cartella {2} con nome {3}", IdObjectDOC, 0, localfileCOA, localfilenameCOA);
                            }

                            attachments.Add(localfileCOA);

                            // allega anche DDT cliente
                            int IdDocDDT = 0;
                            int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDDOC_DDT"].Value.ToString(), out IdDocDDT);
                            localfilenameDDT = radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString().Replace("/", "_") + ".PDF";

                            bool bOKDownload = kw.downloadDoc(IdObjectDOC, IdDocDDT, tempPathDownload, localfilenameDDT);

                            if (bOKDownload == false)
                            { }
                            else
                            {
                                attachments.Add(Path.Combine(tempPathDownload, localfilenameDDT));
                                allegatoDDT = true;
                                dettaglioCOA += string.Format("\r\nDOCUMENTO DI TRASPORTO/DELIVERY NOTE: {0} ----", radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());


                            }



                        }


                        // invio singolo
                        Application.DoEvents();

                        if (bOKUpload)
                        {

                            body = string.Format(Properties.Settings.Default.sendMailCOA, dettaglioCOA);

                            if (IdObjectDOC > 0)
                            {
                                if (Properties.Settings.Default.UseVbs)
                                {
                                    textBoxLOG.Text += "\r\n Invio tramite Lotus";

                                    subject = string.Format("{0}", subjectCOA);
                                    dettaglioCOA = string.Format("CERTIFICATO DI ANALISI/C.O.A. - ARTICOLO/ITEM: {0} {1} ",
                                    radGridViewCOA.SelectedRows[i].Cells["CODART"].Value.ToString(),
                                    radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());

                                    if (allegatoDDT)
                                    {
                                        dettaglioCOA += string.Format("\r\nDOCUMENTO DI TRASPORTO/DELIVERY NOTE: {0} ----", radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());
                                    }


                                    body = string.Format(Properties.Settings.Default.sendMailCOAVbsLotus, dettaglioCOA);

                                    //if (Notifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                    if (Notifica.SendNotifyVBSLotus(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                    {
                                        NotificaCOA cNotifica = new NotificaCOA();

                                        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                                        if (chkSimulazioneCOA.Checked == false)
                                        {
                                            if (IdObjectDOC > 0)
                                            {
                                                // aggirono registro
                                                int idtesta = 0;
                                                int idriga = 0;

                                                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
                                                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

                                                updateDataInvioCOA(idtesta, idriga);
                                            }

                                        }

                                        textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);

                                        radGridViewCOA.ShowRowHeaderColumn = true;

                                        // store della mail inviata
                                        try
                                        {
                                            toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                            IdObjectSentMail = 0;

                                            List<int> links = new List<int>();
                                            links.Add(IdObjectDOC);

                                            Application.DoEvents();
                                            IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()
                                                , System.DateTime.Now
                                                , "coa"
                                                , address
                                                , addressCC
                                                , addressBCC
                                                , subject
                                                , body
                                                , ""
                                                , links
                                                , null
                                                , false
                                                , true
                                                , allegatoDDT);
                                            log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                            textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.Message);
                                        }

                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.DrawFill = true;
                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.DrawFill = true;
                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                                        Application.DoEvents();
                                    }
                                    else
                                    {
                                        log.LogSomething(string.Format("ERRORE - Invio mail NON riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                                        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                    }

                                }
                                else
                                {
                                    if (Properties.Settings.Default.UseCdo)
                                    {
                                        subject = string.Format("{0}", subjectCOA);

                                        NotificaCOA cNotifica = new NotificaCOA();

                                        if (cNotifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                        {
                                            log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                                            if (chkSimulazioneCOA.Checked == false)
                                            {
                                                if (IdObjectDOC > 0)
                                                {
                                                    // aggirono registro
                                                    int idtesta = 0;
                                                    int idriga = 0;

                                                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
                                                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

                                                    updateDataInvioCOA(idtesta, idriga);
                                                }

                                            }
                                            textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);

                                            radGridViewCOA.ShowRowHeaderColumn = true;

                                            // store della mail inviata

                                            toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                            IdObjectSentMail = 0;

                                            List<int> links = new List<int>();
                                            links.Add(IdObjectDOC);

                                            Application.DoEvents();
                                            IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()
                                                , System.DateTime.Now
                                                , "coa"
                                                , address
                                                , addressCC
                                                , addressBCC
                                                , subject
                                                , body
                                                , ""
                                                , links
                                                , null
                                                , false
                                                , true
                                                , allegatoDDT);

                                            log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                            textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);


                                            radGridViewCOA.SelectedRows[i].Cells[1].Style.DrawFill = true;
                                            radGridViewCOA.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                            radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                                            radGridViewCOA.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                                            radGridViewCOA.SelectedRows[i].Cells[2].Style.DrawFill = true;
                                            radGridViewCOA.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                                            radGridViewCOA.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                                            Application.DoEvents();
                                        }
                                        else
                                        {
                                            log.LogSomething(string.Format("ERRORE - Invio mail NON riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                                            textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                                            radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                        }
                                    }
                                }
                            }
                        }

                    }

                    toolStripProgressBar1.Visible = false;
                    checkBoxInterrompiInvio.Enabled = false;

                    MessageBox.Show("Invio completato!");
                }
            }
        }
        //private void btnSendMailCOA_Click(object sender, EventArgs e)
        //{
        //    string address = Properties.Settings.Default.sendMailBCCSimulazioneCOA; //;kavanzi@italcom.biz";
        //    string addressCC = "";  //Properties.Settings.Default.sendMailBCCSimulazione; //"alfredo.deangelo@gmail.com;m.michieletti@zschimmer-schwarz.com";
        //    string addressBCC = Properties.Settings.Default.sendMailBCCCOA; // "knosmail@gmail.com;m.michieletti@zschimmer-schwarz.com";
        //    string body = "";
        //    string subject = "";
        //    string dettaglioCOA = "";

        //    bool bOKUpload = false;

        //    string codclifor = "";
        //    string codart = "";


        //    int IdObjectDOC = 0;

        //    string localfilenameCOA = "";
        //    string localfileCOA = "";
        //    string subjectCOA = "";

        //    // modifiche per invio DDT insieme al COA
        //    string localfilenameDDT = "";
        //    string tempPathDownload = Path.Combine(Application.StartupPath, "TEMP");
        //    if (!Directory.Exists(tempPathDownload))
        //    {
        //        Directory.CreateDirectory(tempPathDownload);
        //    }

        //    cleanTempFolder(tempPathDownload);

        //    string msg = "";

        //    int IdObjectSentMail = 0;


        //    //Knos
        //    if (refreshKnosLogin() == false)
        //        return;

        //    toolStripProgressBar1.Minimum = 0;
        //    toolStripProgressBar1.Maximum = radGridViewCOA.SelectedRows.Count + 1;
        //    toolStripProgressBar1.Value = 1;
        //    toolStripProgressBar1.Step = 1;
        //    toolStripProgressBar1.Visible = true;

        //    log.LogSomething(string.Format("Nr mail da inviare: {0}", radGridViewCOA.SelectedRows.Count));

        //    //checkBoxInterrompiInvio.Enabled = true;

        //    if (radGridViewCOA.SelectedRows.Count > 0)
        //    {
        //        msg = string.Format("Procedo con l'invio delle notifiche {0}", radGridViewCOA.SelectedRows.Count);

        //        if (MessageBox.Show(msg, "Invio Certificati di Analisi", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
        //        {

        //            for (int i = 0; i < radGridViewCOA.SelectedRows.Count; i++)
        //            {
        //                toolStripProgressBar1.Value += 1;
        //                toolStripProgressBar1.Text = string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridViewCOA.SelectedRows.Count);
        //                log.LogSomething(string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridViewCOA.SelectedRows.Count));

        //                var attachments = new List<string>();

        //                bool allegatoDDT = false;

        //                localfileCOA = "";
        //                IdObjectDOC = 0;
        //                dettaglioCOA = "";

        //                codclifor = radGridViewCOA.SelectedRows[i].Cells["CLIENTE"].Value.ToString();
        //                codart = radGridViewCOA.SelectedRows[i].Cells["CODART"].Value.ToString();
        //                localfileCOA = radGridViewCOA.SelectedRows[i].Cells["FILENAME"].Value.ToString();
        //                localfilenameCOA = "[C.O.A.]-" + radGridViewCOA.SelectedRows[i].Cells["FILENAME"].Value.ToString().Substring(radGridViewCOA.SelectedRows[i].Cells["FILENAME"].Value.ToString().LastIndexOf("\\") + 1).Replace("#", "");

        //                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_DOC"].Value.ToString(), out IdObjectDOC);

        //                if (chkSimulazioneCOA.Checked == false)
        //                {
        //                    // destinatari reali
        //                    address = radGridViewCOA.SelectedRows[i].Cells["EMAIL_CLIENTE"].Value.ToString();
        //                }

        //                log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));


        //                if (IdObjectDOC > 0)
        //                {
        //                    subjectCOA = string.Format(Properties.Settings.Default.sendMailCOASubject, 
        //                        radGridViewCOA.SelectedRows[i].Cells["CLIENTE"].Value.ToString());

        //                    dettaglioCOA = string.Format("\r\n CERTIFICATO DI ANALISI/C.O.A. - ARTICOLO/ITEM: {0} {1} ", 
        //                        radGridViewCOA.SelectedRows[i].Cells["CODART"].Value.ToString(), 
        //                        radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());
        //                    toolStripStatusLabel1.Text = dettaglioCOA;

        //                    Application.DoEvents();


        //                    bOKUpload = kw.UploadFileCertificato(IdObjectDOC, 0, localfileCOA, "COA", localfilenameCOA,0, "", "");

        //                    if (bOKUpload == false)
        //                    {
        //                        txtLog.Text += string.Format("\r\nNON sono riuscito a ALLEGARE il file della pubblicazione IdObject {0} IdDoc {1} nella cartella {2} con nome {3}", IdObjectDOC, 0, localfileCOA, localfilenameCOA);
        //                    }

        //                    attachments.Add(localfileCOA);

        //                    // allega anche DDT cliente
        //                    int IdDocDDT = 0;
        //                    int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDDOC_DDT"].Value.ToString(), out IdDocDDT);
        //                    localfilenameDDT = radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString().Replace("/", "_") + ".PDF";

        //                    bool bOKDownload = kw.downloadDoc(IdObjectDOC, IdDocDDT, tempPathDownload, localfilenameDDT);

        //                    if (bOKDownload == false)
        //                    { }
        //                    else
        //                    {
        //                        attachments.Add(Path.Combine(tempPathDownload, localfilenameDDT));
        //                        allegatoDDT = true;
        //                        dettaglioCOA += string.Format("\r\nDOCUMENTO DI TRASPORTO/DELIVERY NOTE: {0} ----", radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());


        //                    }



        //                }


        //                // invio singolo
        //                Application.DoEvents();

        //                if (bOKUpload)
        //                {

        //                    body = string.Format(Properties.Settings.Default.sendMailCOA, dettaglioCOA);

        //                    if (IdObjectDOC > 0)
        //                    {


        //                        if (Properties.Settings.Default.UseLotus)
        //                        {

        //                            subject = string.Format("{0}", subjectCOA);


        //                            Notifica cNotifica = new ToDoNotificheBSC.Notifica();
        //                            if (cNotifica.SendNotifyLotus(address, subject, body, attachments, null, checkBoxPopUpMailCOA.Checked, addressCC, addressBCC) == true)
        //                            {

        //                                log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

        //                                if (chkSimulazioneCOA.Checked == false)
        //                                {
        //                                    if (IdObjectDOC > 0)
        //                                    {
        //                                        // aggirono registro
        //                                        int idtesta = 0;
        //                                        int idriga = 0;

        //                                        int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
        //                                        int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

        //                                        updateDataInvioCOA(idtesta, idriga);
        //                                    }

        //                                }
        //                                textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);

        //                                radGridViewCOA.ShowRowHeaderColumn = true;

        //                                // store della mail inviata

        //                                toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
        //                                IdObjectSentMail = 0;

        //                                List<int> links = new List<int>();
        //                                links.Add(IdObjectDOC);

        //                                Application.DoEvents();
        //                                IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()
        //                                    , System.DateTime.Now
        //                                    , "coa"
        //                                    , address
        //                                    , addressCC
        //                                    , addressBCC
        //                                    , subject
        //                                    , body
        //                                    , ""
        //                                    , links
        //                                    , null
        //                                    , false
        //                                    , true
        //                                    , allegatoDDT);
        //                                log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
        //                                textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);


        //                                radGridViewCOA.SelectedRows[i].Cells[1].Style.DrawFill = true;
        //                                radGridViewCOA.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
        //                                radGridViewCOA.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
        //                                radGridViewCOA.SelectedRows[i].Cells[2].Style.DrawFill = true;
        //                                radGridViewCOA.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                radGridViewCOA.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

        //                                Application.DoEvents();
        //                            }
        //                            else
        //                            {
        //                                log.LogSomething(string.Format("ERRORE - Invio mail NON riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
        //                                textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
        //                                radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
        //                            }



        //                        }
        //                        else
        //                        {


        //                            if (Properties.Settings.Default.UseVbs)
        //                            {
        //                                textBoxLOG.Text += "\r\n Invio tramite Lotus";

        //                                subject = string.Format("{0}", subjectCOA);
        //                                dettaglioCOA = string.Format("CERTIFICATO DI ANALISI/C.O.A. - ARTICOLO/ITEM: {0} {1} ",
        //                                radGridViewCOA.SelectedRows[i].Cells["CODART"].Value.ToString(),
        //                                radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());

        //                                if (allegatoDDT)
        //                                {
        //                                    dettaglioCOA += string.Format("\r\nDOCUMENTO DI TRASPORTO/DELIVERY NOTE: {0} ----", radGridViewCOA.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString());
        //                                }


        //                                body = string.Format(Properties.Settings.Default.sendMailCOAVbsLotus, dettaglioCOA);

        //                                //if (Notifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
        //                                if (Notifica.SendNotifyVBSLotus(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
        //                                {
        //                                    NotificaCOA cNotifica = new NotificaCOA();

        //                                    if (cNotifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
        //                                    {
        //                                        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

        //                                        if (chkSimulazioneCOA.Checked == false)
        //                                        {
        //                                            if (IdObjectDOC > 0)
        //                                            {
        //                                                // aggirono registro
        //                                                int idtesta = 0;
        //                                                int idriga = 0;

        //                                                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
        //                                                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

        //                                                updateDataInvioCOA(idtesta, idriga);
        //                                            }

        //                                        }
        //                                        textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);

        //                                        radGridViewCOA.ShowRowHeaderColumn = true;

        //                                        // store della mail inviata

        //                                        toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
        //                                        IdObjectSentMail = 0;

        //                                        List<int> links = new List<int>();
        //                                        links.Add(IdObjectDOC);

        //                                        Application.DoEvents();
        //                                        IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()
        //                                            , System.DateTime.Now
        //                                            , "coa"
        //                                            , address
        //                                            , addressCC
        //                                            , addressBCC
        //                                            , subject
        //                                            , body
        //                                            , ""
        //                                            , links
        //                                            , null
        //                                            , false
        //                                            , true
        //                                            , allegatoDDT);
        //                                        log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
        //                                        textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);


        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.DrawFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.DrawFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

        //                                        Application.DoEvents();
        //                                    }
        //                                    else
        //                                    {
        //                                        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
        //                                        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
        //                                    }
        //                                }
        //                            }
        //                            else
        //                            {
        //                                if (Properties.Settings.Default.UseCdo)
        //                                {
        //                                    subject = string.Format("{0}", subjectCOA);

        //                                    NotificaCOA cNotifica = new NotificaCOA();

        //                                    if (cNotifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
        //                                    {
        //                                        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

        //                                        if (chkSimulazioneCOA.Checked == false)
        //                                        {
        //                                            if (IdObjectDOC > 0)
        //                                            {
        //                                                // aggirono registro
        //                                                int idtesta = 0;
        //                                                int idriga = 0;

        //                                                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDTESTA"].Value.ToString(), out idtesta);
        //                                                int.TryParse(radGridViewCOA.SelectedRows[i].Cells["IDRIGA"].Value.ToString(), out idriga);

        //                                                updateDataInvioCOA(idtesta, idriga);
        //                                            }

        //                                        }
        //                                        textBoxLOG.Text += string.Format("\r\n OK {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);

        //                                        radGridViewCOA.ShowRowHeaderColumn = true;

        //                                        // store della mail inviata

        //                                        toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
        //                                        IdObjectSentMail = 0;

        //                                        List<int> links = new List<int>();
        //                                        links.Add(IdObjectDOC);

        //                                        Application.DoEvents();
        //                                        IdObjectSentMail = kw.StoreEmailSent(3, "2", radGridViewCOA.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()
        //                                            , System.DateTime.Now
        //                                            , "coa"
        //                                            , address
        //                                            , addressCC
        //                                            , addressBCC
        //                                            , subject
        //                                            , body
        //                                            , ""
        //                                            , links
        //                                            , null
        //                                            , false
        //                                            , true
        //                                            , allegatoDDT);

        //                                        log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
        //                                        textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);


        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.DrawFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.DrawFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

        //                                        Application.DoEvents();
        //                                    }
        //                                    else
        //                                    {
        //                                        log.LogSomething(string.Format("ERRORE - Invio mail NON riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
        //                                        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    //if (Notifica.SendNotifyMAPI(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
        //                                    if (Notifica.SendNotifyMAPILotus(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
        //                                    {
        //                                        radGridViewCOA.ShowRowHeaderColumn = true;

        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.DrawFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.DrawFill = true;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
        //                                        radGridViewCOA.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

        //                                        Application.DoEvents();
        //                                        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
        //                                        textBoxLOG.Text += string.Format("\r\n - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC);

        //                                    }
        //                                    else
        //                                    {
        //                                        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
        //                                        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
        //                                        radGridViewCOA.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
        //                                    }
        //                                }
        //                            }

        //                        }
        //                    }
        //                }

        //            }

        //            toolStripProgressBar1.Visible = false;
        //            checkBoxInterrompiInvio.Enabled = false;

        //            MessageBox.Show("Invio completato!");
        //        }
        //    }
        //}

        private void updateDataInvioCOA(int idt, int idr)
        {
            string dt = System.DateTime.Today.ToShortDateString();

            string strSQL = string.Format("UPDATE EXTRARIGHEDOC SET DATAINVIOCOA = getdate() WHERE IDTESTA = {0} AND IDRIGA = {1}", idt, idr);


            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(strSQL))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    cmd.ExecuteNonQuery();
                }

            }
        }


        private void radGridViewCOA_CurrentRowChanging(object sender, Telerik.WinControls.UI.CurrentRowChangingEventArgs e)
        {
            if (radGridViewCOA.Rows.Count > 0)
            {
                if (e.NewRow.Index >= 0)
                    datirigaCOA(e.NewRow.Index);
            }
        }


        private void datirigaCOA(int r)
        {
            string tooltip = "Cliente: {0} | " +
                "Articolo: {1} \r\n" +
                "Indirizzo email: {2} | " +
                "Documento: {3} \r\n" +
                "File: {4} | " +
                "Ultimo invio: {5} \r\n";

            string cliente = radGridViewCOA.Rows[r].Cells["CLIENTE"].Value.ToString();// string.Format("{0} - {1}", radGridViewCOA.Rows[r].Cells["CODCLIFOR"].Value.ToString(), radGridViewCOA.Rows[r].Cells["RAGIONESOCIALE"].Value.ToString());
            string articolo = radGridViewCOA.Rows[r].Cells["CODART"].Value.ToString();

            string address = radGridViewCOA.Rows[r].Cells["EMAIL_CLIENTE"].Value.ToString();
            string addressCC = radGridViewCOA.Rows[r].Cells["DOCUMENTO"].Value.ToString();

            string fileSCH = radGridViewCOA.Rows[r].Cells["FILENAME"].Value.ToString();
            string ultimoinvioSCH = "";
            if (radGridViewCOA.Rows[r].Cells["DATAINVIOCOA"].Value != null)
                ultimoinvioSCH = radGridViewCOA.Rows[r].Cells["DATAINVIOCOA"].Value.ToString();

            textBoxToolTipCOA.Text = string.Format(tooltip, cliente, articolo, address, addressCC, fileSCH, ultimoinvioSCH);
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

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {

                try
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand(commandtext, cn))
                    {

                        for (int i = 0; i < radGridViewAllegatiSS.SelectedRows.Count; i++)
                        {
                            if (radGridViewAllegatiSS.SelectedRows[i].Cells[0].Value.ToString() != null)
                            {
                                foreach(Allegato a in listaallegati)
                                {
                                    
                                    Application.DoEvents();

                                    toolStripStatusLabel1.Text = string.Format("Allego file: {0} alla pubblicazione con IdObject: {1}", a.Path, radGridViewAllegatiSS.SelectedRows[i].Cells["IDOBJECT_SCH"].Value.ToString());

                                    cmd.CommandText = string.Format(commandtext, radGridViewAllegatiSS.SelectedRows[i].Cells["IDOBJECT_SCH"].Value.ToString(), a.Path, a.Descrizione);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                        }

                        toolStripStatusLabel1.Text = string.Format("Caricamento completato");
                    }

                }

                catch (SqlException ex)
                {
                    MessageBox.Show(string.Format("Errore SQL SERVER: {0} - {1}", Properties.Settings.Default.MetodoConnectionString, ex.Message));

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

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {

                try
                {
                    cn.Open();

                    toolStripStatusLabel1.Text = string.Format("caricamento dati in corso..........");

                    radGridViewAllegatiSS.EnableFiltering = false;
                    radGridViewAllegatiSS.ShowFilteringRow = false;

                    using (SqlCommand cmd = new SqlCommand(commandtext, cn))
                    {

                        //cmd.Parameters.AddWithValue("DATADOC", dateTimePickerDa.Value);
                        //cmd.ExecuteNonQuery();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        da.Fill(dt);

                        radGridViewAllegatiSS.DataSource = dt;

                        for (int i = 0; i < radGridViewAllegatiSS.Columns.Count; i++)
                        {
                            if (radGridViewAllegatiSS.Columns[i].FieldName.StartsWith("IDOBJECT"))
                            {
                                radGridViewAllegatiSS.Columns[i].IsVisible = false;
                            }
                            else
                            {
                                radGridViewAllegatiSS.Columns[i].BestFit();
                            }
                        }

                        radGridViewAllegatiSS.AutoScroll = true;
                        radGridViewAllegatiSS.Refresh();

                        toolStripStatusLabel1.Text = string.Format("Caricamento completato");

                    }


                    radGridViewAllegatiSS.EnableFiltering = true;
                    radGridViewAllegatiSS.ShowFilteringRow = true;
                    radGridViewAllegatiSS.EnableAlternatingRowColor = true;
                    radGridViewAllegatiSS.MultiSelect = true;

                }

                catch (SqlException ex)
                {
                    MessageBox.Show(string.Format("Errore SQL SERVER: {0} - {1}", Properties.Settings.Default.MetodoConnectionString, ex.Message));

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

                dataGridView1.Rows.Add(fi.Name, fi.Name, p);
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
            panel1.Visible = chkAllegati.Checked;
        }

        private void radGridViewCOA_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            int IdObjectBOL = 0;
            int IdDocBOL = 0;
            int IdObjectSCH = 0;
            int IdDocSCH = 0;


            //http://vsrv2k8bsn2:8780/KnoS_Catalog/0/0000035964/0001/1426089157015/Zetesol%20MGS.doc
            string url = "{0}";

            if (e.ColumnIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (radGridViewCOA.Rows[e.RowIndex].Cells["filename"].ColumnInfo.Index == e.ColumnIndex)
                {
                    url = string.Format(url, radGridViewCOA.Rows[e.RowIndex].Cells["filename"].Value.ToString());
                    //url = url.Replace("#", "_");

                    webBrowserCOA.Navigate(url);
                }


                if (radGridViewCOA.Columns[e.ColumnIndex].HeaderText.ToUpper() == "DOCUMENTO")
                {
                    url = "{0}/KnoS_Catalog/0/{1}/{2}/{3}";
                    url = string.Format(url, txtKnosUrl.Text, radGridViewCOA.Rows[e.RowIndex].Cells["IDOBJECT_DOC"].Value.ToString(), radGridViewCOA.Rows[e.RowIndex].Cells["IDDOC_DDT"].Value.ToString(), radGridViewCOA.Rows[e.RowIndex].Cells["FILENAME_DDT"].Value.ToString().Substring(0, 15) + ".PDF");
                    url = url.Replace("#", "_");

                    webBrowserCOA.Navigate(url);
                }


            }
            else
            {
                // selezione celle
                Debug.Print(string.Format("ci: {0}  - ri: {1}", e.ColumnIndex, e.RowIndex));

                if ((e.RowIndex == -1) && (e.ColumnIndex == -1))
                {
                    radGridView1.SelectAll();
                }

                if ((e.RowIndex > -1) && (e.ColumnIndex == -1))
                {

                    if (e.Row.Group != null)
                    {
                        for (int x = 0; x < e.Row.Parent.ChildRows.Count; x++)
                        {

                            e.Row.Parent.ChildRows[x].IsSelected = true;
                            Application.DoEvents();

                        }


                    }

                }




            }
        }

        private void radGridViewEpy_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {

            string url = "{0}";

            if (e.ColumnIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (radGridViewEpy.Rows[e.RowIndex].Cells["filename"].ColumnInfo.Index == e.ColumnIndex)
                {
                    url = string.Format(url, Path.Combine(lblPathEpy.Text, radGridViewEpy.Rows[e.RowIndex].Cells["lingua"].Value.ToString(), radGridViewEpy.Rows[e.RowIndex].Cells["filename"].Value.ToString()));
                    //url = url.Replace("#", "_");

                    webBrowserSSEPY.Navigate(url);
                }


            }
            else
            {
                // selezione celle
                radGridViewEpy.MultiSelect = true;
                if ((e.RowIndex == -1) && (e.ColumnIndex == -1))
                {
                    radGridViewEpy.SelectAll();
                }

                if ((e.RowIndex > -1) && (e.ColumnIndex == -1))
                {

                    if (e.Row.Group != null)
                    {
                        for (int x = 0; x < e.Row.Parent.ChildRows.Count; x++)
                        {

                            e.Row.Parent.ChildRows[x].IsSelected = true;
                            Application.DoEvents();

                        }


                    }

                }




            }
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
                MessageBox.Show(string.Format("Sito KnoS non trovato o non accessibile!", Properties.Settings.Default.KnoS_URL), "Inizializzazione programma firma", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            
        }

        private void radGridViewEpy_Click(object sender, EventArgs e)
        {

        }

        private void chkSimulazione_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxPopUpMail_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxInterrompiInvio_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnSalvaImpostazioni_Click(object sender, EventArgs e)
        {
            clsRadGridSettings.SaveColumnsSettings(radGridView1, s, "SchedeSicurezza");            
        }

        private void cmbI_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                string sf = cmbImpostazioni.SelectedItem.ToString();
                radGridView1.LoadLayout(sf);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delle righe ordine {0}", ex.Message));
            }
        }

        private void btnArchivioRMI_Click(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;

            System.DateTime dtUltimoAggiornamento;

            Application.DoEvents();

            log.LogSomething("Caricamento articoli RMI");

            DataTable xx = getArticoliRMI();

            log.LogSomething("Caricamento articoli schede COMPLETATO");

            List<string> pathlist = new List<string>();

            string filename = "";
            string codart = "";
            string revisione = "";
            string lingua = "";
            DateTime datamodifica = System.DateTime.Now;
            string codartMetodo = "";
            string descrizioneMetodo = "";
            int IDOBJECT = 0;
            string subpath = "";
            string nomeEPY = "";

            DataTable dt = new DataTable();
            dt.Columns.Add("filename");
            dt.Columns.Add("codart");
            dt.Columns.Add("revisione");
            dt.Columns.Add("lingua");
            dt.Columns.Add("datamodifica");
            dt.Columns.Add("codartMetodo");
            dt.Columns.Add("descrizioneMetodo");
            dt.Columns.Add("IDOBJECT");
            dt.Columns.Add("subpath");
            dt.Columns.Add("nomeEPY");

            int i = 0;

            Application.DoEvents();

            foreach (string path in Properties.Settings.Default.PathRMI)
            {
                if (Directory.Exists(path))
                {
                    lblPathRMI.Text = path;

                    log.LogSomething("Inizio ricerca file RMI");

                    toolStripStatusLabel1.Text = string.Format("ciclo sugli articoli che ho trovato {0}", xx.Rows.Count.ToString());
                    Application.DoEvents();

                    // ciclo sugli articoli che ho trovato
                    for (int xi = 0; xi < xx.Rows.Count; xi++)
                    {
                        string xi_filename = xx.Rows[xi]["NOMEEPY"].ToString();

                        if (xi_filename != "")
                        {
                            log.LogSomething("Inizio ricerca file EPY  * " + xi_filename + " * ");

                            string[] files = Directory.GetFiles(path, "*" + xi_filename + "*", SearchOption.AllDirectories);

                            foreach (string f in files)
                            {
                                i += 1;



                                FileInfo fi = new FileInfo(f);
                                string.Format("file: {0}", fi.FullName);

                                //if (fi.LastWriteTime > dtUltimoAggiornamento)
                                //{
                                filename = Path.GetFileName(fi.FullName);
                                string[] fname = filename.Split('_');
                                filename = fi.Name;

                                lingua = "";
                                codart = xx.Rows[xi]["CODICE"].ToString();
                                revisione = "";
                                nomeEPY = "";

                                //if (fname.Count() == 4)
                                //    nomeEPY = fname[3];

                                datamodifica = fi.LastWriteTime;
                                subpath = Path.GetFullPath(fi.FullName);

                                // articolo Metodo
                                //List<string> artmetodo = getArticoloMetodo(codart);

                                codartMetodo = codart;
                                descrizioneMetodo = xx.Rows[xi]["descrizioneMetodo"].ToString();
                                int.TryParse(xx.Rows[xi]["IDOBJECT"].ToString(), out IDOBJECT);

                                toolStripStatusLabel1.Text = string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length);

                                log.LogSomething(string.Format("file: {0}  nr {1} - {2}", filename, i, files.Length));

                                Application.DoEvents();

                                object[] orow = { filename, codart, revisione, lingua, datamodifica, codartMetodo, descrizioneMetodo, IDOBJECT, subpath, nomeEPY };
                                dt.Rows.Add(orow);
                                //}

                            }

                        }

                    }

                }



            }


            radGridViewRMI.DataSource = null;
            radGridViewRMI.Rows.Clear();
            radGridViewRMI.DataSource = dt;

            radGridViewRMI.EnableFiltering = true;
            radGridViewRMI.EnableSorting = true;
            radGridViewRMI.SortDescriptors.Add("codart", ListSortDirection.Ascending);
            radGridViewRMI.SortDescriptors.Add("lingua", ListSortDirection.Ascending);
            radGridViewRMI.SortDescriptors.Add("revisione", ListSortDirection.Ascending);


            this.radGridViewRMI.BestFitColumns(Telerik.WinControls.UI.BestFitColumnMode.DisplayedDataCells);

            lblNrFilesRMI.Text = "Nr Files: " + radGridViewRMI.Rows.Count.ToString();

            MessageBox.Show("Caricamento effettuato!");
            //tabControl1.SelectedIndex = 3;

            if (this.Visible == false)
                this.Visible = true;
        }

        DataTable getArticoliRMI()
        {
            DataTable x = new DataTable();

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(string.Format("SELECT * FROM VISTA_RELAZIONIARTICOLIRMI")))
                {
                    cn.Open();
                    cmd.Connection = cn;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);

                    da.Fill(x);

                }

            }
            return x;
        }
        

        private void radGridViewRMI_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {


            string url = "{0}";

            if (e.ColumnIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (radGridViewRMI.Rows[e.RowIndex].Cells["filename"].ColumnInfo.Index == e.ColumnIndex)
                {
                    url = string.Format(url, Path.Combine(lblPathRMI.Text, radGridViewEpy.Rows[e.RowIndex].Cells["filename"].Value.ToString()));
                    //url = url.Replace("#", "_");

                    webBrowserRMI.Navigate(url);
                }


            }
            //else
            //{
            //    // selezione celle
            //    radGridViewRMI.MultiSelect = true;
            //    if ((e.RowIndex == -1) && (e.ColumnIndex == -1))
            //    {
            //        radGridViewRMI.SelectAll();
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

        private void radGridViewRMI_CellClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {


            string url = "{0}";

            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                if (radGridViewRMI.Rows[e.RowIndex].Cells["subpath"].ColumnInfo.Index == e.ColumnIndex)
                {
                    url = string.Format(url, Path.Combine(lblPathRMI.Text, radGridViewRMI.Rows[e.RowIndex].Cells["subpath"].Value.ToString()));
                    //url = url.Replace("#", "_");

                    webBrowserRMI.Navigate(url);
                }


            }
        }

        private void btnInvioKnosRMI_Click(object sender, EventArgs e)
        {
            string filename;
            string descrizioneMetodo;
            string codlingua;
            string revisione;
            string subpath = "";
            string nCodCliente = "";

            int IdObjectCliAll = 38750;
            int IdObjectCli = 0;
            int IdObjectArticolo = 0;
            int IdObjectScheda = 0;
            IKnosObjectSelector os = KnosInstance.Client.CreateKnosObjectSelector();

            List<int> schedeaggiornate = new List<int>();


            //Knos
            if (refreshKnosLogin() == false)
                return;

            string searchExpression = " IdClass = 5013 and exists(select 1 from object_linkage x where x.idparent = _idobject_ and idchild = {0}) and exists(select 1 from object_linkage x where x.idparent = _idobject_ and idchild = {1})";

            string searchExpressionCLI = " IdClass = 2 and varchar_04 = 'C' + RIGHT('      ' +  '{0}', 6)";

            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.Minimum = 1;

            for (int i = 0; i < radGridViewRMI.SelectedRows.Count; i++)
            {
                toolStripProgressBar1.Maximum = radGridViewRMI.RowCount;
                toolStripProgressBar1.Step = 1;
                toolStripProgressBar1.PerformStep();

                IdObjectArticolo = 0;
                IdObjectScheda = 0;
                IdObjectCli = 0;

                filename = descrizioneMetodo = codlingua = "";

                os.Reset(true);


                Application.DoEvents();

                int.TryParse(radGridViewRMI.SelectedRows[i].Cells["IDOBJECT"].Value.ToString(), out IdObjectArticolo);
                filename = radGridViewRMI.SelectedRows[i].Cells["filename"].Value.ToString().Replace("#", "");
                string codartMetodo = radGridViewRMI.SelectedRows[i].Cells["codartMetodo"].Value.ToString();
                descrizioneMetodo = radGridViewRMI.SelectedRows[i].Cells["descrizioneMetodo"].Value.ToString();
                codlingua = radGridViewRMI.SelectedRows[i].Cells["lingua"].Value.ToString();
                revisione = radGridViewRMI.SelectedRows[i].Cells["revisione"].Value.ToString();
                subpath = radGridViewRMI.SelectedRows[i].Cells["subpath"].Value.ToString();

                string title = string.Format("{0} [{1}]", descrizioneMetodo, codartMetodo);

                Application.DoEvents();
                toolStripStatusLabel1.Text = string.Format("Allego file: {0}, {1}, {2} ({3}/{4})", filename, descrizioneMetodo, codlingua, i, radGridViewRMI.RowCount);

                if (IdObjectArticolo > 0)
                {
                    // gestione cliente specifico
                    nCodCliente = filename.Substring(filename.LastIndexOf('_')+1).ToLower().Replace(".pdf", "");

                    log.LogSomething("Codice cliente: " + string.Format(searchExpressionCLI, nCodCliente));

                    if (nCodCliente != "")
                    {
                        os.SearchExpression = string.Format(searchExpressionCLI, nCodCliente);
                        os.GetPage(1);

                        if (os.RecordCount == 1)
                        {
                            IdObjectCli = os.GetItem(0).IdObject;
                        }

                        os.Reset(true);
                    }

                    if (IdObjectCli > 0)
                    {
                        os.SearchExpression = string.Format(searchExpression, IdObjectCli, IdObjectArticolo);
                    }
                    else
                    {
                        os.SearchExpression = string.Format(searchExpression, IdObjectCliAll, IdObjectArticolo);
                    }

                    os.GetPage(1);

                    if (os.RecordCount == 1)
                    {
                        // scheda esistente
                        IdObjectScheda = os.GetItem(0).IdObject;

                        if (schedeaggiornate != null)
                        {
                            if (schedeaggiornate.Contains(IdObjectScheda) == false)
                            {
                                kw.DeleteFiles(IdObjectScheda, -1);
                                schedeaggiornate.Add(IdObjectScheda);

                                if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateRMI, "") == false)
                                {
                                    if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateRMIRedazione, "") == false)
                                    {
                                        log.LogSomething(string.Format("Errore nella transizione di stato della scheda! {0}", IdObjectScheda));
                                    }
                                    else
                                    {
                                        log.LogSomething(string.Format("Eseguita transizione di stato della scheda! {0} {1}", IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateRMIRedazione));
                                    }
                                }
                                else
                                {
                                    log.LogSomething(string.Format("Eseguita transizione di stato della scheda! {0} {1}", IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateRMI));
                                }

                            }

                        }
                    }
                    else
                    {
                        // la creo
                        IKnosMultivalueEditor meC = KnosInstance.Client.CreateKnosMultivalueEditor();
                        if (IdObjectCli > 0)
                        {
                            meC.AddValue(IdObjectCli);
                        }
                        else
                        {
                            meC.AddValue(IdObjectCliAll);
                        }
                            
                        IKnosMultivalueEditor meA = KnosInstance.Client.CreateKnosMultivalueEditor();
                        meA.AddValue(IdObjectArticolo);

                        IKnosObjectMaker om = KnosInstance.Client.CreateKnosObjectMaker();
                        om.IdClass = 5013;
                        om.SetAttrValue("object_19", meC, EnumKnosDataType.ObjectListType);
                        om.SetAttrValue("object_5036", meA, EnumKnosDataType.ObjectListType);
                        om.SetAttrValue("title", title);
                        om.CreateObject(out IdObjectScheda);


                    }

                    if (IdObjectScheda == 0)
                    {
                        MessageBox.Show("Nessuna scheda trovata o creata!");
                    }
                    else
                    {

                        //upload file
                        if (kw.UploadFileCertificato(IdObjectScheda, 0, subpath, string.Format("{0}-{1}", codlingua, descrizioneMetodo), filename, 0, "", revisione))
                        {
                            radGridViewRMI.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                            Application.DoEvents();
                        }


                        if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateRMI, "") == false)
                        {
                            if (kw.EseguiAzioneWS(IdObjectScheda, Properties.Settings.Default.KnoS_IdActionUpdateRMIRedazione, "") == false)
                            {
                                log.LogSomething(string.Format("Errore nella transizione di stato della scheda! {0}", IdObjectScheda));
                            }
                        }

                    }




                }




            }

            // aggiornamento registro schede
            aggiornaRegistro();


            //DateTime dtUltimoAggiornamento;
            //DateTime.TryParse(txtUltimoAggiornamento.Text, out dtUltimoAggiornamento);
            //Properties.Settings.Default.DataAggiornamento = dtUltimoAggiornamento;
            //Properties.Settings.Default.Save();

            MessageBox.Show("Aggiornamento completato!", "Importazione schede da Epy", MessageBoxButtons.OK);
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
                MessageBox.Show(string.Format("Si è verificato un errore nel caricamento del layout della tabella dei delle righe ordine {0}", ex.Message));
            }

        }

        void check_columns()
        {
            try
            {
                string sXML = Path.Combine(s, System.Environment.UserName + ".xml");
                string tXML = Path.Combine(s, "template_msds" + ".xml");

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
            catch (Exception ex)
            {

                log.LogSomething(string.Format("ERRORE IN CARICAMENTO IMPOSTAZIONI COLONNE \r\n{0}", ex.Message));

            }

        }

        private void radGridViewCOA_CellFormatting(object sender, Telerik.WinControls.UI.CellFormattingEventArgs e)
        {
            if ((e.ColumnIndex > 0) && (e.Row.Cells["IDOBJECT_DOC"].Value.ToString() == "0"))
            {
                e.CellElement.DrawFill = true;
                e.CellElement.BackColor = Color.Red;
                e.CellElement.NumberOfColors = 1;
            }
            else
            {
                e.CellElement.DrawFill = true;
                e.CellElement.BackColor = Color.White;
                e.CellElement.NumberOfColors = 1;
            }
        }

        private void cmbPathEpy_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblPathEpy.Text = cmbPathEpy.Items[cmbPathEpy.SelectedIndex].ToString();
        }

        private void cmdPathH_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblPathH.Text = cmbPathH.Items[cmbPathH.SelectedIndex].ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveGridSettings(radGridView1);
        }
    }


}

