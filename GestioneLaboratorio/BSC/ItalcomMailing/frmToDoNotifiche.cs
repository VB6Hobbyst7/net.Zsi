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

namespace ToDoNotificheBSC
{

    public partial class frmToDoNotifiche : Form
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

        List<Allegato> allegati = new List<Allegato>();

        public class KnoSWrapper 
        {
            

            IKnosObject knosObject ;
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
                                lvAttr.Items[lvAttr.Items.Count - 1].SubItems[1].Text = string.Format("{0} - {1}",  knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_04").ToString(),  knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_05").ToString());
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

                                    fileName =  fileUrl = fileDescr = fileLocalPath = "";
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
                                    lvAttr.Items[lvAttr.Items.Count - 1].SubItems[1].Text = string.Format("{0} - {1}",  knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_04").ToString(),  knosObjectCliente.AttrValueList.GetItemByColumnName("varchar_05").ToString());
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
                                s.Text = string.Format("Caricamento certificati in corso..... ({0}/{1})", (i+1).ToString(), nrCertificatiTot.ToString());
                                // verifico attributi della pubblicazione crtificato per poter capire che tipo di azione può effettuare l'utente concui 
                                // ci si è loggati a KnoS

                                UtenteTecnico = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_51").ToString();
                                try
                                {
                                    DataPrimaFirma = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("datetime_08").ToString();
                                }
                                catch(Exception ex)
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
                    if(_idDoc >= 0)
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

                if ((_idObject > 0) && (_idDoc>0))
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

            public bool downloadDoc(int _idCertificato, string _fnPattern = "", string filePath = "", string filename = "")
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
                        if (knosObject.DocumentList.GetItem(i).FileName.StartsWith(_fnPattern.Substring(17)))
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


            public bool AddDestinatarioCapoCommessa(int _idObject)
            {
                bool bAddDestinatario = false;
                int _idSubject = 0;
                bool retvalue = false;

                _idSubject = GetIdSubjectByName(ToDoNotificheBSC.frmToDoNotifiche.CurrentCapoCommessa);

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


            public bool GetSignImage(int _idObject, ListView lvFirme, string _signer )
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
                DataPrimaFirma =  CapoCommessaSost = ""; //DataSecondaFirma = ResponsabileTecnicoSost = 

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
                    DateTime.TryParse(DataChiusuraPDL,out  x);

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
                , List<string> IdObjectCliList
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
                )
            {
                int idobjectMail = 0;
                int idObjectCLI = 0;
                
                IKnosObjectMaker kom = KnosInstance.Client.CreateKnosObjectMaker();

                kom.Reset();

                kom.IdClass = idClass;
                kom.SetAttrValue("enum_27", Tipo);

                IKnosMultivalueEditor kme = KnosInstance.Client.CreateKnosMultivalueEditor();
                if (IdObjectCliList != null)
                {
                    foreach (string IdObjectCli in IdObjectCliList)
                    {

                        int.TryParse(IdObjectCli, out idObjectCLI);
                        kme.AddValue(idObjectCLI);
                    }
                }

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
                kom.SetAttrValue("smallint_03", true);


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

       
 
        public frmToDoNotifiche()
        {
            InitializeComponent();
        }


        
        class itemConf
        {
            string _key;
            string _sql;
            string _oggi;


            public string Key
            {
                get
                {
                    return this._key;
                }
                set
                {
                    this._key = value;
                }
            }

            public string Sql
            {
                get
                {
                    return this._sql;
                }
                set
                {
                    this._sql = value;
                }
            }

            public string Oggi
            {
                get
                {
                    return this._oggi;
                }
                set
                {
                    this._oggi = value;
                }
            }
        }

        private void frmToDoNotificheBSC_Load(object sender, EventArgs e)
        {

            log = new Logger();
            log.Setup();
            log.LogSomething("Start servizio");

            this.Text = string.Format("Invio notifiche ({0})", Application.ProductVersion);

            try
            {
           
                opened = false;

                bool.TryParse(Properties.Settings.Default.sendMailPopUp, out notifyPopUp);

                //Knos
                if (kw.Inizializza(Properties.Settings.Default.KnoS_URL) == true)
                {
                    //kw.CurrentUser = Properties.Settings.Default.KnoS_User;
                    //kw.PWD = Properties.Settings.Default.KnoS_PWD;

                    txtKnosUrl.Text = Properties.Settings.Default.KnoS_URL;
                    txtKnoSUser.Text = kw.CurrentUser;

                    Application.DoEvents();


                    txtKnosUrl.ReadOnly = txtKnoSUser.ReadOnly = txtKnoSPassword.ReadOnly = true;
                    btnKnoSLogin.Enabled = false;

                    statusStrip1.Text = string.Format("");
                }
                else
                {
                    txtKnosUrl.Text = Properties.Settings.Default.KnoS_URL;
                    MessageBox.Show(string.Format("Sito KnoS non trovato o non accessibile!", Properties.Settings.Default.KnoS_URL), "Inizializzazione programma firma", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                Application.DoEvents();


                // gestione invio allegati Knos
                System.Collections.Specialized.StringCollection cmbItems = Properties.Settings.Default.MetodoCommandDocs;



                List<itemConf> items = new List<itemConf>();
                 
                foreach (string f in cmbItems)
                {
                    string[] x = f.Split('|');

                    itemConf i = new itemConf();
                    i.Key = x[0];
                    i.Sql = x[1];
                    i.Oggi = x[2];



                    items.Add(i);


                }
                cmbInvioDocKnos.DataSource = items;
                GridViewDataRowInfo r = (GridViewDataRowInfo)cmbInvioDocKnos.SelectedItem;

                chkModificatiOggi.Checked = (r.Cells[2].Value.ToString() == "1");


            }
            catch (Exception ex)
            {
                lblPNGFirma.Text = "";
                //return;
            }

            richTextBoxMSG.Text = Properties.Settings.Default.sendMailComunicazioni;

            Application.DoEvents();

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

        private void button4_Click_1(object sender, EventArgs e)
        {

            

        }


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

            bool loadIMpostazioni = true;
            
            Cursor c = Cursors.WaitCursor;

            if (radGridView1.DataSource != null)
            {
                loadIMpostazioni = false;
            }



            string commandtext = Properties.Settings.Default.MetodoCommand;


            if (txtbArticolo.Text != "")
            {
                commandtext += string.Format("  WHERE EXISTS (SELECT 1 FROM ZS_VISTA_NOTIFICHEBSC v WHERE (v.CODART LIKE '%{0}%' OR v.DESCRIZIONEARTICOLO LIKE '%{0}%') AND v.CODCLIFOR = VISTA_NOTIFICHECLIENTICONTATTI.CODCONTO)", txtbArticolo.Text);


            }

            radGridView1.DataSource = null;
            radGridView1.AutoSizeColumnsMode = Telerik.WinControls.UI.GridViewAutoSizeColumnsMode.Fill;
            //commandtext += " WHERE DATADOC >= @DATADOC";

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {

                try
                {
                    cn.Open();
                    
                    toolStripStatusLabel1.Text = string.Format("caricamento dati in corso..........");

                    radGridView1.EnableFiltering = false;
                    radGridView1.ShowFilteringRow = false;

                    using (SqlCommand cmd = new SqlCommand(commandtext, cn))
                    {

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
                                //radGridView1.Columns[i].BestFit();
                                radGridView1.Columns[i].AutoSizeMode = Telerik.WinControls.UI.BestFitColumnMode.AllCells;
                            }
                        }


                        radGridView1.AutoScroll = true;
                        radGridView1.Refresh();

                        toolStripStatusLabel1.Text = string.Format("Caricamento completato");

                    }


                    radGridView1.EnableFiltering = true;
                    radGridView1.ShowFilteringRow = true;
                    radGridView1.EnableAlternatingRowColor = true;
                    radGridView1.MultiSelect = true;
                    
                    

                    radGridView1.Refresh();

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


            if (loadIMpostazioni == true)
            {
                LoadGridSettings(radGridView1);
            }

            txtSearch_Search.Text = strW;

            Cursor.Current = Cursors.Default;

            lblPNGFirma.Text = string.Format("Nr articoli trovati: {0}", radGridView1.Rows.Count.ToString());
        }



        private void btnSchedaPDL_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }

        private void btnKnoSLogin_Click(object sender, EventArgs e)
        {
            //kw.CurrentUser = txtKnoSUser.Text;
            //kw.PWD = txtKnoSPassword.Text;
            //if (kw.Inizializza(txtKnosUrl.Text) == true)
            //{
            //    btnKnoSLogin.BackColor = Color.LightGreen;
            //    Properties.Settings.Default.KnoS_URL = txtKnosUrl.Text;
            //    Properties.Settings.Default.KnoS_User = txtKnoSUser.Text;
            //    Properties.Settings.Default.KnoS_PWD = txtKnoSPassword.Text;
            //    Properties.Settings.Default.Save();

            //}
            
        }

        private void SaveGridSettings(DataGridView dg)
        {
            //salva le impostazioni della gridview in un file XML per utente
            string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings_{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

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


        private void LoadGridSettings(DataGridView dg)
        {
            //carica le impostazioni della gridview in un file XML per utente
            string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings_{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            DataTable dt = new DataTable();
            dt.ReadXmlSchema(pathGridSettings);

            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                dg.Columns[col.ColumnName].DisplayIndex = i;
                i++;
            }

        }


        private void SaveGridSettings(Telerik.WinControls.UI.RadGridView dg)
        {
            // salva le impostazioni della gridview in un file XML per utente
            string pathGridSettings = Path.Combine(Application.StartupPath, "GridLayout", string.Format("gridsettings_{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            if (!Directory.Exists(Path.Combine(Application.StartupPath, "GridLayout")))
            {
                Directory.CreateDirectory(Path.Combine(Application.StartupPath, "GridLayout"));
            }
            // salva le impostazioni della gridview in un file XML per utente
            //string pathGridSettings = Path.Combine(Application.StartupPath, string.Format("gridsettings{0}_{1}.xml", txtKnoSUser.Text, dg.Name));


            string sf = Path.Combine(pathGridSettings);


            dg.SaveLayout(sf);

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


        private void LoadGridSettings(Telerik.WinControls.UI.RadGridView dg)
        {
            //// salva le impostazioni della gridview in un file XML per utente
            string pathGridSettings = Path.Combine(Application.StartupPath, "GridLayout", string.Format("gridsettings_{0}_{1}.xml", txtKnoSUser.Text, dg.Name));

            try
            {
                dg.LoadLayout(pathGridSettings);
            }
            catch { }
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
            string ragionesociale = "";
            string codart = "";

            
            int IdDocBOL = 0;
            int IdDocSCH = 0;
            string localfileSCH = "";

            string msg = "";

            var attachments = new List<string>();
            int IdObjectSentMail = 0;
            List<string> IdObjectCLIList = new List<string>();

            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = radGridView1.SelectedRows.Count+1;
            toolStripProgressBar1.Value = 1;
            toolStripProgressBar1.Step = 1;
            toolStripProgressBar1.Visible = true;

            string addressBCCSimulazione = "";

            log.LogSomething(string.Format("Nr mail da inviare: {0}", radGridView1.SelectedRows.Count));

            string tempPathDownload = Path.Combine(Application.StartupPath, "TEMP");
            //if (!Directory.Exists(tempPathDownload))
            //{
            //    Directory.CreateDirectory(tempPathDownload);
            //}

            cleanTempFolder(tempPathDownload);



            checkBoxInterrompiInvio.Enabled = true;

            if ((textBoxSubject.Text == "") || (richTextBoxMSG.Text == ""))
            {
                MessageBox.Show("Inserire un titolo ed un testo del messaggio che si vuole inviare!", "Invio Mail Clienti", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabControl1.SelectedIndex = 2;
                textBoxSubject.Focus();
                return;
            }

            Application.DoEvents();

            if (allegati.Count > 0)
            {
                //dettaglioSCHAllegati = "<br /> ALLEGATI / ATTACHEMENTS:";

                foreach (Allegato a in allegati)
                {



                    //    dettaglioSCHAllegati += string.Format("<br /> - {0}", a.FileName);
                    localfileSCH = Path.Combine(tempPathDownload, a.FileName);
                    File.Copy(a.Path.Replace("file://", ""), localfileSCH, true);
                    attachments.Add(localfileSCH);

                }

            }

            body = richTextBoxMSG.Text; //string.Format(Properties.Settings.Default.sendMailSchede, dettaglioBOL, dettaglioSCH, dettaglioSCHAllegati);
            
            radGridView1.ShowRowHeaderColumn = true;

            if (radGridView1.SelectedRows.Count > 0)
            {
                msg = string.Format("Procedo con l'invio delle notifiche {0}", radGridView1.SelectedRows.Count);

                if (MessageBox.Show(msg, "Invio Mail Clienti", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    if (chkAgente.Checked == true)
                    {
                        // invio singolo con agente in CC
                        for (int i = 0; i < radGridView1.SelectedRows.Count; i++)
                        {
                            toolStripProgressBar1.Value += 1;
                            toolStripProgressBar1.Text = string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridView1.SelectedRows.Count);
                            log.LogSomething(string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridView1.SelectedRows.Count));



                            localfileSCH = "";
                            dettaglioBOL = dettaglioSCH = "";

                            ragionesociale = radGridView1.SelectedRows[i].Cells["DSCCONTO1"].Value.ToString();

                            if (chkSimulazione.Checked == false)
                            {
                                // destinatari reali
                                address = radGridView1.SelectedRows[i].Cells["EMAIL_CONTATTO"].Value.ToString();
                                addressCC = radGridView1.SelectedRows[i].Cells["EMAIL_AGENTI"].Value.ToString();

                                if (address.Contains(addressCC) && addressCC.Length > 0)
                                    address = address.Replace(addressCC, "");

                                if (!IdObjectCLIList.Contains(radGridView1.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()))
                                {
                                    IdObjectCLIList.Add(radGridView1.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString());
                                }

                            }

                            //addressBCC += radGridView1.SelectedRows[i].Cells["EMAIL_CONTATTO"].Value.ToString();

                            log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));

                            subject = string.Format(textBoxSubject.Text);
                            dettaglioSCH = richTextBoxMSG.Text;


                            // invio singolo
                            Application.DoEvents();

                            if (InviaMail(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                            {

                                // store della mail inviata
                                try
                                {
                                    toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                    IdObjectSentMail = 0;
                                    Application.DoEvents();
                                    IdObjectSentMail = kw.StoreEmailSent(3, "2", IdObjectCLIList, System.DateTime.Now, "comunicazioni", address, addressCC, addressBCC, subject, body, "", null, attachments);
                                    log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                    textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }


                            }
                            else
                            {
                                radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                return;
                            }

                            radGridView1.SelectedRows[i].Cells[1].Style.DrawFill = true;
                            radGridView1.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                            radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                            radGridView1.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                            radGridView1.SelectedRows[i].Cells[2].Style.DrawFill = true;
                            radGridView1.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                            radGridView1.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                            Application.DoEvents();
                            
                        }
                    }
                    else
                    {

                        
                        subject = string.Format(textBoxSubject.Text);
                        dettaglioSCH = richTextBoxMSG.Text;

                        // invio raggruppato solo al contatto in BCC
                        for (int i = 0; i < radGridView1.SelectedRows.Count; i++)
                        {


                            toolStripProgressBar1.Value += 1;
                            toolStripProgressBar1.Text = string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridView1.SelectedRows.Count);
                            log.LogSomething(string.Format("Record {0}/{1}", toolStripProgressBar1.Value, radGridView1.SelectedRows.Count));

                            

                            localfileSCH = "";
                            dettaglioBOL = dettaglioSCH = "";

                            ragionesociale = radGridView1.SelectedRows[i].Cells["DSCCONTO1"].Value.ToString();


                            if (!addressBCCSimulazione.Contains(radGridView1.SelectedRows[i].Cells["EMAIL_CONTATTO"].Value.ToString()))
                                addressBCCSimulazione += radGridView1.SelectedRows[i].Cells["EMAIL_CONTATTO"].Value.ToString();

                            if (addressBCCSimulazione.EndsWith(";"))
                            {

                            }
                            else
                            {
                                addressBCCSimulazione += ";";
                            }

                            if (chkSimulazione.Checked == false)
                            {
                                // destinatari reali
                                address = Properties.Settings.Default.sendMailBCC;
                                addressCC = "";
                                addressBCC+= radGridView1.SelectedRows[i].Cells["EMAIL_CONTATTO"].Value.ToString();

                                if (address.Contains(addressCC) && addressCC.Length > 0)
                                    address = address.Replace(addressCC, "");



                            }
                            else
                            {
                                address = Properties.Settings.Default.sendMailBCCSimulazione;
                                addressCC = "";
                                addressBCC = "";



                            }

                            log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));

                            

                            if (addressBCC.EndsWith(";"))
                            { 
                            
                            }
                            else
                            {
                                addressBCC += ";";
                            }


                            if (!IdObjectCLIList.Contains(radGridView1.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()))
                            {
                                IdObjectCLIList.Add(radGridView1.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString());
                            }


                            if (chkSimulazione.Checked == false)
                            {
                                if ((addressBCC.Length > 250) || (i == radGridView1.SelectedRows.Count - 1))
                                {
                                    log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));

                                    // invio singolo
                                    //Application.DoEvents();

                                    if (InviaMail(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                    {
                                        textBoxLOG.Text += string.Format("\r\n --- elenco indirizzi spedizione BCC {0}", addressBCC);
                                        textBoxLOG.Text += string.Format("\r\n --- elenco indirizzi spedizione BCC-Simulazione {0}", addressBCCSimulazione);

                                        // store della mail inviata
                                        try
                                        {
                                            toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                            IdObjectSentMail = 0;
                                            //Application.DoEvents();
                                            IdObjectSentMail = kw.StoreEmailSent(3, "2", IdObjectCLIList, System.DateTime.Now, "comunicazioni", address, addressCC, addressBCC, subject, body, "", null, attachments);
                                            log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                            textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.Message);
                                        }

                                        addressBCC = addressBCCSimulazione = "";
                                        IdObjectCLIList.Clear();
                                    }
                                    else
                                    {
                                        radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                                        addressBCC = addressBCCSimulazione = "";
                                        IdObjectCLIList.Clear();

                                        return;
                                    }
                                }
                            }
                            else
                            {
                                if ((addressBCCSimulazione.Length > 250) || (i == radGridView1.SelectedRows.Count - 1))
                                {
                                    if (InviaMail(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                                    {
                                        textBoxLOG.Text += string.Format("\r\n --- elenco indirizzi spedizione BCC {0}", addressBCC);
                                        textBoxLOG.Text += string.Format("\r\n --- elenco indirizzi spedizione BCC-Simulazione {0}", addressBCCSimulazione);
                                        
                                    }
                                    addressBCC = addressBCCSimulazione = "";
                                    IdObjectCLIList.Clear();
                                }

                            }

                            radGridView1.SelectedRows[i].Cells[1].Style.DrawFill = true;
                            radGridView1.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                            radGridView1.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                            radGridView1.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                            radGridView1.SelectedRows[i].Cells[2].Style.DrawFill = true;
                            radGridView1.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                            radGridView1.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                            //Application.DoEvents();

 
                    
                        }

                    }

                    toolStripProgressBar1.Visible = false;
                    checkBoxInterrompiInvio.Enabled = false;

                    MessageBox.Show("Invio completato!");
                }
            }
        }


        public bool InviaMail(string address, string subject, string body, List<string> attachments, List<string> bodyImages, bool bWeb, string addressCC, string addressBCC)
        { 
            if (checkBoxInterrompiInvio.Checked == true)
            {
                MessageBox.Show("Invio email interrotto dall'utente!", "Invio notifiche", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                checkBoxInterrompiInvio.Enabled = false;
                checkBoxInterrompiInvio.Checked = false;

                return false;
            }

            if (Properties.Settings.Default.UseCdo)
            {
                //subject = string.Format("{0}  {1}", subjectBOL, subjectSCH);
                if (Properties.Settings.Default.UseVbs)
                {

                    //body = string.Format(Properties.Settings.Default.sendMailComunicazioniVbsLotus);

                    ////if (Notifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                    //if (Notifica.SendNotifyVBSLotus(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                    //{
                    //    log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));

                    //    if (chkSimulazione.Checked == false)
                    //    {

                    //    }
                    //    else
                    //    {
                    //        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                    //        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                    //    }
                    //}

                    if (Notifica.SendNotifyCdo(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                    {
                        log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                        
                        textBoxLOG.Text += string.Format("\r\n - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC);

                    }
                    else
                    {
                        log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                        textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);

                    }
                }
            }
            else
            {
                if (Notifica.SendNotifyMAPI(address, subject, body, attachments, checkBoxPopUpMail.Checked, addressCC, addressBCC) == true)
                {
                    log.LogSomething(string.Format("Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                    textBoxLOG.Text += string.Format("\r\n - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC);

                }
                else
                {
                    log.LogSomething(string.Format("ERRORE - Invio mail riuscito {0} - {1} - {2} - {3} - {4} - {5}", address, subject, body, attachments.Count, addressCC, addressBCC));
                    textBoxLOG.Text += string.Format("\r\n ERRORE {0} {1} {2} {3} {4} {5} {6}", System.DateTime.Now.ToLongTimeString(), address, subject, body, attachments[0], addressCC, addressBCC);
                }
            }
        
            return true;
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


        private static void aggiornaRegistro()
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


        DataTable getArticoliSchede()
        {
            DataTable x = new DataTable();

            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand(string.Format("SELECT * FROM VISTA_RELAZIONIARTICOLIBSC")))
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
                checkBoxInterrompiInvio.Checked = true;
            }

            return bOK;

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
                "Indirizzo email: {1} | " + 
                "Indirizzo email agenti: {2} \r\n" +
                "Indirizzo email contatto: {3} | " + 
                "Ultimo invio: {4} \r\n";

            string cliente = string.Format("{0} - {1}", radGridView1.Rows[r].Cells["CODCONTO"].Value.ToString(), radGridView1.Rows[r].Cells["DSCCONTO1"].Value.ToString());
            string address = radGridView1.Rows[r].Cells["EMAIL_CLIENTE"].Value.ToString();
            string addressCC = radGridView1.Rows[r].Cells["EMAIL_AGENTI"].Value.ToString();
            string addressBCC = radGridView1.Rows[r].Cells["EMAIL_CONTATTO"].Value.ToString();
            string ultimoinvioSCH = "";
            
            //if (radGridView1.Rows[r].Cells["DATAULTIMOINVIO_SCH"].Value != null)
            //    ultimoinvioSCH = radGridView1.Rows[r].Cells["DATAULTIMOINVIO_SCH"].Value.ToString();

            textBoxToolTip.Text = string.Format(tooltip, cliente, address, addressCC, addressBCC, ultimoinvioSCH);
        }

        private void radGridView1_CurrentRowChanging(object sender, Telerik.WinControls.UI.CurrentRowChangingEventArgs e)
        {
            if (radGridView1.Rows.Count > 0)
            {
                if (e.NewRow.Index >= 0)
                    datiriga(e.NewRow.Index);
            }
        }






        private void btnAllegati_Click(object sender, EventArgs e)
        {
            
            frmUpload f = new frmUpload();
            
            f.ShowDialog();

            foreach (DataGridViewRow r in f.dataGridView1.Rows)
            {
                if (r.Cells[0].Value != null)
                {
                    Allegato a = new Allegato(r.Cells[0].Value.ToString(), r.Cells[0].Value.ToString(), r.Cells[2].Value.ToString());

                    allegati.Add(a);

                    dataGridView1.Rows.Add(r.Cells[0].Value.ToString(), r.Cells[2].Value.ToString());
                }

            }


        }

        private void MasterTemplate_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
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

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("Eliminare elenco allegati?", "Gestione allegati mail", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                allegati.Clear();
                dataGridView1.Rows.Clear();
            
            }
        }

        private void btnLoadDoc_Click(object sender, EventArgs e)
        {
            bool loadIMpostazioni = true;
            GridViewDataRowInfo r = (GridViewDataRowInfo)cmbInvioDocKnos.SelectedItem;

            Cursor c = Cursors.WaitCursor;

            if (rgvInvioDocMetodo.DataSource != null)
            {
                loadIMpostazioni = false;
            }


            string sql = r.Cells[1].Value.ToString();

            if (chkModificatiOggi.Checked == true)
            {
                sql = sql.Replace(" WHERE ", " WHERE OGGI = 0 AND ");
            }


            rgvInvioDocMetodo.DataSource = getDocumentiDaInviare(sql);

            if (loadIMpostazioni == true)
            {
                LoadGridSettings(rgvInvioDocMetodo);
            }

            c = Cursors.Default;

        }


        DataTable getDocumentiDaInviare(string sql)
        {
            DataTable x = new DataTable();

            try
            {
                using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.MetodoConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql))
                    {
                        cn.Open();
                        cmd.Connection = cn;
                        SqlDataAdapter da = new SqlDataAdapter(cmd);

                        da.Fill(x);

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            return x;
        }

        

        private void rgvInvioDocMetodo_CellClick(object sender, GridViewCellEventArgs e)
        {

            if (e.RowIndex >= 0)
            {

                string url = "{0}/KnoS_Catalog/0/{1}/{2}/{3}";


                url = string.Format(url, txtKnosUrl.Text, rgvInvioDocMetodo.SelectedRows[0].Cells["IDOBJECT_DOC"].Value.ToString(), rgvInvioDocMetodo.SelectedRows[0].Cells["IDDOC"].Value.ToString(), rgvInvioDocMetodo.SelectedRows[0].Cells["FILENAME"].Value.ToString().Substring(0, 15) + ".PDF");
                url = url.Replace("#", "_");

                webBrowser2.Navigate(url);

                //frmWebBrowser frm = new frmWebBrowser();
                //frm.Text = url;
                //frm.url = url;
                //frm.ShowDialog();
            }
        }

        private void SendMail(RadGridView rgv)
        {
            string address = Properties.Settings.Default.sendMailBCCSimulazione; //;kavanzi@italcom.biz";
            string addressCC = "";  //Properties.Settings.Default.sendMailBCCSimulazione; //"alfredo.deangelo@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string addressBCC = Properties.Settings.Default.sendMailBCC; // "knosmail@gmail.com;m.michieletti@zschimmer-schwarz.com";
            string body = "";
            string subject = "";
            string dettaglioDOC = "";

            bool bOKDownload = false;

            string codclifor = "";
            string ragionesociale = "";
            string filenamepattern = "";
            string documento = "";


            string localfile = "";

            string msg = "";

            int IDOBJECT = 0;

            var attachments = new List<string>();
            int IdObjectSentMail = 0;
            List<string> IdObjectCLIList = new List<string>();
            List<int> IdObjectDOCList = new List<int>();

            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = rgv.SelectedRows.Count + 1;
            toolStripProgressBar1.Value = 1;
            toolStripProgressBar1.Step = 1;
            toolStripProgressBar1.Visible = true;

            string addressBCCSimulazione = "";

            log.LogSomething(string.Format("Nr mail da inviare: {0}", rgv.SelectedRows.Count));

            string tempPathDownload = Path.Combine(Application.StartupPath, "TEMP");

            cleanTempFolder(tempPathDownload);

            checkBoxInterrompiInvio.Enabled = true;

            textBoxSubject.Text = Properties.Settings.Default.sendMailDocSubject;

            if ((textBoxSubject.Text == "") || (richTextBoxMSG.Text == ""))
            {
                MessageBox.Show("Inserire un titolo ed un testo del messaggio che si vuole inviare!", "Invio Mail Clienti", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tabControl1.SelectedIndex = 2;
                textBoxSubject.Focus();
                return;
            }

            Application.DoEvents();

            body = richTextBoxMSG.Text; //string.Format(Properties.Settings.Default.sendMailSchede, dettaglioBOL, dettaglioSCH, dettaglioSCHAllegati);

            rgv.ShowRowHeaderColumn = true;

            if (rgv.SelectedRows.Count > 0)
            {
                msg = string.Format("Procedo con l'invio delle notifiche {0}", rgv.SelectedRows.Count);

                if (MessageBox.Show(msg, "Invio Mail Clienti", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    // invio singolo con agente in CC
                    for (int i = 0; i < rgv.SelectedRows.Count; i++)
                    {
                        toolStripProgressBar1.Value += 1;
                        toolStripProgressBar1.Text = string.Format("Record {0}/{1}", toolStripProgressBar1.Value, rgv.SelectedRows.Count);
                        log.LogSomething(string.Format("Record {0}/{1}", toolStripProgressBar1.Value, rgv.SelectedRows.Count));

                        int.TryParse(rgv.SelectedRows[i].Cells["IDOBJECT_DOC"].Value.ToString(), out IDOBJECT);

                        filenamepattern = rgv.SelectedRows[i].Cells["FILENAME"].Value.ToString();
                        localfile = Path.Combine(tempPathDownload, filenamepattern);

                        if (kw.downloadDoc(IDOBJECT, filenamepattern, tempPathDownload, filenamepattern))
                        {
                            attachments.Add(Path.Combine(tempPathDownload, localfile));
                        }

                        dettaglioDOC = "";

                        ragionesociale = rgv.SelectedRows[i].Cells["RAGIONESOCIALE"].Value.ToString();
                        documento = rgv.SelectedRows[i].Cells["DOCUMENTO"].Value.ToString();



                        if (chkSimulazioneDOC.Checked == false)
                        {
                            // destinatari reali
                            address = rgv.SelectedRows[i].Cells["EMAIL_CONTATTO"].Value.ToString();

                            addressCC = rgv.SelectedRows[i].Cells["EMAIL_AGENTI"].Value.ToString();

                            if (address.Contains(addressCC) && addressCC.Length > 0)
                                address = address.Replace(addressCC, "");

                        }

                        if (!IdObjectCLIList.Contains(rgv.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString()))
                        {
                            IdObjectCLIList.Add(rgv.SelectedRows[i].Cells["IDOBJECT_CLI"].Value.ToString());
                        }

                        if (!IdObjectDOCList.Contains(int.Parse(rgv.SelectedRows[i].Cells["IDOBJECT_DOC"].Value.ToString())))
                        {
                            IdObjectDOCList.Add(int.Parse(rgv.SelectedRows[i].Cells["IDOBJECT_DOC"].Value.ToString()));
                        }

                        log.LogSomething(string.Format("Invio a : {0} - {1}", address, addressCC));

                        subject = string.Format(textBoxSubject.Text, documento, ragionesociale);
                        dettaglioDOC = richTextBoxMSG.Text;

                        // invio singolo
                        Application.DoEvents();

                        if (InviaMail(address, subject, body, attachments, null, true, addressCC, addressBCC) == true)
                        {

                            // store della mail inviata
                            try
                            {
                                toolStripStatusLabel1.Text = string.Format("Archiviazione mail in Knos");
                                IdObjectSentMail = 0;
                                Application.DoEvents();
                                IdObjectSentMail = kw.StoreEmailSent(3, "2", IdObjectCLIList, System.DateTime.Now, "comunicazioni", address, addressCC, addressBCC, subject, body, "", IdObjectDOCList, attachments);
                                log.LogSomething(string.Format("Archiviazione mail in Knos {0}", IdObjectSentMail));
                                textBoxLOG.Text += string.Format("\r\n --- Archiviazione mail in Knos {0}", IdObjectSentMail);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }


                        }
                        else
                        {
                            rgv.SelectedRows[i].Cells[1].Style.BackColor = Color.Red;
                            return;
                        }

                        rgv.SelectedRows[i].Cells[1].Style.DrawFill = true;
                        rgv.SelectedRows[i].Cells[1].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                        rgv.SelectedRows[i].Cells[1].Style.BackColor = Color.Lime;
                        rgv.SelectedRows[i].Cells[1].Style.CustomizeFill = true;
                        rgv.SelectedRows[i].Cells[2].Style.DrawFill = true;
                        rgv.SelectedRows[i].Cells[2].Style.GradientStyle = Telerik.WinControls.GradientStyles.Solid;
                        rgv.SelectedRows[i].Cells[2].Style.BackColor = Color.Lime;

                        Application.DoEvents();

                        IdObjectCLIList.Clear();
                        IdObjectDOCList.Clear();
                        attachments.Clear();

                    }

                    toolStripProgressBar1.Visible = false;
                    checkBoxInterrompiInvio.Enabled = false;

                    MessageBox.Show("Invio completato!");
                }
            }
        }

        private void btnSendDocMetodo_Click(object sender, EventArgs e)
        {
            SendMail(rgvInvioDocMetodo);
        }

        private void frmToDoNotifiche_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveGridSettings(dataGridView1);
            SaveGridSettings(radGridView1);
            SaveGridSettings(rgvInvioDocMetodo);
        }
    }


}

