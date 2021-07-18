using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;

using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;

using System.Runtime.InteropServices;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

//using Microsoft.Office;
using Microsoft.Office.Interop.Word;


using Knos;
using Knos.API.NET;
using Knos.API.COM;
using System.Net;

using System.Reflection;
using System.Diagnostics;

using KnosCSSignLibrary;

using Outlook = Microsoft.Office.Interop.Outlook;
using SendFileTo;

using System.Xml;


namespace SignRTFPDF
{

    public partial class frmFirma : Form
    {

        public static bool bSoloNotifica = false;

        public bool notifyPopUp = true;
        public static string CurrentResponsabileTecnico;
        public static string CurrentCapoCommessa;
        public static string CurrentTecnico;
        public static int CurrentIdObject;
        public static int CurrentIdObjectCertificato;
        public static int CurrentIdDocCertificato;
        public static int CurrentIDStatusPDL;
        public static string CurrentStatusNamePDL;
        public static int CurrentIdAction;
        public static string CurrentAttrNameData;
        public static string CurrentFileName;
        public static string CurrentFileDescr;
        public static string CurrentPDFPDLUrl;
        public static string strFilePDFPDL;




        public static int nrCertificati1F;
        public static int nrCertificati2F; 
        public static int nrCertificatiTot; 
        public static int nrCertificatiUtente; 
        public static int nrCertificatiUtente1F; 
        public static int nrCertificatiUtente2F; 
        public static int nrCertificatiUtente1FDaFirmare; 
        public static int nrCertificatiUtente2FDaFirmare;

        public static string CurrentPDLTitle = "";



        public class KnoSWrapper 
        {

            
            public string CurrentPWD = "";

            IKnosObject knosObject ;
            IKnosObject knosObjectCertificato;
            IKnosObjectMaker knosObjectMaker;
            IKnosObject knosObjectCliente;

            int cIdSubject = 0;
            string cUserName = "";

            public string DefaultSite;
            public string CurrentUser;


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
                            ikr.ClearAll();

                            //ikr = KnosInstance.Client.Login(CurrentUser, CurrentPWD, out cIdSubject);

                            if (cIdSubject > 0)
                            {
                                DefaultSite = _defaultSite;
                                CurrentUser = cUserName;
                                retvalue = true;
                            }
                            else
                            {
                                MessageBox.Show("Verificare le credenziali dell'utente");

                                retvalue = false;
                            }

                        }
                        else
                        {
                            //MessageBox.Show("Utente non loggato da Internet Explorer");
                            ikr.ClearAll();

                            ikr = KnosInstance.Client.Login(CurrentUser, CurrentPWD, out cIdSubject);

                            if (cIdSubject > 0)
                            {
                                DefaultSite = _defaultSite;
                                CurrentUser = cUserName;
                                retvalue = true;
                            }
                            else
                            {
                                MessageBox.Show("Verificare le credenziali dell'utente");

                                retvalue = false;
                            }

                            retvalue = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format("Impossibile aprire KnoS all'indirizzo {0}", _defaultSite));

                        retvalue = false;
                    }
                }
                
                
                return retvalue;
            
            
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


            public DataTable GetSostituti()
            {
                DataTable retTable = new DataTable();
                string pathFirma = "";

                retTable.Columns.Add("Utente");
                retTable.Columns.Add("Responsabile");
                retTable.Columns.Add("CapoCommessa");
                retTable.Columns.Add("PathFirma");


                IKnosObjectSelector ks = KnosInstance.Client.CreateKnosObjectSelector();
                ks.SearchExpression = "IdClass = 133";
                ks.SelectIdView = 127;
                ks.PageSize = 100;

                IKnosResult ikr = ks.GetPage(1);

                try
                {
                    if (ikr.HasErrors == false)
                    {
                        for (int i = 0; i < ks.ItemCount; i++)
                        {

                            int idObjectSost = ks.GetItem(i).IdObject;

                            //MessageBox.Show(ks.GetItem(i).IdObject.ToString(), "Id Utente");

                            ikr = ks.GetItem(i).GetObjectLinks(idObjectSost);

                            if ((ikr.HasErrors == false) && (ks.GetItem(i).LinkList.ItemCount > 0))
                            {
                                if (ks.GetItem(i).LinkList.GetItem(0).Url.StartsWith("file:"))
                                {
                                    pathFirma = ks.GetItem(i).LinkList.GetItem(0).Url.ToString().Replace(@"file://", "");
                                }
                                else
                                {
                                    pathFirma = ks.GetItem(i).LinkList.GetItem(0).Url.ToString();
                                }

                            }
                            else {
                                //MessageBox.Show(string.Format("Problemi con il reperimento dei dati dell'utente con IdObject {0}", ks.GetItem(i).IdObject.ToString()), "Id Utente");
                            }


                            retTable.Rows.Add(ks.GetItem(i).AttrValueList.GetItem(0).ToString(), ks.GetItem(i).AttrValueList.GetItem(1).ToString(), ks.GetItem(i).AttrValueList.GetItem(2).ToString(), pathFirma);
                        }
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Errore : {0}", ex.Message));
                }
                return retTable;
            }

            
            public bool GetPDL(int _idObject, ListView lvAttr, DataGridView dgCertificati, ListView lvFirme, StatusStrip s)
            { 
                bool retvalue = false;
                string fileName = "";
                string fileUrl = "";
                string fileLocalPath = "";
                int fileIdDoc = 0;
                string fileDescr = "";

                CurrentPDLTitle = "";

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
                    CurrentPDLTitle = knosObject.AttrValueList.GetItemByColumnName("Title").ToString();
                    
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
                                //UtenteResponsabileTecnico = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_52").ToString();
                                //DataSecondaFirma = knosObjectCertificato.AttrValueList.GetItemByColumnName("datetime_09").ToString();
                                UtenteCapoCommessa = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_53").ToString();
                                //UtenteResponsabileTecnicoSost = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_54").ToString();
                                UtenteCapoCommessaSost = knosObjectCertificato.AttrValueList.GetItemByColumnName("varchar_55").ToString();

                                CurrentCapoCommessa = UtenteCapoCommessa;

                                if (UtenteCapoCommessaSost != "")
                                {
                                    CurrentCapoCommessa = UtenteCapoCommessaSost;
                                }

                                //ikr = knosObjectViewList.GetItem(j).GetObjectDocuments();
                                ikr = knosObjectViewList.GetItem(j).GetObjectLinks();
                                if (ikr.HasErrors == false)
                                {

                                    fileName =  fileUrl = fileDescr = fileLocalPath = "";
                                    fileIdDoc = 0;
                                    if (knosObjectViewList.GetItem(j).LinkList.ItemCount == 1)
                                    {
                                        fileName = knosObjectViewList.GetItem(j).LinkList.GetItem(0).Url;
                                        fileUrl = knosObjectViewList.GetItem(j).LinkList.GetItem(0).Url;
                                        fileIdDoc = knosObjectViewList.GetItem(j).LinkList.GetItem(0).IdLink;
                                        fileDescr = knosObjectViewList.GetItem(j).LinkList.GetItem(0).LinkDescr;

                                        // pulizia dei file locali
                                        fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);
                                        if (File.Exists(fileLocalPath))
                                        {
                                            File.Delete(fileLocalPath);
                                        }
                                        ////download local del file
                                        //File.Delete(Path.Combine(Path.GetTempPath(), fileName));
                                        ////ikr = knosObjectViewList.GetItem(j).DocumentList.GetItem(0).DownloadFile(Path.GetTempPath(), fileName);
                                        ////if (ikr.HasErrors == false)
                                        ////{
                                        ////}
                                    }


                                    if (knosObjectViewList.GetItem(j).IdStatus == SignFiles.KnoS_Certificato_IdStatusIniziale)
                                    {
                                        //if (UtenteTecnico == CurrentUser)
                                        //{
                                            nrCertificatiUtente1FDaFirmare += 1;
                                        //}
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


                                    if (knosObjectViewList.GetItem(j).IdStatus == SignFiles.KnoS_Certificato_IdStatus2F)
                                    {
                                        nrCertificati2F += 1;

                                        if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                                        {
                                            nrCertificatiUtente2F += 1;
                                        }

                                    }



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


                            //if ((nrCertificati2F == nrCertificatiTot) && (CurrentCapoCommessa == CurrentUser))
                            if ((nrCertificatiUtente1FDaFirmare == 0) && (CurrentCapoCommessa == CurrentUser))
                            {
                                SignFiles.tipofirma = 2;
                            }

                            // reupero PDF certificato
                            knosObject.GetObjectDocuments();

                            if (knosObject.DocumentList.ItemCount > 0)
                            {
                                for (int iDocumento = 0; iDocumento < knosObject.DocumentList.ItemCount; iDocumento++)
                                {
                                    if (knosObject.DocumentList.GetItem(iDocumento).FileName.StartsWith(strFilePDFPDL))
                                    {
                                        //knosObject.DocumentList.GetItem(iDocumento).CurrentVersion 
                                        CurrentPDFPDLUrl = knosObject.DocumentList.GetItem(iDocumento).GetUrl();
                                    }

                                }
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

                CurrentPDLTitle = "";

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


                        CurrentPDLTitle = knosObject.AttrValueList.GetItemByColumnName("Title").ToString();


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
                                UtenteCapoCommessa = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_53").ToString();
                                try
                                {
                                    UtenteResponsabileTecnicoSost = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_54").ToString();
                                }
                                catch (Exception ex)
                                {
                                    UtenteResponsabileTecnicoSost = "";
                                }
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


                                //ikr = knosObjectSelectorCertificati.GetItem(i).GetObjectDocuments();
                                ikr = knosObjectSelectorCertificati.GetItem(i).GetObjectLinks();
                                if (ikr.HasErrors == false)
                                {

                                    fileName = fileUrl = fileDescr = fileLocalPath = "";
                                    fileIdDoc = 0;
                                    if (knosObjectSelectorCertificati.GetItem(i).LinkList.ItemCount == 1)
                                    {
                                        fileName = knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(0).LinkDescr;
                                        fileUrl = knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(0).Url;
                                        fileIdDoc = knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(0).IdLink;
                                        fileDescr = knosObjectSelectorCertificati.GetItem(i).LinkList.GetItem(0).LinkDescr;

                                        // pulizia dei file locali
                                        fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);
                                        if (File.Exists(fileLocalPath))
                                        {
                                            File.Delete(fileLocalPath);
                                        }
                                    }


                                    if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatusIniziale)
                                    {
                                        //if (UtenteTecnico == CurrentUser)
                                        //{
                                            nrCertificatiUtente1FDaFirmare += 1;
                                        //}
                                    }

                                    //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                                    //{
                                    //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                                    //    {
                                    //        nrCertificatiUtente2FDaFirmare += 1;
                                    //    }

                                    //}

                                    //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                                    //{
                                    //    nrCertificati1F += 1;

                                    //    if (UtenteTecnico == CurrentUser)
                                    //    {
                                    //        nrCertificatiUtente1F += 1;
                                    //    }
                                    //}


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

                                //if ((nrCertificati2F == nrCertificatiTot) && (CurrentCapoCommessa == CurrentUser))
                                if ((nrCertificatiUtente1FDaFirmare == 0) && (CurrentCapoCommessa == CurrentUser))
                                {
                                    SignFiles.tipofirma = 2;
                                }



                            }

                        // reupero PDF certificato
                        knosObject.GetObjectDocuments();

                        if (knosObject.DocumentList.ItemCount > 0)
                        {
                            for (int iDocumento = 0; iDocumento < knosObject.DocumentList.ItemCount; iDocumento++)
                            {
                                if (knosObject.DocumentList.GetItem(iDocumento).FileName.StartsWith(strFilePDFPDL))
                                {
                                    //knosObject.DocumentList.GetItem(iDocumento).CurrentVersion 
                                    CurrentPDFPDLUrl = knosObject.DocumentList.GetItem(iDocumento).GetUrl();
                                }
                                
                            }
                        }

                        

                        retvalue = true;
                    }


                    else
                    {
                        if (_idObject == 0)
                        {
                            MessageBox.Show("Non è stato scelto un PDL", "Apertura PDL", MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                        }
                        else
                        {
                            MessageBox.Show(ikr.GetError(1).Description);
                        }
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
                        UtenteResponsabileTecnico = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_52").ToString();
                        try
                        {
                            DataSecondaFirma = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("datetime_09").ToString();
                        }
                        catch (Exception ex)
                        {
                            DataSecondaFirma = "";
                        }
                        UtenteCapoCommessa = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_53").ToString();
                        try
                        {
                            UtenteResponsabileTecnicoSost = knosObjectSelectorCertificati.GetItem(i).AttrValueList.GetItemByColumnName("varchar_54").ToString();
                        }
                        catch (Exception ex)
                        {
                            UtenteResponsabileTecnicoSost = "";
                        }
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
                                //fileLocalPath = Path.Combine(Path.GetTempPath(), fileName);
                                //File.Delete(Path.Combine(fileLocalPath));
                            }


                            if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatusIniziale)
                            {
                                //if (UtenteTecnico == CurrentUser)
                                //{
                                    nrCertificatiUtente1FDaFirmare += 1;
                                //}
                            }

                            //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                            //{
                            //    if ((UtenteResponsabileTecnico == CurrentUser) || (UtenteResponsabileTecnicoSost == CurrentUser))
                            //    {
                            //        nrCertificatiUtente2FDaFirmare += 1;
                            //    }

                            //}

                            //if (knosObjectSelectorCertificati.GetItem(i).IdStatus == SignFiles.KnoS_Certificato_IdStatus1F)
                            //{
                            //    nrCertificati1F += 1;

                            //    if (UtenteTecnico == CurrentUser)
                            //    {
                            //        nrCertificatiUtente1F += 1;
                            //    }
                            //}


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
            

            public bool UploadFileCertificato(int _idObject, 
                int _idDoc, 
                string _filePath, 
                string _fileDescr, 
                string _fileName, 
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
                    ui.FilePath = _filePath;
                    ui.UploadType = EnumKnosUploadType.BarcodeUpload;

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
                        retvalue = false;

                        MessageBox.Show(knosResult.GetError(0).ToString());
                        return retvalue;
                    }

                    // aggiornamento attributo cambio stato 
                    if (_attrNameDate != "")
                    {
                        knosObjectMaker.SetAttrValue(_attrNameDate, System.DateTime.Now, EnumKnosDataType.DateTimeType);

                        knosResult = knosObjectMaker.UpdateObject(_idObject);
                    }

                    if (knosResult.HasErrors)
                    {
                        c = Cursors.Default;

                        MessageBox.Show(knosResult.GetError(0).ToString());
                        retvalue = false;
                    }
                    else
                    {
                        retvalue = true;
                    }
                }


                c = Cursors.Default;
                return retvalue;


            }



            public bool ActionCertificato(int _idObject,
                int _actionWF,
                string _attrNameDate
                )
            {
                bool retvalue = false;

                // upload file
                IKnosObjectMaker knosObjectMaker = KnosInstance.Client.CreateKnosObjectMaker();

                Cursor.Current = Cursors.WaitCursor;
                
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
                    if (_attrNameDate != "")
                    {
                        knosObjectMaker.SetAttrValue(_attrNameDate, System.DateTime.Now, EnumKnosDataType.DateTimeType);

                        knosResult = knosObjectMaker.UpdateObject(_idObject);
                    }

                    if (knosResult.HasWarningsErrors)
                    {
                        Cursor.Current = Cursors.Default;

                        MessageBox.Show(knosResult.ToString());
                    }
                    else
                    {
                        retvalue = true;
                    }
                }


                Cursor.Current = Cursors.Default;
                return retvalue;


            }


            public int CreateSurvey()
            {
                int IdObjectSurvey = 0;
                int IdClassSurvey = 0;

                try
                {
                    IdClassSurvey = Properties.Settings.Default.KnoS_IdClassSurvey;
                    
                    IKnosObjectMaker kom = KnosInstance.Client.CreateKnosObjectMaker();
                    kom.IdClass = IdClassSurvey;

                    kom.SetAttrValue("Title", "Survey del Piano di Lavoro " + CurrentPDLTitle);
                    IKnosResult kr = kom.CreateObject(out IdObjectSurvey);

                    if (kr.NoWarningsErrors)
                    {
                        // lo linko alla pubblicazione PDL
                        if (CurrentIdObject > 0)
                        {
                            kom.Reset();

                            IKnosLink kl = KnosInstance.Client.CreateKnosLink();
                            kl.IdObjectTo = IdObjectSurvey;
                            kom.LinkEditor.AddValue(kl);
                            kr = kom.UpdateObject(CurrentIdObject);

                            if (kr.NoWarningsErrors)
                            {

                            }
                            else
                            {
                                MessageBox.Show("Errore nel collegamento della pubblicazione Survey al PDL");
                            }
                        
                        }

                    }
                    else
                    {
                        MessageBox.Show("Errore nella creazione della pubblicazione Survey da collegare al PDL");
                    }


                }
                catch (Exception ex)
                { }


                return IdObjectSurvey;

            
            
            
            
            
            }






            public bool downloadDoc(int _idCertificato, int _idDoc = 1, string filePath = "")
            {
                IKnosResult ikr;
                bool bOK = false;

                ikr = knosObjectCertificato.GetObjectDocuments(_idCertificato);
                if (ikr.HasErrors == false)
                {
                    //download local del file
                    //File.Delete(filePath);
                    ikr = knosObjectCertificato.DocumentList.GetItem(0).DownloadFile(Path.GetTempPath(), knosObjectCertificato.DocumentList.GetItem(0).FileName);
                    
                }
                    
                if (ikr.HasErrors == false)
                {
                    return true;
                }
                else
                {
                    MessageBox.Show(ikr.GetError(0).ToString(), "Errore in Download allegato");
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



        public class TipoFirma
        {
            int id;

            public int Id
            {
                get { return id; }
                set { id = value; }
            }

            string description;

            public string Description
            {
                get { return description; }
                set { description = value; }
            }

            public TipoFirma(int id, string description) { this.id = id; this.description = description; }
        }

        /*
     * see: TextRenderInfo & RenderListener classes here:
     * http://api.itextpdf.com/itext/
     * 
     * and Google "itextsharp extract 
         * "
     */
        public class MyImageRenderListener : IRenderListener
        {
            public void RenderText(TextRenderInfo renderInfo) { }
            public void BeginTextBlock() { }
            public void EndTextBlock() { }

            public List<byte[]> Images = new List<byte[]>();
            public List<string> ImageNames = new List<string>();
            public void RenderImage(ImageRenderInfo renderInfo)
            {
                PdfImageObject image = renderInfo.GetImage();
                try
                {
                    image = renderInfo.GetImage();
                    if (image == null) return;

                    ImageNames.Add(string.Format(
                      "Image{0}.{1}", renderInfo.GetRef().Number, image.GetFileType()
                    ));
                    using (MemoryStream ms = new MemoryStream(image.GetImageAsBytes()))
                    {
                        Images.Add(ms.ToArray());
                    }
                }
                catch (IOException ie)
                {
                    /*
                     * pass-through; image type not supported by iText[Sharp]; e.g. jbig2
                    */
                }
            }
        }


        
        
 
        public frmFirma()
        {
            InitializeComponent();
        }


        private void btnSignImage_Click(object sender, EventArgs e)
        {
            //openFileDialog1.ShowDialog();
            //if (openFileDialog1.FileName != null)
            //{
            //    lblPNGFirma.Text = openFileDialog1.FileName;

            //    if (File.Exists(lblPNGFirma.Text))
            //    {
            //        System.Drawing.Image x = System.Drawing.Image.FromFile(lblPNGFirma.Text);

            //        //pictureBoxFirma.Image = x;

            //        pictureBoxFirma.Image = FixedSize(x, pictureBoxFirma.Width, pictureBoxFirma.Height);
            //        pictureBoxFirma.Visible = true;
                    
            //        txtStringSignIMG.Text = SignRTFPDF.SignFiles.GetStringFromPNG(lblPNGFirma.Text, 100, 100);
            //    }
                
            //}
        }

        static System.Drawing.Image FixedSize(System.Drawing.Image imgPhoto, int Width, int Height)
        {
            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = ((float)Width / (float)sourceWidth);
            nPercentH = ((float)Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
                destX = System.Convert.ToInt16((Width -
                              (sourceWidth * nPercent)) / 2);
            }
            else
            {
                nPercent = nPercentW;
                destY = System.Convert.ToInt16((Height -
                              (sourceHeight * nPercent)) / 2);
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            Bitmap bmPhoto = new Bitmap(Width, Height,
                              PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                             imgPhoto.VerticalResolution);

            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            grPhoto.Clear(Color.LightGray);
            grPhoto.InterpolationMode =
                    System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            grPhoto.DrawImage(imgPhoto,
                new System.Drawing.Rectangle(destX, destY, destWidth, destHeight),
                new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                GraphicsUnit.Pixel);

            grPhoto.Dispose();
            return bmPhoto;
        }



        private void  frmFirma_Load(object sender, EventArgs e)
        {
            // certificati
            cboTipoDispositivo.SelectedIndex = 1;

            this.Text = string.Format("TEC EUROLAB - Firma Certificati PDL ({0})", System.Windows.Forms.Application.ProductVersion);


            try
            {

                opened = false;

                int.TryParse(Properties.Settings.Default.KnoS_IdStatoIniziale, out SignFiles.KnoS_Certificato_IdStatusIniziale);
                int.TryParse(Properties.Settings.Default.KnoS_IdStato1F, out SignFiles.KnoS_Certificato_IdStatus1F);
                int.TryParse(Properties.Settings.Default.KnoS_IdStato2F, out SignFiles.KnoS_Certificato_IdStatus2F);
                int.TryParse(Properties.Settings.Default.KnoS_IdStatoPDLFirmato, out SignFiles.KnoS_PDL_IdStatusPDLFirmato);
                int.TryParse(Properties.Settings.Default.KnoS_IdStatoPDLDaFirmare, out SignFiles.KnoS_PDL_IdStatusPDLDaFirmare);

                int.TryParse(Properties.Settings.Default.KnoS_IdActionPDLDaFirmare, out SignFiles.KnoS_IdActionPDFdaFirmare);
                int.TryParse(Properties.Settings.Default.KnoS_IdActionPDFFirmato, out SignFiles.KnoS_IdActionPDFFirmato);

                int.TryParse(Properties.Settings.Default.KnoS_IdActionPDLFirmatoCERT, out SignFiles.KnoS_IdActionPDLFirmatoCERT);

                bool.TryParse(Properties.Settings.Default.sendMailPopUp, out notifyPopUp);


                System.Windows.Forms.Application.DoEvents();
                statusStrip1.Text = string.Format("File inizializzazione: {0}", SignFiles.startXML);

                if ((SignFiles.startXML_idobject == 0))
                {
                    if (SignFiles.startXML != "")
                    {
                        //                        MessageBox.Show(SignFiles.startXML);
                        SignFiles.LoadInfoXML(SignFiles.startXML);
                    }
                    else
                    {
                        // caricamento delle impostazioni provenienti dal file aperto
                        if (File.Exists(SignFiles.startXML) == false)
                        {
                            //MessageBox.Show(string.Format("File di inizializzazione {0} non trovato o non accessibile!", SignFiles.startXML), "Inizializzazione programma firma", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    
                    }
                }

                if (kw.Inizializza(SignFiles.startXML_baseurl) == true)
                {

                    txtKnosUrl.Text = kw.DefaultSite;
                    txtKnoSUser.Text = kw.CurrentUser;

                    statusStrip1.Text = string.Format("Caricamento utenti sostituti.....");
                    System.Windows.Forms.Application.DoEvents();

                    loadSostituti();                    
                    
                    txtIdPDL.Text = SignFiles.startXML_idobject.ToString();

                    int.TryParse(SignFiles.startXML_idobject_certificato.ToString(), out CurrentIdObjectCertificato);

                    //if (SignFiles.startXML_idobject_certificato > 0)
                    //{
                    //    statusStrip1.Text = string.Format("Caricamento Certificato PDL in corso.....");
                    //    System.Windows.Forms.Application.DoEvents();

                    //    LoadCertificatoPDL();
                    //    dataGridViewCertificati.Top = label8.Top;
                    //    dataGridViewCertificati.BringToFront();
                    //}
                    //else
                    {

                        statusStrip1.Text = string.Format("Caricamento PDL in corso.....");
                        System.Windows.Forms.Application.DoEvents();

                        LoadPDL();
                    }

                    txtKnosUrl.ReadOnly = txtKnoSUser.ReadOnly = txtKnoSPassword.ReadOnly = true;
                    btnKnoSLogin.Enabled = false;

                    statusStrip1.Text = string.Format("");
                    System.Windows.Forms.Application.DoEvents();
                }
                else
                {
                    txtKnosUrl.Text = SignFiles.startXML_baseurl;
                    MessageBox.Show(string.Format("Sito KnoS non trovato o non accessibile!", SignFiles.startXML_baseurl), "Inizializzazione programam firma", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                

                statusStrip1.Text = string.Format("PDL Caricato");
                System.Windows.Forms.Application.DoEvents();

                if ((bSoloNotifica) || (SignFiles.bNotificaCapocommessa))
                {
                    NotificaCapocommessa();

                    this.Close();
                }


                if (SignFiles.testSignPDF == false)
                {
                    //splitContainer1.Panel2.Width = 6;

                }            
            }
            catch (Exception ex)
            {
                //lblPNGFirma.Text = "file firma NON indicato";
                //return;
            }


            if (opened == false)
            {
                //    winWordControl1.HideCommandBars();
                //    LoadDoc();

                // mi posiziono sul certificato scelto eventualmente passato nel fiel di configurazione
                for (int j = 0; j < dataGridViewCertificati.Rows.Count; j++)
                {
                    if (int.Parse(dataGridViewCertificati.Rows[j].Cells["IdObject"].Value.ToString()) == SignFiles.startXML_idobject_certificato)
                    {
                        LoadPDFCertificato(j);
                        dataGridViewCertificati.Rows[j].Selected = true;
                        //tabControl1.SelectedIndex = 1;
                        tabControl1.SelectedIndex = 0;
                        break;
                    }
                    else
                    {
                        tabControl1.SelectedIndex = 1;
                    
                    }
                }

            }

            statusStrip1.Text = string.Format("PDL Caricato");
            System.Windows.Forms.Application.DoEvents();

        }
       
        private void exRichTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            char currentKey = (char)e.KeyCode;
            bool modifier = e.Control || e.Alt || e.Shift;

            bool nonvalid = char.IsNumber(currentKey) || char.IsLetter(currentKey) || char.IsSymbol(currentKey) || char.IsWhiteSpace(currentKey) || char.IsPunctuation(currentKey);
            if (nonvalid)
                e.SuppressKeyPress = true;
            return;
        }

 
        public bool CheckImagesByPixel(string fname1, string fname2)
        {
            int count1 = 0;
            int count2 = 0;
            Bitmap img1;
            Bitmap img2;

            bool flag = true;

            string img1_ref, img2_ref;
            img1 = new Bitmap(fname1);
            img2 = new Bitmap(fname2);

            if (img1.Width == img2.Width && img1.Height == img2.Height)
            {
                for (int i = 0; i < img1.Width; i++)
                {
                    for (int j = 0; j < img1.Height; j++)
                    {
                        img1_ref = img1.GetPixel(i, j).ToString();
                        img2_ref = img2.GetPixel(i, j).ToString();

                        if (img1_ref != img2_ref)
                        {
                            count2++;
                            flag = false;
                            break;
                        }
                        count1++;
                    }
                }
                if (flag == false)
                {
                    //MessageBox.Show("Sorry, Images are not same , wrong pixels found " + count2 );
                }
                else
                {
                    //MessageBox.Show(" Images are same , " + count1 + " same pixels found and " + count2 + " wrong pixels found");
                }
            }
            else
            {
                flag = false;
                //MessageBox.Show("can not compare this images");
            }
            img1.Dispose();
            img2.Dispose();
            img1 = img2 = null;
            
            return flag;
        }

        public object listener { get; set; }

        

        public static void CombineMultiplePDFs(string[] fileNames, string outFile)
        {

            // step 1: creation of a document-object
            iTextSharp.text.Document document = new iTextSharp.text.Document();

            // step 2: we create a writer that listens to the document
            PdfCopy writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
            if (writer == null)
            {
                return;
            }

            // step 3: we open the document
            document.Open();

            foreach (string fileName in fileNames)
            {
                if (File.Exists(fileName))
                {

                    // we create a reader for a certain document
                    PdfReader reader = new PdfReader(fileName);
                    reader.ConsolidateNamedDestinations();

                    // step 4: we add content
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        writer.AddPage(page);
                    }

                    PRAcroForm form = reader.AcroForm;
                    if (form != null)
                    {
                        writer.CopyAcroForm(reader);
                    }

                    reader.Close();
                }
            }

            // step 5: we close the document and writer
            writer.Close();
            document.Close();


            //int pageOffset = 0;
            ////int f = 0;
            //iTextSharp.text.Document document = null;
            //PdfCopy writer = null;
            //PdfReader reader = null;

            ////MessageBox.Show(String.Format("Nr certificati da unire {0}......", fileNames.Length));
            //try
            //{
            //    //File.Create(Path.GetRandomFileName());


        
            //    for (int f = 0; f < fileNames.Length; f++ )
            //    {
            //        //if (f == 0)
            //        //{

            //        //}

            //        if (fileNames[f] != "")
            //        {
            //            MessageBox.Show(string.Format(" certificato {0} ",fileNames[f]));

            //            // we create a reader for a certain document
            //            reader = new PdfReader(fileNames[f]);
            //            reader.ConsolidateNamedDestinations();
            //            // we retrieve the total number of pages
            //            int n = reader.NumberOfPages;
            //            pageOffset += n;

            //            if (document == null)
            //            {
            //                // step 1: creation of a document-object
            //                document = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));
            //                // step 2: we create a writer that listens to the document
            //                writer = new PdfCopy(document, new FileStream(outFile, FileMode.Create));
            //                // step 3: we open the document
            //                document.Open();
            //            }


            //            // step 4: we add content
            //            for (int i = 0; i < n; )
            //            {
            //                ++i;
            //                if (writer != null)
            //                {
            //                    PdfImportedPage page = writer.GetImportedPage(reader, i);
            //                    writer.AddPage(page);
            //                }
            //            }
            //            PRAcroForm form = reader.AcroForm;
            //            if (form != null && writer != null)
            //            {
            //                writer.CopyAcroForm(reader);
            //            }

            //            if (reader != null)
            //            {
            //                reader.Close();
            //                reader.Dispose();
            //            }

            //        }

            //        f++;
            //    }

            //    // step 5: we close the document
            //    if (document != null)
            //    {
            //        document.Close();
            //    }

            //    writer.Dispose();

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message); 
            //}
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            string[] filesToMerge = {SignFiles.tempOriginalPDF, SignFiles.tempSignedPDF};

            CombineMultiplePDFs(filesToMerge, SignFiles.tempMergePDF);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                //axAcroPDF1.LoadFile(SignRTFPDF.SignFiles.tempOriginalPDF);
                //Uri u = new Uri(SignFiles.tempOriginalPDF);
                //webBrowser1.Url = u;
                //webBrowser1.Refresh();
                //webBrowser1.Navigate(SignFiles.tempOriginalPDF);
            }


            if (tabControl1.SelectedIndex == 1)
            {
                GetAzioneCertificati();
            }
        }

        private void txtIdPDL_Leave(object sender, EventArgs e)
        {
            
        }

        private void LoadPDL()
        {

            btnFirmaCapoCommessa.Enabled = false;
            int _intRes = 0;

            if (int.TryParse(txtIdPDL.Text, out _intRes) == true)
            {
                
                Cursor.Current = Cursors.WaitCursor;

                CurrentIdObject = _intRes;
                toolStripStatusLabel1.Text = "Caricamento dati e certificati del PDL.....";

                kw.GetPDLSelector(_intRes, listViewAttr, dataGridViewCertificati, lvFileFirma, statusStrip1);

                if (dataGridViewCertificati.Rows.Count > 0)
                {
                    foreach (DataGridViewColumn dvc in dataGridViewCertificati.Columns)
                    {
                        dvc.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    dataGridViewCertificati.Sort(dataGridViewCertificati.Columns["IdObject"], ListSortDirection.Ascending);
                }
                // stato PDL
                btnPDLStatus.Text = CurrentStatusNamePDL;
                btnPDLStatus.Tag = CurrentPDFPDLUrl;

                //if (SignFiles.tipofirma > 0)
                //{
                    tabControl1.SelectedIndex = 1;
                //}


                // controllo stato certificati
                bool bFirmabileCapocommessa = true;
                string CapoCommessa = "";
                string CapoCommessaSost = "";



                if (dataGridViewCertificati.Rows.Count > 0)
                {
                    for( int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
                    {
                        if ((int.Parse(dataGridViewCertificati.Rows[i].Cells["IdStatus"].Value.ToString()) != SignFiles.KnoS_Certificato_IdStatus1F))
                        {
                            bFirmabileCapocommessa = false;
                        }
                        else
                        { 
                            //        CurrentCapoCommessa = dataGridViewCertificati["CapoCommessa", _RowIndex].Value.ToString();
                            //        if (dataGridViewCertificati["CapoCommessaSost", _RowIndex].Value.ToString() != "")
                            //        {
                            //            CurrentCapoCommessa = dataGridViewCertificati["CapoCommessaSost", _RowIndex].Value.ToString();
                            //        }
                            if (CapoCommessa == "")
                            {
                                CapoCommessa = dataGridViewCertificati["CapoCommessa", i].Value.ToString();
                            }
                            if (CapoCommessaSost == "")
                            {
                                CapoCommessaSost = dataGridViewCertificati["CapoCommessaSost", i].Value.ToString();
                            }
                        }
                    }
                }

                //btnFirmaCapoCommessa.Enabled = bFirmabileCapocommessa;
                toolStripStatusLabel1.Text = "Controllo capocommessa " + CapoCommessa;

                if (bFirmabileCapocommessa == true)
                {
                    btnFirmaCapoCommessa.Enabled = ((kw.CurrentUser == CapoCommessa) || (kw.CurrentUser == CapoCommessaSost));
                }

                toolStripStatusLabel1.Text = "";

            }
        }


        private void LoadCertificatoPDL()
        {
            int _intRes = 0;

            Cursor.Current = Cursors.WaitCursor;

            CurrentIdObject = _intRes;
            toolStripStatusLabel1.Text = "Caricamento dati e certificati del PDL.....";
            kw.GetCertificatoPDLSelector(_intRes, listViewAttr, dataGridViewCertificati, lvFileFirma, statusStrip1);

            // bloccaggio ordinamento colonne
            if (dataGridViewCertificati.Rows.Count > 0)
            {
                foreach (DataGridViewColumn dvc in dataGridViewCertificati.Columns)
                {
                    dvc.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                dataGridViewCertificati.Sort(dataGridViewCertificati.Columns["IdObject"], ListSortDirection.Ascending);
            }

            // stato PDL
            btnPDLStatus.Text = CurrentStatusNamePDL;
            btnPDLStatus.Tag = CurrentPDFPDLUrl;

            btnFirmaCapoCommessa.Enabled = (kw.CurrentUser == CurrentCapoCommessa);
            
            //GetAzioneCertificati();

            //if (SignFiles.tipofirma > 0)
            //{
            //    tabControl1.SelectedIndex = 1;
            //}

        }

        private void dataGridViewCertificati_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.ColumnIndex == 0) && (e.RowIndex > -1))
            {
                // premuto il bottone "firma"
                //LoadPDFCertificato(e.RowIndex);
                tabControl1.SelectedIndex = 0;



            }
        }

        //private void LoadWordCertificato(int _RowIndex)
        //{
        //    // resetto il backcolor della grid
        //    for (int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
        //    {
        //        dataGridViewCertificati.Rows[i].DefaultCellStyle.BackColor = Color.White;
        //    }
        //    dataGridViewCertificati.Rows[_RowIndex].DefaultCellStyle.BackColor = Color.Lime;


        //    try
        //    {
        //        string filePath = dataGridViewCertificati.Rows[_RowIndex].Cells["LocalFile"].Value.ToString();

        //        int.TryParse(dataGridViewCertificati.Rows[_RowIndex].Cells["IdObject"].Value.ToString(), out CurrentIdObjectCertificato);

        //        //CurrentResponsabileTecnico = dataGridViewCertificati["ResponsabileTecnico", _RowIndex].Value.ToString();
        //        //if (dataGridViewCertificati["ResponsabileTecnicoSost", _RowIndex].Value.ToString() != "")
        //        //{
        //        //    CurrentResponsabileTecnico = dataGridViewCertificati["ResponsabileTecnicoSost", _RowIndex].Value.ToString();
        //        //}

        //        CurrentCapoCommessa = dataGridViewCertificati["CapoCommessa", _RowIndex].Value.ToString();
        //        if (dataGridViewCertificati["CapoCommessaSost", _RowIndex].Value.ToString() != "")
        //        {
        //            CurrentCapoCommessa = dataGridViewCertificati["CapoCommessaSost", _RowIndex].Value.ToString();
        //        }


        //        foreach (ListViewItem li in lvCCSost.Items)
        //        {
        //            li.BackColor = Color.White;

        //            //if (li.Text == dataGridViewCertificati["CapoCommessa", _RowIndex].Value.ToString())
        //            if (li.Text == CurrentCapoCommessa)
        //            {
        //                li.BackColor = Color.Lime;
        //            }

        //        }


        //        if (File.Exists(dataGridViewCertificati["LocalFile", _RowIndex].Value.ToString()))
        //        {

        //        }
        //        else
        //        {
        //            if (kw.downloadDoc(CurrentIdObjectCertificato, 1, filePath))
        //            {

        //            }
        //            else
        //            {

        //            }
        //        }


        //        if (File.Exists(dataGridViewCertificati["Url", _RowIndex].Value.ToString()))
        //        {

        //            //SignFiles.tempOriginalPDF = dataGridViewCertificati["LocalFile", _RowIndex].Value.ToString();

        //            toolStripStatusLabel1.Text = "Caricamento file Word.....";

        //            CurrentIdObjectCertificato = int.Parse(dataGridViewCertificati["IdObject", _RowIndex].Value.ToString());

        //            OpenMicrosoftWord(dataGridViewCertificati["Url", _RowIndex].Value.ToString());




        //            //if ((dataGridViewCertificati["Tecnico", _RowIndex].Value.ToString() == kw.CurrentUser) && (dataGridViewCertificati["IdStatus", _RowIndex].Value.ToString() == Properties.Settings.Default.KnoS_IdStatoIniziale))
        //            //{
        //            //    int.TryParse(Properties.Settings.Default.KnoS_IdAction1F, out CurrentIdAction);
        //            //    CurrentAttrNameData = Properties.Settings.Default.KnoS_AttrNameData1F;
        //            //    SignFiles.tipofirma = 0;
        //            //}

        //            //if (((dataGridViewCertificati["ResponsabileTecnicoSost", _RowIndex].Value.ToString() == kw.CurrentUser) || (dataGridViewCertificati["ResponsabileTecnico", _RowIndex].Value.ToString() == kw.CurrentUser)) && (dataGridViewCertificati["IdStatus", _RowIndex].Value.ToString() == Properties.Settings.Default.KnoS_IdStato1F))
        //            //{
        //            //    int.TryParse(Properties.Settings.Default.KnoS_IdAction2F, out CurrentIdAction);
        //            //    CurrentAttrNameData = Properties.Settings.Default.KnoS_AttrNameData2F;
        //            //    SignFiles.tipofirma = 1;
        //            //}

        //            int.TryParse(dataGridViewCertificati["IdObject", _RowIndex].Value.ToString(), out CurrentIdObjectCertificato);
        //            int.TryParse(dataGridViewCertificati["IdDoc", _RowIndex].Value.ToString(), out CurrentIdDocCertificato);

        //            CurrentFileName = dataGridViewCertificati["File", _RowIndex].Value.ToString();
        //            CurrentFileDescr = dataGridViewCertificati["FileDescr", _RowIndex].Value.ToString();


        //            //tabControl1.SelectedIndex = 0;


        //            toolStripStatusLabel1.Text = "";


        //        }
        //        else
        //        {
        //            MessageBox.Show("manca il file PDF");
        //            //txtMSG.Text += string.Format("FILE PDF NON TROVATO");
        //            SignFiles.tempOriginalPDF = "";
        //        }


        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Errore nel caricamento del file PDF \r\n" + ex.Message);
        //    }        


        //}

        public static void OpenMicrosoftWord(string filePath)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "WINWORD.EXE";
            startInfo.Arguments = filePath;
            Process.Start(startInfo);
        }

        private void LoadPDFCertificato(int _RowIndex)
        {

            // resetto il backcolor della grid
            for (int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
            {
                dataGridViewCertificati.Rows[i].DefaultCellStyle.BackColor = Color.White;
            }
            dataGridViewCertificati.Rows[_RowIndex].DefaultCellStyle.BackColor = Color.Lime;


            //try
            //{
            //    string filePath = dataGridViewCertificati.Rows[_RowIndex].Cells["LocalFile"].Value.ToString();

            //    int.TryParse(dataGridViewCertificati.Rows[_RowIndex].Cells["IdObject"].Value.ToString(), out CurrentIdObjectCertificato);

            //    CurrentResponsabileTecnico = dataGridViewCertificati["ResponsabileTecnico", _RowIndex].Value.ToString();
            //    if (dataGridViewCertificati["ResponsabileTecnicoSost", _RowIndex].Value.ToString() != "")
            //    {
            //        CurrentResponsabileTecnico = dataGridViewCertificati["ResponsabileTecnicoSost", _RowIndex].Value.ToString();
            //    }

            //    CurrentCapoCommessa = dataGridViewCertificati["CapoCommessa", _RowIndex].Value.ToString();
            //    if (dataGridViewCertificati["CapoCommessaSost", _RowIndex].Value.ToString() != "")
            //    {
            //        CurrentCapoCommessa = dataGridViewCertificati["CapoCommessaSost", _RowIndex].Value.ToString();
            //    }


            //    foreach (ListViewItem li in lvCCSost.Items)
            //    {
            //        li.BackColor = Color.White;

            //        //if (li.Text == dataGridViewCertificati["CapoCommessa", _RowIndex].Value.ToString())
            //        if (li.Text == CurrentCapoCommessa)
            //        {
            //            li.BackColor = Color.Lime;
            //        }

            //    }


            //    //if (File.Exists(dataGridViewCertificati["LocalFile", _RowIndex].Value.ToString()))
            //    //{

            //    //}
            //    //else
            //    //{
            //    //    if (kw.downloadDoc(CurrentIdObjectCertificato, 1, filePath))
            //    //    {

            //    //    }
            //    //    else
            //    //    {

            //    //    }
            //    //}


            //    //if (File.Exists(dataGridViewCertificati["LocalFile", _RowIndex].Value.ToString()))
            //    //{

            //    //    SignFiles.tempOriginalPDF = dataGridViewCertificati["LocalFile", _RowIndex].Value.ToString();

                    
            //    //    toolStripStatusLabel1.Text = "Caricamento file PDF.....";

            //    //    CurrentIdObjectCertificato = int.Parse(dataGridViewCertificati["IdObject", _RowIndex].Value.ToString());


            //    //    if ((dataGridViewCertificati["Tecnico", _RowIndex].Value.ToString() == kw.CurrentUser) && (dataGridViewCertificati["IdStatus", _RowIndex].Value.ToString() == Properties.Settings.Default.KnoS_IdStatoIniziale))
            //    //    {
            //    //        int.TryParse(Properties.Settings.Default.KnoS_IdAction1F, out CurrentIdAction);
            //    //        CurrentAttrNameData = Properties.Settings.Default.KnoS_AttrNameData1F;
            //    //        SignFiles.tipofirma = 0;
            //    //    }

            //    //    if (((dataGridViewCertificati["ResponsabileTecnicoSost", _RowIndex].Value.ToString() == kw.CurrentUser) || (dataGridViewCertificati["ResponsabileTecnico", _RowIndex].Value.ToString() == kw.CurrentUser)) && (dataGridViewCertificati["IdStatus", _RowIndex].Value.ToString() == Properties.Settings.Default.KnoS_IdStato1F))
            //    //    {
            //    //        int.TryParse(Properties.Settings.Default.KnoS_IdAction2F, out CurrentIdAction);
            //    //        CurrentAttrNameData = Properties.Settings.Default.KnoS_AttrNameData2F;
            //    //        SignFiles.tipofirma = 1;
            //    //    }

            //    //    int.TryParse(dataGridViewCertificati["IdObject", _RowIndex].Value.ToString(), out CurrentIdObjectCertificato);
            //    //    int.TryParse(dataGridViewCertificati["IdDoc", _RowIndex].Value.ToString(), out CurrentIdDocCertificato);

            //    //    CurrentFileName = dataGridViewCertificati["File", _RowIndex].Value.ToString();
            //    //    CurrentFileDescr = dataGridViewCertificati["FileDescr", _RowIndex].Value.ToString();


            //    //    tabControl1.SelectedIndex = 0;


            //    //    toolStripStatusLabel1.Text = "";


            //    //}
            //    //else
            //    //{
            //    //    MessageBox.Show("manca il file PDF");
            //    //    //txtMSG.Text += string.Format("FILE PDF NON TROVATO");
            //    //    SignFiles.tempOriginalPDF = "";
            //    //}


            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Errore nel caricamento del file PDF \r\n" + ex.Message);
            //}


        }


        private void GetAzioneCertificati()
        {

            toolStripStatusLabel1.Text = "Determino stato certificati....";

            for (int j = 0; j < dataGridViewCertificati.Rows.Count; j++)
            {
                //MessageBox.Show(j.ToString());
                DataGridViewButtonCell b = (DataGridViewButtonCell)(dataGridViewCertificati.Rows[j].Cells["Firma"]);

                if ((int.Parse(dataGridViewCertificati.Rows[j].Cells["IdObject"].Value.ToString()) == SignFiles.startXML_idobject_certificato) && (int.Parse(dataGridViewCertificati.Rows[j].Cells["IdStatus"].Value.ToString()) == SignFiles.KnoS_Certificato_IdStatusIniziale))
                {
                    b.Style.ForeColor = Color.Red;
                    b.Value = "Fine Prova";
                    b.FlatStyle = FlatStyle.Popup;
                    dataGridViewCertificati.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                    //tabControl1.SelectedIndex = 0;
                }



                //if (SignFiles.tipofirma == 1)
                //{
                    if ((dataGridViewCertificati.Rows[j].Cells["ResponsabileTecnicoSost"].Value.ToString() == kw.CurrentUser) && (int.Parse(dataGridViewCertificati.Rows[j].Cells["IdStatus"].Value.ToString()) == SignFiles.KnoS_Certificato_IdStatus1F))
                    {
                        if ((string)b.Value == null)
                        {
                            b.Style.ForeColor = Color.Navy;
                            b.Value = "Applica Firma";
                            b.FlatStyle = FlatStyle.Popup;
                        }
                    }
                    else
                    {
                        if ((dataGridViewCertificati.Rows[j].Cells["ResponsabileTecnico"].Value.ToString() == kw.CurrentUser) && (int.Parse(dataGridViewCertificati.Rows[j].Cells["IdStatus"].Value.ToString()) == SignFiles.KnoS_Certificato_IdStatus1F))
                        {
                            if ((string)b.Value == null)
                            {
                                b.Style.ForeColor = Color.Navy;
                                b.Value = "Applica Firma";
                                b.FlatStyle = FlatStyle.Popup;
                            }

                        }
                    }
                //}



            }

            toolStripStatusLabel1.Text = "";
            Cursor.Current = Cursors.Default;


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


        private void btnFirmaCapoCommessa_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            string fileCertificato = "";
            string filelocale = "";
            
            nrCertificatiTot = dataGridViewCertificati.Rows.Count;

            string[] filesToMerge = new string[nrCertificatiTot];

            System.Windows.Forms.Application.DoEvents();



            for (int iRow = 0; iRow < dataGridViewCertificati.Rows.Count; iRow++)
            {

                fileCertificato = dataGridViewCertificati["Url", iRow].Value.ToString().Replace(@"file://", "");
                filelocale = dataGridViewCertificati["LocalFile", iRow].Value.ToString();
                
                toolStripStatusLabel1.Text = String.Format("File locale {0}", filelocale);

                // lo cancello se per caso ne trovo uno
                if (File.Exists(filelocale))
                {
                    File.Delete(filelocale);
                }



                if (filelocale != "")
                {
                    toolStripStatusLabel1.Text = String.Format("Download file Word linkato al certificato {0}/{1}......", (iRow + 1).ToString(), nrCertificatiTot);

                    if (fileCertificato != "")
                    {
                        if (File.Exists(fileCertificato))
                        {
                            File.Copy(fileCertificato, filelocale);
                            Word2PDF(filelocale);
                        }
                        else
                        {
                            // non è stato trovato il file del certificato
                            MessageBox.Show(string.Format("Non ho trovato il file {0} verifica se effettivamente c'è e se hai i permessi per copiarlo", fileCertificato), "Copia locale del file certificato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Cursor.Current = Cursors.Default;
                            return;
                        }
                    }

                    // preparo l'elenco 
                    filesToMerge[iRow] = Convert.ToString(filelocale).ToUpper().Replace(".DOC", ".PDF");

                }
                else
                {

                    filesToMerge[iRow] = "";
                
                }
            }


            toolStripStatusLabel1.Text = "Preparo il PDF unendo i certificati......";

            //if (nrCertificati2F == nrCertificatiTot)
            //{
                SignFiles.tempMergePDF = Path.Combine(Path.GetTempPath(), strFilePDFPDL + ".PDF");

            //}

            toolStripStatusLabel1.Text = "Inizio unione i certificati......" + SignFiles.tempMergePDF;

            CombineMultiplePDFs(filesToMerge, SignFiles.tempMergePDF);

            toolStripStatusLabel1.Text = "PDF del piano di lavoro pronto per la firma...... allego alla pubblicazione";
            

            try
            {
                //// firma con Intesi PKNET
                //KnosCSSignLibrary.SmartCardClass.Inizializza();
                //KnosCSSignLibrary.SmartCardClass.loadEnvelope();

                //KnosCSSignLibrary.CustomSignData csd;
                //KnosCSSignLibrary.SmartCardClass.loadSignerInfo();
                //File.Delete(SignFiles.tempMergePDF + "_signed.pdf");

                //// firma
                //KnosCSSignLibrary.Sign.SignReason = "Firma certificata";
                //KnosCSSignLibrary.Sign.SignLocation = "KnoS";


                //if (KnosCSSignLibrary.SmartCardClass.PDFSignFileEX(SignFiles.tempMergePDF, SignFiles.tempMergePDF + "_signed.pdf", cboSmartCardCert.SelectedItem.ToString(), System.DateTime.Now) == true)
                //{

                bool bFirmaOK = false;
                int nrTentativiFirma = 0;

                while (nrTentativiFirma < 3)
                {
                    nrTentativiFirma += 1;

                    bFirmaOK = ApplicaFirmaCertificato();

                    if (bFirmaOK == true)
                    {
                        break;
                    }
                }

                if (bFirmaOK == true)
                {
                    //// firma andata a buon fine
                    //KnosCSSignLibrary.SmartCardClass.unloadEnvelope();


                    if (kw.UploadFileCertificato(CurrentIdObject, 0, SignFiles.tempMergePDF + "_signed.pdf", strFilePDFPDL, Path.GetFileName(SignFiles.tempMergePDF + "_signed.pdf"), SignFiles.KnoS_IdActionPDFFirmato, Properties.Settings.Default.KnoS_AttrNameDataFirmaPDL))
                    {
                        toolStripStatusLabel1.Text = "Ricarico il PDL firmato...............";

                        try
                        {
                            for (int i = 0; i < filesToMerge.Length; i++)
                            {
                                File.Delete(filesToMerge[i].ToString());
                            }

                            File.Delete(SignFiles.tempMergePDF + "_signed.pdf");
                            File.Delete(SignFiles.tempMergePDF);
                        }
                        catch (Exception ex)
                        {

                        }

                    }


                    // transizione di stato dei certficati
                    for (int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
                    {
                        toolStripStatusLabel1.Text = string.Format("Transizione di stato certificato {0} ....", dataGridViewCertificati["IdObject", i].Value);
                        if (kw.ActionCertificato(int.Parse(dataGridViewCertificati["IdObject", i].Value.ToString()), SignFiles.KnoS_IdActionPDLFirmatoCERT, "") == false)
                        {
                            MessageBox.Show(string.Format("La transizione di stato del certificato con IdObject {0} in Firmato Capo Commessa NON è andata a buon fine", dataGridViewCertificati["IdObject", i].Value), "Transizione di stato", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    toolStripStatusLabel1.Text = "";

                    //ricarico il PDL
                    MessageBox.Show("Procedura firma digitale completata, attendere il caricamento del PDL firmato", "Firma digitale PDL", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //LoadPDL();


                }
                else
                {
                    //KnosCSSignLibrary.SmartCardClass.unloadEnvelope();
                    MessageBox.Show("Provare a firmare nuovamente il Piano di Lavoro", "Firma digitale PDL", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                }
                
                LoadPDL();
                
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Accertarsi di aver inserito la BusinessKey e caricare il certificato digitale", "Firma digitale PDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }

            toolStripStatusLabel1.Text = "";

        }


        private bool ApplicaFirmaCertificato()
        {
            bool bOK = false;

            // firma con Intesi PKNET
            KnosCSSignLibrary.SmartCardClass.Inizializza();
            KnosCSSignLibrary.SmartCardClass.loadEnvelope();

            KnosCSSignLibrary.CustomSignData csd;
            KnosCSSignLibrary.SmartCardClass.loadSignerInfo();
            File.Delete(SignFiles.tempMergePDF + "_signed.pdf");

            // firma
            KnosCSSignLibrary.Sign.SignReason = "Firma certificata";
            KnosCSSignLibrary.Sign.SignLocation = "KnoS";

            bOK = KnosCSSignLibrary.SmartCardClass.PDFSignFileEX(SignFiles.tempMergePDF, SignFiles.tempMergePDF + "_signed.pdf", cboSmartCardCert.SelectedItem.ToString(), System.DateTime.Now);

            // firma andata a buon fine
            KnosCSSignLibrary.SmartCardClass.unloadEnvelope();

            return bOK;
            
        }



        // converto il doc in PDF
        private bool Word2PDF(string wordFile)
        {

            bool bOK = false;

            // Create a new Microsoft Word application object
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {

                // C# doesn't have optional arguments so we'll need a dummy value
                object oMissing = System.Reflection.Missing.Value;

                word.Visible = false;
                word.ScreenUpdating = false;

                // Cast as Object for word Open method
                Object filename = (Object)wordFile;

                // Use the dummy value as a placeholder for optional arguments
                doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                object outputFileName = wordFile.Replace(".DOC", ".PDF");
                object fileFormat = WdSaveFormat.wdFormatPDF;

                doc.DeleteAllEditableRanges(WdEditorType.wdEditorEveryone);
                word.ScreenUpdating = false;

                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.                
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;

                bOK = true;
            }
            catch (Exception ex)
            {

            }
            finally
            {

                word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                word.Quit();
                if (doc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                if (word != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(word);

                doc = null;
                word = null;
                GC.Collect(); // force final cleanup! 
            }

            return bOK;

        }


        private void button2_Click(object sender, EventArgs e)
        {
            kw.CurrentUser = txtKnoSUser.Text;
            kw.CurrentPWD = txtKnoSPassword.Text;
            kw.Inizializza(txtKnosUrl.Text);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Form_Prefs f = new Form_Prefs();
            f.ShowDialog();

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            loadSostituti();
            LoadPDL();
        }


        private bool SendNotify(string _address, string _subject, string _body)
        {
            if (Properties.Settings.Default.useMapiMail == true)
            {

                if (_address != "")
                {
                    SendFileTo.MAPI mapi = new MAPI();
                    
                    mapi.AddRecipientTo(_address);

 //                    mapi.AddAttachment(strFilePDF);
                    _body += "\r\n" + string.Format(" per aprire il piano di lavoro col programma di firma fai doppio click sull'allegato oppure \r\ncopia la seguente stringa ed incollala sul percorso di esplora risorse o nel browser \r\nknosapi:OpenTecEurolabCertificatiPDL?id={0}&baseurl={1}", SignFiles.startXML_idobject, SignFiles.startXML_baseurl);


//                    <?xml version="1.0" encoding="utf-8"?>
//                        <KnosEnvelope knosBaseUrl="http://tecsql03" 
                            //knosVersion="7.1.2" 
                            //envelopeVersion="1.0" 
                            //contains="PDL">  
                            //<PDL IdObject="505804" IdObjectCertificato="505805" NotificaCapocommessa="false" />
//                        </KnosEnvelope>
                    
//"

                    string linkFile = Path.Combine(Path.GetTempPath(), "ApriPDL.knos-fr");
                    XmlWriter xmlWriter = XmlWriter.Create(linkFile);

                 
                    xmlWriter.WriteStartDocument();
                    xmlWriter.WriteStartElement("KnosEnvelope");
                    xmlWriter.WriteAttributeString("knosBaseUrl", txtKnosUrl.Text);
                    xmlWriter.WriteAttributeString("knosVersion", "7.1.2");
                    xmlWriter.WriteAttributeString("contains", "PDL");

                    xmlWriter.WriteStartElement("PDL");
                    xmlWriter.WriteAttributeString("IdObject", CurrentIdObject.ToString());
                    xmlWriter.WriteAttributeString("NotificaCapocommessa", "false");

                    xmlWriter.WriteEndElement();

                    xmlWriter.WriteEndElement();
                    
                    xmlWriter.WriteEndDocument();
                    xmlWriter.Close();

                    mapi.AddAttachment(linkFile);
                    mapi.SendMailPopup(_subject, _body);

                }

            }
            else
            {

                try
                {
                    toolStripProgressBar1.Step = 1;
                    toolStripProgressBar1.Minimum = 0;
                    toolStripProgressBar1.Maximum = 7;
                    toolStripProgressBar1.Value = 0;
                    toolStripProgressBar1.Visible = true;
                    toolStripProgressBar1.Width = statusStrip1.Width - toolStripStatusLabel1.Width - 50;

                    // Create the Outlook application by using inline initialization.
                    Outlook.Application oApp = new Outlook.Application();
                    toolStripStatusLabel1.Text = "inizializzo Outlook....";
                    toolStripProgressBar1.PerformStep();

                    // survive to grant access....
                    Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
                    Outlook.MAPIFolder f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    System.Threading.Thread.Sleep(5000);

                    toolStripStatusLabel1.Text = "inizializzo Messaggio....";
                    toolStripProgressBar1.PerformStep();

                    //Create the new message by using the simplest approach.
                    Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                    //Add a recipient.
                    // TODO: Change the following recipient where appropriate.
                    Outlook.Recipient oRecip;
                    if (_address != "")
                    {
                        oRecip = (Outlook.Recipient)oMsg.Recipients.Add(_address);
                        oRecip.Resolve();
                    }
                    toolStripStatusLabel1.Text = "inizializzo Destinatario....";
                    toolStripProgressBar1.PerformStep();

                    //Set the basic properties.
                    oMsg.Subject = _subject;// "This is the subject of the test message";

                    oMsg.Body = _body; // "This is the text in the message.";
                    oMsg.HTMLBody += "\n\r" + string.Format("<a href=\"knosapi:OpenTecEurolabCertificatiPDL?id={0}&baseurl={1}\">Firma</a>", SignFiles.startXML_idobject, SignFiles.startXML_baseurl);
                    toolStripStatusLabel1.Text = "inizializzo titolo e corpo maessaggio....";
                    toolStripProgressBar1.PerformStep();

                    // 11/03/2014 - rimosso l'allegato dalla mail in quanto aggiunto olink diretto nel corpo mail

                    //Outlook.Attachment oAttach;

                    //if (File.Exists(SignFiles.startXML) == true)
                    //{
                    //    String sSource = SignFiles.startXML;
                    //    String sDisplayName = "Apri Programma per la Firma";
                    //    int iPosition = (int)oMsg.Body.Length + 1;
                    //    int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                    //    oAttach = oMsg.Attachments.Add(sSource, iAttachType, iPosition, sDisplayName);
                    //}
                    //else
                    //{ 
                    //    // aggiungere link al PDL

                    //}

                    toolStripStatusLabel1.Text = "inizializzo Allegato per apertura programma firma....";
                    toolStripProgressBar1.PerformStep();


                    // If you want to, display the message.
                    if (notifyPopUp == true)
                    {
                        oMsg.Display(true);  //modal
                        //oMsg.Save();
                    }
                    else
                    {
                        //Send the message.
                        oMsg.Save();
                        oMsg.Send();
                    }


                    //Explicitly release objects.
                    oRecip = null;
                    //oAttach = null;
                    oMsg = null;
                    oApp = null;
                }

                                // Simple error handler.
                catch (Exception e)
                {
                    MessageBox.Show(string.Format("Messaggio da Outlook: \r\n {0} ", e.Message), "Invio Notifica");
                    toolStripStatusLabel1.Text = "";
                    toolStripProgressBar1.Visible = false;
                    return true;

                }
                finally
                {
                    toolStripStatusLabel1.Text = "";
                    toolStripProgressBar1.Visible = false;
                }

            }
            //Default return value.
            return true;
        
        }

        private void button4_Click_1(object sender, EventArgs e)
        {

            

        }

        private void btnPDLStatus_Click(object sender, EventArgs e)
        {
            if (webBrowser1.Visible == true)
            {
                webBrowser1.Navigate("about:blank");
                webBrowser1.SendToBack();
                webBrowser1.Visible = false;
                panelBrowser.Width = 50;
                panelBrowser.SendToBack();
                panelBrowser.Visible = false;

            }
            else
            {
                if (btnPDLStatus.Tag != null)
                {
                    webBrowser1.Navigate(btnPDLStatus.Tag.ToString());
                    webBrowser1.BringToFront();
                    webBrowser1.Visible = true;

                    panelBrowser.Width = panelBrowser.Parent.Width - 20;
                    panelBrowser.BringToFront();
                    panelBrowser.Visible = true;
                }
            }
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

        private void btnGetCertificates_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            toolStripStatusLabel1.Text = "Caricamento certificati digitali in corso.......";

            if (cboTipoDispositivo.SelectedIndex == 0)
            {
                //demo (locali)
                GetCertificates();
            }
            else
            {
                //smartcard
                if (GetCertificatesSC())
                {
                    btnGetCertificates.ImageIndex = 4;
                }
                else
                {
                    Cursor.Current = Cursors.Default;
                    return;
                }
            }

            Cursor.Current = Cursors.Default;

            if (cboTipoDispositivo.SelectedIndex == 0)
            {
                Utils.ComboItem<Certificato> c = (Utils.ComboItem<Certificato>)cboCertificates.SelectedItem;

                MessageBox.Show(c.Obj.Nome, "Informazioni sul certificato selezionato", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                Utils.ComboItem<Certificato> c = (Utils.ComboItem<Certificato>)cboSmartCardCert.SelectedItem;

                //MessageBox.Show(c.Obj.Nome, "Informazioni sul certificato selezionato", MessageBoxButtons.OK, MessageBoxIcon.Information);

                btnDettagliCertificato.BackColor = Color.Lime;

                KnosCSSignLibrary.SmartCardClass.Inizializza();
                KnosCSSignLibrary.SmartCardClass.loadEnvelope();
                KnosCSSignLibrary.SmartCardClass.loadSignerInfo();
            }


        }

        private void loadCert()
        {
            try
            {
                Card card = cboSmartCards.SelectedItem as Utils.ComboItem<Card>;
                cboSmartCardCert.Items.Clear();
                foreach (KeyValuePair<String, Certificato> kvpCert in card.Certificati)
                {
                    cboSmartCardCert.Items.Add(new Utils.ComboItem<Certificato>(kvpCert.Value, kvpCert.Value.Alias));
                }

                if (cboSmartCardCert.Items.Count > 0) cboSmartCardCert.SelectedIndex = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }



        private void GetCertificates()
        {
            Cursor c;
            toolStripStatusLabel1.Text = "Caricamento certificati digitali in corso.......";

            try
            {
                c = Cursors.WaitCursor;

                if (cboTipoDispositivo.SelectedIndex == 0)
                {
                    //demo (locali)
                }
                else
                {
                    //smartcard
                }



                // controllo dispositivo connesso
                //System.Threading.Thread.Sleep(2000);

                //string PIN = "";
                //if (customControls.InputBox("Inserire PIN:", "Dispositivo Firme Digitali", ref PIN) == DialogResult.OK)
                //{
                //    if (PIN != "sash17ne")
                //    {
                //        toolStripStatusLabel1.Text = "";
                //        MessageBox.Show("Inserire il dispositivo e digitare il PIN corretto!", "Dispositivo Firme Digitali", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1);
                //        c = Cursors.Default;
                //        return;
                //    }
                //}


                KnosCSSignLibrary.Sign.GetCertificates();

                cboCertificates.DataSource = new BindingSource(KnosCSSignLibrary.Sign.SignCertificates, null);

                cboCertificates.DisplayMember = "Value";

                cboCertificates.ValueMember = "Key";

                btnGetCertificates.ImageIndex = 4;
                toolStripStatusLabel1.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                toolStripStatusLabel1.Text = "";
                c = Cursors.Default;
            }
        }


        private bool GetCertificatesSC()
        {
            Cursor c;
            bool bOK = false;

            toolStripStatusLabel1.Text = "Caricamento certificati digitali in corso.......";

            try
            {
                c = Cursors.WaitCursor;

                SmartCardClass.Inizializza();
                SmartCardClass.RilevaSmartCard();


                cboSmartCards.Items.Clear();

                foreach (KeyValuePair<String, Card> kvpCards in SmartCardClass.SmartCardList)
                {
                    cboSmartCards.Items.Add(new Utils.ComboItem<Card>(kvpCards.Value, kvpCards.Value.Nome));
                }

                cboSmartCards.SelectedIndex = 0;

                loadCert();
                bOK = true;



            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                MessageBox.Show("Verificare che il dispositivo di firma sia collegato e attivo", "Dispositivo di firma NON rilevato", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            toolStripStatusLabel1.Text = "";
            return bOK;

        }




        private void btnDettagliCertificato_Click(object sender, EventArgs e)
        {
            if (cboTipoDispositivo.SelectedIndex == 0)
            {
                Utils.ComboItem<Certificato> c = (Utils.ComboItem<Certificato>)cboCertificates.SelectedItem;

                MessageBox.Show(c.Obj.Nome, "Informazioni sul certificato selezionato", MessageBoxButtons.OK, MessageBoxIcon.Information); 
            }
            else
            {
                Utils.ComboItem<Certificato> c = (Utils.ComboItem<Certificato>)cboSmartCardCert.SelectedItem;

                MessageBox.Show(c.Obj.Nome, "Informazioni sul certificato selezionato", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cboTipoDispositivo_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboSmartCards.Visible = cboSmartCardCert.Visible = (cboTipoDispositivo.SelectedIndex == 1);
            cboCertificates.Visible = (cboTipoDispositivo.SelectedIndex == 0);
        }

        private void btnSendMail_Click(object sender, EventArgs e)
        {

            string _tecnico = "";
            string _responsabile = "";
            string _capocommessa = "";
            string _mailaddress = "";
            string _mailsubject = "";
            string _mailbody = "";
            string _dettaglipdl = "";



                //for (int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
                //{
                //    //if (int.Parse(dataGridViewCertificati["IdObject", i].Value.ToString()) == CurrentIdObjectCertificato)
                //    //{
                //    _tecnico = dataGridViewCertificati["Tecnico", i].Value.ToString();
                //    _responsabile = dataGridViewCertificati["ResponsabileTecnico", i].Value.ToString();
                //    _capocommessa = dataGridViewCertificati["CapoCommessa", i].Value.ToString();
                //    //break;
                //    //}
                //}


            
            if ((nrCertificati2F == nrCertificatiTot) && (_responsabile == kw.CurrentUser))
            {

                // transizione di stato del PDL in "In attesa di firma"
                if (SignFiles.KnoS_IdActionPDFdaFirmare > 0)
                {
                    kw.UploadFileCertificato(CurrentIdObject, 0, "", "", "", SignFiles.KnoS_IdActionPDFdaFirmare, "");
                }

                // notifica al capo commessa
                _mailsubject = string.Format("Notifica al Capo Commessa", "");
                _mailbody = Properties.Settings.Default.sendMailCapocommessaMessage;
                _mailaddress = kw.GetEmailSubjectByName(_capocommessa);


                for (int liItem = 0; liItem < listViewAttr.Items.Count; liItem++)
                {
                    _dettaglipdl += string.Format("\r\n  - {0} - {1}", listViewAttr.Items[liItem].SubItems[0].Text, listViewAttr.Items[liItem].SubItems[1].Text);
                }

                _mailbody = string.Format(_mailbody, _dettaglipdl, _capocommessa);

                MessageBox.Show(string.Format("Notifica al Capo Commessa {0} - {1} ", _capocommessa, _mailaddress));

                SendNotify(_mailaddress, _mailsubject, _mailbody);

            }


            //if (nrCertificati1F >= nrCertificatiUtente1F)
            if ((nrCertificatiUtente1FDaFirmare == 0) && (_tecnico == kw.CurrentUser) && (SignFiles.tipofirma == 0))
            {

                // notifica al capo commessa
                _mailsubject = string.Format("Notifica al Responsabile Tecnico", "");
                _mailbody = Properties.Settings.Default.sendMailResponsabileMessage;
                _mailaddress = kw.GetEmailSubjectByName(_responsabile);


                for (int liItem = 0; liItem < listViewAttr.Items.Count; liItem++)
                {
                    _dettaglipdl += string.Format("\r\n  - {0} - {1}", listViewAttr.Items[liItem].SubItems[0].Text, listViewAttr.Items[liItem].SubItems[1].Text);
                }

                _mailbody = string.Format(_mailbody, _dettaglipdl, _responsabile);

                MessageBox.Show(string.Format("Notifica al Responsabile Tecnico {0} - {1} ", _responsabile, _mailaddress));

                SendNotify(_mailaddress, _mailsubject, _mailbody);

            }


            if (_mailaddress == "")
            {
                if (dataGridViewCertificati.SelectedRows.Count > 0)
                {

                    _tecnico = dataGridViewCertificati["Tecnico", dataGridViewCertificati.SelectedRows[0].Index].Value.ToString();
                    //_responsabile = dataGridViewCertificati["ResponsabileTecnico", dataGridViewCertificati.SelectedRows[0].Index].Value.ToString();
                    _capocommessa = dataGridViewCertificati["CapoCommessa", dataGridViewCertificati.SelectedRows[0].Index].Value.ToString();
                }
                else
                { 
                  
                
                }

                _mailaddress = string.Format("{0};{1};{2}",kw.GetEmailSubjectByName(_tecnico),kw.GetEmailSubjectByName(_responsabile),kw.GetEmailSubjectByName(_capocommessa));


                for (int liItem = 0; liItem < listViewAttr.Items.Count; liItem++)
                {
                    _dettaglipdl += string.Format("\r\n  - {0} - {1}", listViewAttr.Items[liItem].SubItems[0].Text, listViewAttr.Items[liItem].SubItems[1].Text);
                }

                _mailsubject = "Testo";
                _mailbody = string.Format(_dettaglipdl);
                //_mailbody += "\n\r" + string.Format("<a href=\"knosapi:OpenTecEurolabCertificatiPDL?id={0}&baseurl={1}\">Firma</a>", SignFiles.startXML_idobject, SignFiles.startXML_baseurl);
                SendNotify(_mailaddress, _mailsubject, _mailbody);
            }

        }


        private void loadSostituti()
        {
            DataTable x = kw.GetSostituti();

            lvCCSost.Items.Clear();
            lvFileFirma.Items.Clear();

            for (int i=0; i <x.Rows.Count; i++)
            {


                if (x.Rows[i][1].ToString() == "1")
                {
                    lvCCSost.Items.Add(x.Rows[i][2].ToString());
                    lvCCSost.Items[lvCCSost.Items.Count-1].SubItems.Add(x.Rows[i][3].ToString());
                }

                if (x.Rows[i][2].ToString() == kw.CurrentUser)
                {
                    lvFileFirma.Items.Add(x.Rows[i][2].ToString());
                    lvFileFirma.Items[lvFileFirma.Items.Count-1].SubItems.Add(x.Rows[i][3].ToString());

                }
            }
        }

        private void lvSost_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            bool rt = false;
            string msgText = "Vuoi impostare {0} come Capo Commessa Sostituto?";
            string msgCaption = "Cambio Capo Commessa";

            if (((ListView)sender).Name == "lvRTSost")
            {
                rt = true;
                msgText = "Vuoi impostare {0} come Responsabile Tecnico Sostituto?";
                msgCaption = "Cambio Responsabile Tecnico ";
            }


            if (MessageBox.Show(string.Format(msgText, ((ListView)sender).SelectedItems[0].Text), msgCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                if (rt)
                { 
                    // idobject certificato
                    if (kw.SetSostituto(CurrentIdObjectCertificato, "varchar_54", ((ListView)sender).SelectedItems[0].Text) == true)
                    {
                        MessageBox.Show("Responsabile Tecnico Sostituto assegnato correttamente!", msgCaption);
                    }
                }
                else
                { 
                    // tutti i certificati
                    foreach (DataGridViewRow drv in dataGridViewCertificati.Rows)
                    {
                        // idobject certificato
                        if (kw.SetSostituto(int.Parse(drv.Cells["IdObject"].Value.ToString()), "varchar_55", ((ListView)sender).SelectedItems[0].Text) == false)
                        {
                            MessageBox.Show("ERRORE nell'assegnazione del Capo Commessa Sostituto assegnato correttamente! \r\n VERIFICARE LA PUBBLICAZIONE DEL CERTIFICATO IN KNOS", msgCaption);
                            break;
                        }
                    }
                
                    MessageBox.Show("Capo Commessa Sostituto assegnato correttamente!", msgCaption);


                    toolStripStatusLabel1.Text = "Ricarico il PDL......";
                    LoadPDL();
                    toolStripStatusLabel1.Text = "";

                }


            
            }

        }

        private void dataGridViewCertificati_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Uri u = new Uri(dataGridViewCertificati.Rows[e.RowIndex].Cells["Url"].Value.ToString());

            OpenMicrosoftWord(u.ToString());
            //webBrowser1.Url = u;
            //webBrowser1.Navigate(u);
            ////webBrowser1.Width = panelBrowser.Width;
            //webBrowser1.Visible = true;
            //panelBrowser.Visible = true;
            //panelBrowser.Width = panelBrowser.Parent.Width - 20;
            //panelBrowser.BringToFront();
        }

        private void btnCloseWebBrowser_Click(object sender, EventArgs e)
        {
            panelBrowser.Visible = false;
            panelBrowser.Width = 50;
            panelBrowser.SendToBack();
        }

        private void btnMyCertificates_Click(object sender, EventArgs e)
        {
            //dataGridViewMyCertificates.DataSource = kw.GetMyCertificates();
        }

        private void btnSchedaPDL_Click(object sender, EventArgs e)
        {
            //tabControl1.SelectedIndex = 1;
        }

        private void btnFirmaPDF_Click(object sender, EventArgs e)
        {
        //    if (!File.Exists(lblPNGFirma.Text))
        //    {
        //        MessageBox.Show("Scegliere un file firma da applicare al documento");
        //        return;
        //    }



        //    Byte[] pageBytes;
        //    bool bFound = false;
        //    iTextSharp.text.Image maskImage;

        //    string fileNameExisting = SignRTFPDF.SignFiles.tempOriginalPDF;
        //    string fileNameNew = SignRTFPDF.SignFiles.tempSignedPDF;

        //    txtMSG.Text = string.Format("Inizio procedura firma con sostituzione immagini segnaposto");
        //    txtMSG.Text += string.Format("\r\n - file: {0}", fileNameExisting);


        //    iTextSharp.text.pdf.PdfReader pdf = new iTextSharp.text.pdf.PdfReader(fileNameExisting);

        //    iTextSharp.text.pdf.PdfStamper stp = new iTextSharp.text.pdf.PdfStamper(pdf, new FileStream(fileNameNew,
        //    FileMode.Create));
        //    iTextSharp.text.pdf.PdfWriter writer = stp.Writer;
        //    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(lblPNGFirma.Text);

        //    // da verificare
        //    //List<string> signaturenames = pdf.AcroFields.GetSignatureNames();


        //    //PRTokeniser token;
        //    //string tknValue = String.Empty;
        //    //PRTokeniser.TokType tknType;

        //    //for(int i=0; i<signaturenames.Count; i++)
        //    //{
        //    //    txtMSG.Text += string.Format("\r\n signaturename {0} : {1}", i.ToString(), signaturenames[i].ToString());
        //    //}

        //    txtMSG.Text += string.Format("\r\n -Nr Pagine: {0}", pdf.NumberOfPages);


        //    if (SignFiles.tipofirma == -1)
        //    {
        //        txtMSG.Text += string.Format("\r\n tipo firma : {0} ------------", "NESSUNO");
        //    }

        //    if (SignFiles.tipofirma == 0)
        //    {
        //        txtMSG.Text += string.Format("\r\n tipo firma : {0} ------------", "TECNICO");
        //    }
        //    if (SignFiles.tipofirma == 1)
        //    {
        //        txtMSG.Text += string.Format("\r\n tipo firma : {0} ------------", "RESPONSABILE TECNICO");
        //    }
        //    if (SignFiles.tipofirma == 2)
        //    {
        //        txtMSG.Text += string.Format("\r\n tipo firma : {0} ------------", "CAPO COMMESSA");
        //    }

        //    pictureBox1.BackColor = pictureBox2.BackColor = pictureBox3.BackColor = Color.Transparent;

        //    string tempB0 = "";
        //    tempB0 = Path.Combine(System.Windows.Forms.Application.StartupPath, "temp0.bmp");
        //    File.Delete(tempB0);
        //    pictureBox1.Image.Save(tempB0);
        //    string tempB1 = "";
        //    tempB1 = Path.Combine(System.Windows.Forms.Application.StartupPath, "temp1.bmp");
        //    File.Delete(tempB1);
        //    pictureBox2.Image.Save(tempB1);
        //    string tempB2 = "";
        //    tempB2 = Path.Combine(System.Windows.Forms.Application.StartupPath, "temp2.bmp");
        //    File.Delete(tempB2);
        //    pictureBox3.Image.Save(tempB2);



        //    for (int i = 1; i <= pdf.NumberOfPages; i++)
        //    {
        //        //txtMSG.Text += string.Format("\r\n ---------- page: {0} ------------", i.ToString());
        //        #region todo
        //        /*                

        //                        pageBytes = pdf.GetPageContent(i);

        //                        if (pageBytes != null)
        //                        {
        //                            RandomAccessFileOrArray r = new RandomAccessFileOrArray(pageBytes);
        //                            token = new PRTokeniser(r);

        //                            while(token.NextToken())
        //                            {


        //                                tknType = token.TokenType;
        //                                tknValue = token.StringValue;
        //                                //txtMSG.Text += string.Format("\r\n type: {0} - value: {1}", tknType.ToString(), tknValue.ToString());

        //                                if ((tknType == PRTokeniser.TokType.STRING) )
        //                                {
        //                                    txtMSG.Text += string.Format("\r\n type: {0} - value: {1}", tknType.ToString(), tknValue.ToString());

        //                                    if (tknValue == "firma")
        //                                    {
        //                                        //sb.Append(token.StringValue)
        //                                        MessageBox.Show("found");
        //                                    }
        //                                }

        //                            }
        //                        }

        //        */

        //        #endregion todo
        //        // sostituzione immagine

        //        iTextSharp.text.pdf.PdfDictionary pg = pdf.GetPageN(i);
        //        iTextSharp.text.pdf.PdfDictionary res =
        //        (iTextSharp.text.pdf.PdfDictionary)iTextSharp.text.pdf.PdfReader.GetPdfObject(pg.Get(iTextSharp.text.pdf.PdfName.RESOURCES));
        //        iTextSharp.text.pdf.PdfDictionary xobj =
        //        (iTextSharp.text.pdf.PdfDictionary)iTextSharp.text.pdf.PdfReader.GetPdfObject(res.Get(iTextSharp.text.pdf.PdfName.XOBJECT));
        //        if (xobj != null)
        //        {
        //            foreach (iTextSharp.text.pdf.PdfName name in xobj.Keys)
        //            {
        //                iTextSharp.text.pdf.PdfObject obj = xobj.Get(name);
        //                if (obj.IsIndirect())
        //                {
        //                    iTextSharp.text.pdf.PdfDictionary tg = (iTextSharp.text.pdf.PdfDictionary)iTextSharp.text.pdf.PdfReader.GetPdfObject(obj);

        //                    if (tg != null)
        //                    {
        //                        iTextSharp.text.pdf.PdfName type =
        //                        (iTextSharp.text.pdf.PdfName)iTextSharp.text.pdf.PdfReader.GetPdfObject(tg.Get(iTextSharp.text.pdf.PdfName.SUBTYPE));
        //                        if (iTextSharp.text.pdf.PdfName.IMAGE.Equals(type))
        //                        {
        //                            iTextSharp.text.pdf.PdfReader.KillIndirect(obj);

        //                            iTextSharp.text.Image newimg = iTextSharp.text.Image.GetInstance(lblPNGFirma.Text);

        //                            byte[] bytes = iTextSharp.text.pdf.PdfReader.GetStreamBytesRaw((iTextSharp.text.pdf.PRStream)tg);

        //                            if ((bytes != null))
        //                            {
        //                                try
        //                                {
        //                                    iTextSharp.text.Image foundimg = iTextSharp.text.Image.GetInstance(bytes);

        //                                    System.IO.MemoryStream MS = new System.IO.MemoryStream(bytes);



        //                                    MS.Position = 0;
        //                                    System.Drawing.Image ImgPDF = System.Drawing.Image.FromStream(MS);
        //                                    pictureBox4.Image = ImgPDF;

        //                                    string temp = Path.Combine(System.Windows.Forms.Application.StartupPath, "temp.bmp");
        //                                    File.Delete(temp);

        //                                    ImgPDF.Save(temp);
        //                                    FileInfo fi = new FileInfo(temp);
        //                                    string lenTemp = fi.Length.ToString();
        //                                    //decimal lenFS = GetFileSizeOnDisk(temp);

        //                                    txtMSG.Text += string.Format("\r\n - trovata immagine W:{0} H:{1} L:{2}", foundimg.Width, foundimg.Height, 0);

        //                                    bFound = false;



        //                                    if (CheckImagesByPixel(temp, tempB0) == true)
        //                                    {
        //                                        bFound = (SignFiles.tipofirma == 0);
        //                                        pictureBox1.BackColor = Color.Green;
        //                                        txtMSG.Text += string.Format("\r\n - trovata immagine segnaposto TECNICO");
        //                                    }


        //                                    if (bFound == false)
        //                                    {
        //                                        if ((CheckImagesByPixel(temp, tempB1) == true))
        //                                        {
        //                                            bFound = (SignFiles.tipofirma == 1);
        //                                            pictureBox2.BackColor = Color.Green;
        //                                            txtMSG.Text += string.Format("\r\n - trovata immagine segnaposto RESPONSABILE TECNICO");
        //                                        }
        //                                    }

        //                                    if (bFound == false)
        //                                    {
        //                                        if ((CheckImagesByPixel(temp, tempB2) == true))
        //                                        {
        //                                            bFound = (SignFiles.tipofirma == 2);
        //                                            pictureBox3.BackColor = Color.Green;
        //                                            txtMSG.Text += string.Format("\r\n - trovata immagine segnaposto CAPO COMMESSA");
        //                                        }
        //                                    }


        //                                    if (bFound == true)
        //                                    {
        //                                        txtMSG.Text += string.Format("\r\n - immagine sostituita con immagine firma");
        //                                        maskImage = newimg.ImageMask;
        //                                        if (maskImage != null)
        //                                            writer.AddDirectImageSimple(maskImage);
        //                                        writer.AddDirectImageSimple(newimg, (iTextSharp.text.pdf.PRIndirectReference)obj);
        //                                        //break;
        //                                    }
        //                                    else
        //                                    {
        //                                        maskImage = foundimg.ImageMask;
        //                                        if (maskImage != null)
        //                                            writer.AddDirectImageSimple(maskImage);
        //                                        writer.AddDirectImageSimple(foundimg, (iTextSharp.text.pdf.PRIndirectReference)obj);
        //                                        //break;
        //                                    }





        //                                }
        //                                catch (Exception ex)
        //                                {
        //                                    //                                            MessageBox.Show(ex.Message);
        //                                    txtMSG.Text += string.Format("\r\n - errore in gestione immagine\r\n - " + ex.Message);


        //                                }
        //                            }



        //                        }
        //                    }

        //                }
        //            }

        //        }


        //    }

        //    stp.FormFlattening = true;

        //    stp.Close();

        //    pdf.Close();

        //    //webBrowser1.Dispose();



        //    File.Copy(SignFiles.tempSignedPDF, SignFiles.tempOriginalPDF, true);
        //    axAcroPDF1.LoadFile(SignFiles.tempSignedPDF);

        //    MessageBox.Show("Firma Applicata! \r\n ATTENDERE il messaggio di completamento dell'upload del file PDF alla pubblicazione certificato.", "Inserimento firma su certicato PDF", MessageBoxButtons.OK);

        //    toolStripStatusLabel1.Text = "Inizio upload file...";

        //    if (kw.UploadFileCertificato(CurrentIdObject, CurrentIdDocCertificato, SignFiles.tempOriginalPDF, CurrentFileDescr, CurrentFileName, CurrentIdAction, CurrentAttrNameData) == true)
        //    {
        //        //kw.GetPDL(CurrentIdObject, listViewAttr, dataGridViewCertificati, lvFileFirma);

        //        // stato PDL
        //        btnPDLStatus.Text = CurrentStatusNamePDL;
        //        btnPDLStatus.Tag = CurrentPDFPDLUrl;


        //        //if (SignFiles.tipofirma == 0)
        //        //{
        //        //    for (int j = 0; j < dataGridViewCertificati.Rows.Count; j++)
        //        //    {
        //        //        if (int.Parse(dataGridViewCertificati.Rows[j].Cells["IdObject"].Value.ToString()) == SignFiles.startXML_idobject_certificato)
        //        //        {
        //        //            DataGridViewButtonCell b = (DataGridViewButtonCell)(dataGridViewCertificati.Rows[j].Cells["Firma"]);

        //        //            b.Style.ForeColor = Color.Red;
        //        //            b.Value = "Fine Prova";
        //        //            b.FlatStyle = FlatStyle.Popup;
        //        //        }



        //        //    }
        //        //}

        //        LoadPDL();

        //        MessageBox.Show("Upload del documento avvenuto con successo!", "Pubblicazione certificato");
        //        string _tecnico = "";
        //        string _responsabile = "";
        //        string _capocommessa = "";
        //        string _mailaddress = "";
        //        string _mailsubject = "";
        //        string _mailbody = "";

        //        // notifica al responsabile tecnico
        //        for (int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
        //        {
        //            if (int.Parse(dataGridViewCertificati["IdObject", i].Value.ToString()) == CurrentIdObjectCertificato)
        //            {
        //                _tecnico = dataGridViewCertificati["Tecnico", i].Value.ToString();
        //                _responsabile = dataGridViewCertificati["ResponsabileTecnico", i].Value.ToString();
        //                _capocommessa = dataGridViewCertificati["CapoCommessa", i].Value.ToString();
        //                break;
        //            }
        //        }


        //        toolStripStatusLabel1.Text = "Verifica Notifiche...";

        //        if ((nrCertificatiUtente2FDaFirmare == 0) && (_responsabile == kw.CurrentUser))
        //        {
        //            // notifica al capo commessa
        //            MessageBox.Show(string.Format("Responsabile Tecnico {0} : Su TUTTI i certificati è stata apposta la firma", _responsabile));



        //        }

        //        if ((nrCertificati2F == nrCertificatiTot) && (_responsabile == kw.CurrentUser))
        //        {
        //            // notifica al capo commessa
        //            _mailaddress = kw.GetEmailSubjectByName(_capocommessa);

        //            _mailbody = string.Format("Notifica al Capo commessa {0} - da parte del Responsabile Tecnico {1}: ", _capocommessa, _responsabile);
        //            for (int liItem = 0; liItem < listViewAttr.Items[0].SubItems.Count; liItem++)
        //            {
        //                _mailbody += string.Format("\r\n  - {0} - {1}", listViewAttr.Items[0].SubItems[liItem].Name, listViewAttr.Items[0].SubItems[liItem].Text);

        //            }

        //            _mailsubject = string.Format("PDL - {0} - Firme Responsabili Tecnici COMPLETATE", "");
        //            MessageBox.Show(string.Format("Notifica al Capo commessa {0} - {1} ", _capocommessa, _mailaddress));
        //            //SendNotify(_mailaddress, _mailsubject, _mailbody);

        //        }


        //        //if (nrCertificati1F >= nrCertificatiUtente1F)
        //        if ((nrCertificatiUtente1FDaFirmare == 0) && (_tecnico == kw.CurrentUser) && (SignFiles.tipofirma == 0))
        //        {
        //            _mailaddress = kw.GetEmailSubjectByName(_responsabile);

        //            _mailbody = string.Format("Notifica al Responsabile Tecnico {0} - da parte del Tecnico {1}: ", _capocommessa, _responsabile);
        //            for (int liItem = 0; liItem < listViewAttr.Items[0].SubItems.Count; liItem++)
        //            {
        //                _mailbody += string.Format("\r\n  - {0} - {1}", listViewAttr.Items[0].SubItems[liItem].Name, listViewAttr.Items[0].SubItems[liItem].Text);

        //            }

        //            MessageBox.Show(string.Format("Notifica al Responsabile Tecnico {0} - {1}: ", _responsabile, _mailaddress));
        //            //SendNotify(_mailaddress, "PDL pronto per la firma digitale", "Testo: PDL pronto per la firma digitale");
        //        }

        //    }

        //    toolStripStatusLabel1.Text = "";

        }


        public bool NotificaCapocommessa()
        {
            string _tecnico = "";
            string _capocommessa = "";
            string _mailaddress = "";
            string _mailsubject = "";
            string _mailbody = "";            
            
            bool bOK = true;

            //bOK = kw.GetPDL(CurrentIdObject, listViewAttr, dataGridViewCertificati, lvFileFirma, statusStrip1);

            for (int i = 0; i < dataGridViewCertificati.Rows.Count; i++)
            {
                if ((int.Parse(dataGridViewCertificati["IdObject", i].Value.ToString()) == CurrentIdObjectCertificato) || (CurrentIdObjectCertificato == 0))
                {
                    _tecnico = dataGridViewCertificati["Tecnico", i].Value.ToString();
                    //_responsabile = dataGridViewCertificati["ResponsabileTecnico", i].Value.ToString();
                    _capocommessa = dataGridViewCertificati["CapoCommessa", i].Value.ToString();
                    break;
                }
            }


            _mailaddress = kw.GetEmailSubjectByName(_capocommessa);

            _mailbody = string.Format("Notifica al Capo commessa {0} - da parte del Tecnico {1}: ", _capocommessa, _tecnico);
            for (int liItem = 0; liItem < listViewAttr.Items.Count; liItem++)
            {
                _mailbody += string.Format("\r\n  - {0} - {1}", listViewAttr.Items[liItem].Text, listViewAttr.Items[liItem].SubItems[1].Text);

            }

            _mailsubject = string.Format("PDL - {0} - Firme Tecnici COMPLETATE", "");
            //MessageBox.Show(string.Format("Notifica al Capo commessa {0} - {1} ", _capocommessa, _mailaddress));
            SendNotify(_mailaddress, _mailsubject, _mailbody);

            return bOK;

        }

        private void btnSurvey_Click(object sender, EventArgs e)
        {

            if (SurveyExists() == false)
            {
                int IdObjectSurvey = kw.CreateSurvey();


                if (IdObjectSurvey > 0)
                {
                    MessageBox.Show("Survey creato correttamente!", "Creazione Survey", MessageBoxButtons.OK);

                }
            }
            else
            {
                MessageBox.Show("Survey già creato e associato all PDL", "Survey", MessageBoxButtons.OK);
            }
        }

        private bool SurveyExists()
        {
            bool bExists = false;
            
            IKnosResult kr;

            IKnosObject kObj = KnosInstance.Client.CreateKnosObject();
            kr = kObj.GetObjectLinks(CurrentIdObject);
            
            if (kr.NoWarningsErrors)
            {
                for(int l = 0; l < kObj.LinkList.ItemCount; l++)
                {
                    if (kObj.LinkList.GetItem(l).Title.ToUpper().StartsWith("SURVEY"))
                    {
                        bExists = true;
                        break;
                    }
                }
            }
            else
            {}


            return bExists;


        
        }


    }
}

