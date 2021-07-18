using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;


namespace ToDoNotificheBSC
{
    

    public class Notifica
    {
        
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);


        public static bool SendNotifyMAPILotus(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            bool popup = false,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = false;


            try
            {
                //SendFileTo.LotusNotes mapi = new SendFileTo.LotusNotes();


                //Dictionary<string, string> lotusconfig = new Dictionary<string, string>();

                //for (int i = 0; i < Properties.Settings.Default.LotusConfig.Count; i++)
                //{
                //    lotusconfig.Add(Properties.Settings.Default.LotusConfig[i].Split('=')[0].ToString(), Properties.Settings.Default.LotusConfig[i].Split('=')[1].ToString());
                //}



                //mapi.SMTPServerName = lotusconfig["SMTPServerName"];
                //mapi.SMTPAUTHUser = lotusconfig["SMTPAUTHUser"];
                //mapi.SMTPAUTHPassword = lotusconfig["SMTPAUTHPassword"];
                //int.TryParse(lotusconfig["SMTPPort"], out mapi.SMTPPort);
                //mapi.SMTPUseSSL = (lotusconfig["SMTPUseSSL"] == "1");
                //mapi.SMTPAUTHUserEmail = string.Format("data\\{0}", lotusconfig["SMTPAUTHUserEmail"]);
                //mapi.SMTPAlwaysSendAs = lotusconfig["SMTPAlwaysSendAs"];


                //mapi.to = _address;
                //mapi.toCC = _addressCC;
                //mapi.toBCC = _addressBCC;
                //mapi.attachments = _attachments;

                //if (popup == true)
                //{
                //    mapi.SendMailPopup(_subject, _body);
                //}
                //else
                //{
                //    mapi.SendMailPopup(_subject, _body);
                //}


            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Errore nella lettura delle impostazioni \r\n{0}", ex.Message));
                bOK = false;
            }


            //Default return value.
            return bOK;

        }


        public static bool SendNotifyOtlk(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            bool popup = true,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = false;

            Logger log;
            log = new Logger();
            log.Setup();
            log.LogSomething("Start - SendNotifyOtlk");

            try
            {
                SendFileTo.OTLK otlk = new SendFileTo.OTLK();
                otlk.SendMail(_address, _addressCC, _subject, _body, popup);

            }
            catch (Exception ex)
            {
                log.LogSomething(ex.Message);

                bOK = false;
            }


            //Default return value.
            return bOK;

        }
        public static bool SendNotifyMAPI(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            bool popup = true,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = false;

            Logger log;
            log = new Logger();
            log.Setup();
            log.LogSomething("Start - SendNotifyMAPI");


            try
            {
                


                SendFileTo.MAPI mapi = new SendFileTo.MAPI();

                mapi.AddRecipientTo(_address);

                if (!string.IsNullOrEmpty(_addressCC))
                    mapi.AddRecipientCC(_addressCC);

                if (!string.IsNullOrEmpty(_addressBCC))
                    mapi.AddRecipientBCC(_addressBCC);

                for (int i = 0; i < _attachments.Count; i++)
                {
                    if (!string.IsNullOrEmpty(_attachments[i]))
                        mapi.AddAttachment(_attachments[i]);
                }

                if (popup)
                {
                    mapi.SendMailPopup(_subject, _body);
                }
                else
                {
                    mapi.SendMailDirect(_subject, _body);
                }


            }
            catch (Exception ex)
            {
                log.LogSomething(ex.Message);

                bOK = false;
            }


            //Default return value.
            return bOK;

        }

        public virtual bool SendNotifyCdo(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            List<string> _bodyimages,
            bool _web = false,
            string _addressCC = "",
            string _addressBCC = "")
        {

            bool bOK = true;

            Logger log;
            log = new Logger();
            log.Setup();
            log.LogSomething("Start - SendNotifyCdo");


            try
            {
                Dictionary<string, string> cdoconfig = new Dictionary<string, string>();

                for (int i = 0; i < Properties.Settings.Default.CdoConfig.Count; i++)
                {
                    cdoconfig.Add(Properties.Settings.Default.CdoConfig[i].Split('=')[0].ToString(), Properties.Settings.Default.CdoConfig[i].Split('=')[1].ToString());
                }


                SendFileTo.Cdo cdo = new SendFileTo.Cdo();
                cdo.SMTPServerName = cdoconfig["SMTPServerName"];
                cdo.SMTPAUTHUser = cdoconfig["SMTPAUTHUser"];
                cdo.SMTPAUTHPassword = cdoconfig["SMTPAUTHPassword"];
                int.TryParse(cdoconfig["SMTPPort"], out cdo.SMTPPort);
                cdo.SMTPUseSSL = (cdoconfig["SMTPUseSSL"] == "1");
                cdo.SMTPAUTHUserEmail = cdoconfig["SMTPAUTHUserEmail"];
                cdo.SMTPAlwaysSendAs = cdoconfig["SMTPAlwaysSendAs"];

                cdo.to = _address;
                cdo.toCC = _addressCC;
                cdo.toBCC = _addressBCC;
                cdo.attachments = _attachments;
                if (_web)
                {
                    cdo.SendMailWeb(_subject, _body, 0);
                }
                else
                {
                    cdo.SendMail(_subject, _body, 0);
                }


            }
            catch (Exception ex)
            {
                log.LogSomething(ex.Message);
                bOK = false;
            }


            //Default return value.
            return bOK;
        }

        public static bool SendNotifyVBSLotus(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            List<string> _bodyimages,
            bool _web = false,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = true;

            Logger log;
            log = new Logger();
            log.Setup();
            log.LogSomething("Start - SendNotifyVBSLotus");

            string VBStext = "";

            try
            {
                Dictionary<string, string> cdoconfig = new Dictionary<string, string>();

                for (int i = 0; i < Properties.Settings.Default.CdoConfig.Count; i++)
                {
                    cdoconfig.Add(Properties.Settings.Default.CdoConfig[i].Split('=')[0].ToString(), Properties.Settings.Default.CdoConfig[i].Split('=')[1].ToString());
                }

                string _addressLotus = string.Format("Array(\"{0}\")", _address.Replace(";", "\",\""));
                string _addressCCLotus = string.Format("Array(\"{0}\")", _addressCC.Replace(";", "\",\""));
                string _addressBCCLotus = string.Format("Array(\"{0}\")", _addressBCC.Replace(";", "\",\""));

                VBStext += "\r\nDim Maildb  'The mail database";
                VBStext += "\r\nDim UserName 'The current users notes name";
                VBStext += "\r\nDim MailDbName 'THe current users notes mail database name";
                VBStext += "\r\nDim MailDoc  'The mail document itself";
                VBStext += "\r\nDim AttachME  'The attachment richtextfile object";
                VBStext += "\r\nDim Session  'The notes session";
                VBStext += "\r\nDim EmbedObj  'The embedded object (Attachment)";
                VBStext += "\r\nSet Session = CreateObject(\"Notes.NotesSession\")";




                VBStext += "\r\nif Session is nothing then";
                VBStext += "\r\n'msgbox \"non inizializzato\"";
                VBStext += "\r\nend if";

                VBStext += "\r\n'Session.Initialize(\"Tuttoesaurito0\")";
                VBStext += "\r\n'msgbox \"inizializzata\"";
                VBStext += "\r\nUserName = Session.UserName";
                VBStext += "\r\n'msgbox UserName";
                VBStext += "\r\nMailDbName = Left(UserName, 1) & Right(UserName, (Len(UserName) - InStr(1, UserName, \" \"))) & \".nsf\"";
                VBStext += "\r\n'msgbox MailDbName";

                VBStext += "\r\n'Open the mail database in notes";
                VBStext += "\r\nSet Maildb = Session.GETDATABASE(\"\", MailDbName)";
                VBStext += "\r\nIf Maildb.ISOPEN = True Then";
                VBStext += "\r\n'msgbox \"'Already open for mail\"";
                VBStext += "\r\nElse";
                VBStext += "\r\nMaildb.OPENMAIL";
                VBStext += "\r\nEnd If";

                VBStext += "\r\n'msgbox \"'Set up the new mail document\"";
                VBStext += "\r\nSet MailDoc = Maildb.CREATEDOCUMENT";
                VBStext += "\r\nMailDoc.Form = \"Memo\"";
                VBStext += string.Format("\r\nMailDoc.sendto = {0}", _addressLotus);
                VBStext += string.Format("\r\nMailDoc.copyto = {0}", _addressCCLotus);

                if (_addressBCC != "")
                {
                    VBStext += "\r\nDim objNotesField";
                    VBStext += string.Format("\r\nSet objNotesField = MailDoc.APPENDITEMVALUE(\"BlindCopyTo\", {0})", _addressBCCLotus);
                }

                VBStext += string.Format("\r\nMailDoc.sendto = {0}", _addressBCCLotus);
                VBStext += string.Format("\r\nMailDoc.Subject = \"{0}\"", _subject);
                VBStext += string.Format("\r\nMailDoc.Body = {0}", _body);
                VBStext += "\r\nMailDoc.SAVEMESSAGEONSEND = true";
                //VBStext += "\r\n'Set up the embedded object and attachment and attach it";
                //VBStext += "\r\nIf Attachment <> \"\" Then";
                //VBStext += "\r\nSet AttachME = MailDoc.CREATERICHTEXTITEM(\"Attachment\")";
                //VBStext += "\r\nSet EmbedObj = AttachME.EMBEDOBJECT(1454, \"\", Attachment, \"Attachment\")";
                //VBStext += "\r\nMailDoc.CREATERICHTEXTITEM(\"Attachment\")";
                //VBStext += "\r\nEnd If";

                int iA = 0;

                foreach (var a in _attachments)
                {

                    if (System.IO.File.Exists(a.Replace("file://", "")))
                    {
                        iA += 1;

                        //VBStext += "\r\nIf Attachment <> \"\" Then";
                        VBStext += string.Format("\r\nDim AttachME{0}", iA.ToString());
                        VBStext += string.Format("\r\nDim EmbedObj{0}", iA.ToString());
                        VBStext += string.Format("\r\nSet AttachME{0} = MailDoc.CREATERICHTEXTITEM(\"Attachment{0}\")", iA.ToString());
                        VBStext += string.Format("\r\nSet EmbedObj{0} = AttachME{0}.EMBEDOBJECT(1454, \"\", \"{1}\", \"Attachment{0}\")", iA.ToString(), a);
                        //VBStext += "\r\nMailDoc.CREATERICHTEXTITEM(\"Attachment\")";
                        //VBStext += "\r\nEnd If";
                    }
                }

                VBStext += "\r\n'msgbox \"'Send the document\"";
                VBStext += "\r\nMailDoc.PostedDate = Now() 'Gets the mail to appear in the sent items folder";
                VBStext += "\r\nMailDoc.SEND 0, Recipient";

                VBStext += "\r\n'msgbox \" 'Clean Up\"";
                VBStext += "\r\n   Set Maildb = Nothing";
                VBStext += "\r\nSet MailDoc = Nothing";
                VBStext += "\r\nSet AttachME = Nothing";
                VBStext += "\r\nSet Session = Nothing";
                VBStext += "\r\nSet EmbedObj = Nothing";

                //VBStext += "\r\n Set emailObj = CreateObject(\"CDO.Message\")";
                //VBStext += string.Format("\r\n emailObj.From = \"{0}\"", cdoconfig["SMTPAUTHUser"]);
                //VBStext += string.Format("\r\n emailObj.To = \"{0}\"", _address);

                //VBStext += string.Format("\r\n emailObj.Subject = \"{0}\"", _subject);
                //VBStext += string.Format("\r\n emailObj.HTMLBody  = \"{0}\"", _body.Replace(System.Environment.NewLine, ""));

                //foreach (var a in _attachments)
                //{
                //    if (System.IO.File.Exists(a.Replace("file://", "")))
                //    {
                //        VBStext += string.Format("\r\n emailObj.AddAttachment \"{0}\"", a);
                //    }
                //}

                //VBStext += string.Format("\r\n Set emailConfig = emailObj.Configuration");

                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpserver\") = \"{0}\"", cdoconfig["SMTPServerName"]);
                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpserverport\") = {0}", cdoconfig["SMTPPort"]);
                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/sendusing\") = 2");
                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate\") = 1");
                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpusessl\") = false");
                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/sendusername\") = \"{0}\"", cdoconfig["SMTPAUTHUser"]);
                //VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/sendpassword\") = \"{0}\"", cdoconfig["SMTPAUTHPassword"]);
                //VBStext += string.Format("\r\n emailConfig.Fields.Update");
                //VBStext += string.Format("\r\n emailObj.Send");

                string VBSfile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "test.vbs");

                //System.IO.StreamWriter tw  = new System.IO.StreamWriter(VBSfile);
                //tw.Write(VBStext);

                System.IO.File.Delete(VBSfile);
                System.IO.File.WriteAllText(VBSfile, VBStext);

                System.Diagnostics.Process scriptProc = new System.Diagnostics.Process();
                scriptProc.StartInfo.FileName = @VBSfile;
                //scriptProc.StartInfo.Arguments = "//B //Nologo {0}";
                scriptProc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                scriptProc.Start();
                scriptProc.WaitForExit();
                scriptProc.Close();


            }
            catch (Exception ex)
            {
                //MessageBox.Show("Errore " + ex.Message);
                log.LogSomething(ex.Message);
                bOK = false;
            }


            //Default return value.
            return bOK;
        }

        public virtual bool SendNotifyLotus(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            List<string> _bodyimages,
            bool _web = false,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = true;

            Logger log;
            log = new Logger();
            log.Setup();
            log.LogSomething("Start - SendNotifyLotus");

            string VBStext = "";

            try
            {
                dynamic Notes = null;
                object db = null;
                dynamic WorkSpace = null;
                dynamic UIdoc = null;
                string userName = null;
                string MailDbName = null;

                Notes = Activator.CreateInstance(Type.GetTypeFromProgID("Notes.NotesSession"));
                userName = Notes.userName;

                MailDbName = userName.Substring(0, 1) + userName.Substring(userName.Length - ((userName.Length - (userName.IndexOf(" ", 0) + 1)))) + ".nsf";

                #region textmail

                ////textBoxLOG.Text += "\r\n" + userName;

                //MailDbName = userName.Substring(0, 1) + userName.Substring(userName.Length - ((userName.Length - (userName.IndexOf(" ", 0) + 1)))) + ".nsf";
                ////textBoxLOG.Text += "\r\n" + MailDbName;

                //db = Notes.GetDataBase(null, MailDbName);
                //WorkSpace = Activator.CreateInstance(Type.GetTypeFromProgID("Notes.NotesUIWorkspace"));
                //WorkSpace.ComposeDocument("", "", "Memo");
                //UIdoc = WorkSpace.currentdocument;
                ////Recipient = "test@email.com";
                ////CCD = "test2@email.com";
                //UIdoc.FieldSetText("EnterSendTo", _address);
                //UIdoc.FieldSetText("EnterCopyTo", _addressCC);
                ////Subject = "Subject";
                //UIdoc.FieldSetText("Subject", _subject);
                //UIdoc.GotoField("Body");
                //UIdoc.INSERTTEXT(_body);

                //UIdoc.Save();
                ////UIdoc.PostedDate = System.DateTime.Now;






                ////Create the notes document 
                //_notesDocument = _notesDataBase.CreateDocument();

                ////Set document type 
                //_notesDocument.ReplaceItemValue(
                //    "Form", "Memo");

                ////sent notes memo fields (To: CC: Bcc: Subject etc) 
                //_notesDocument.ReplaceItemValue(
                //    "SendTo", sSendTo);
                //_notesDocument.ReplaceItemValue(
                //    "CopyTo", sCopyTo);
                //_notesDocument.ReplaceItemValue(
                //    "Subject", sSubject);

                ////Set the body of the email. This allows you to use the appendtext 
                //NotesRichTextItem _richTextItem = _notesDocument.CreateRichTextItem("Body");

                ////add lines to memo email body. the \r\n is needed for each new line. 
                //_richTextItem.AppendText(
                //    "Error: " + errMessage + "\r\n");
                //_richTextItem.AppendText(
                //    "File: " + filename + "\r\n");
                //_richTextItem.AppendText(
                //    "Resolution: " + resolution + "\r\n");
                ////send email & pass in byRef field, this case SendTo (always have this, 





                #endregion


                dynamic LNHeader = null;
                dynamic LNStream = null;

                db = Notes.GetDataBase(null, MailDbName, false);

                LNStream = Notes.CreateStream();
                Notes.ConvertMime = false;

                UIdoc = Notes.CurrentDatabase.CreateDocument();
                UIdoc.ReplaceItemValue("Form", "Memo");

                dynamic body = UIdoc.CreateMIMEEntity();

                

                LNHeader = body.CreateHeader("Subject");
                LNHeader.SetHeaderVal(_subject);

                LNHeader = body.CreateHeader("To");
                LNHeader.SetHeaderVal(_address);

                LNStream.WriteText(_body);
                body.SetContentFromText(LNStream, "text/HTML;charset=UTF-8", 1728);



                //// instantiate a Notes session and workspace
                //Type NotesSession = Type.GetTypeFromProgID("Notes.NotesSession");
                //Type NotesUIWorkspace = Type.GetTypeFromProgID("Notes.NotesUIWorkspace");
                //Object sess = Activator.CreateInstance(NotesSession);
                //Object ws = Activator.CreateInstance(NotesUIWorkspace);

                //// open current user's mail file
                //String mailServer = (String)NotesSession.InvokeMember("GetEnvironmentString", BindingFlags.InvokeMethod, null, sess, new Object[] { "MailServer", true });
                //String mailFile = (String)NotesSession.InvokeMember("GetEnvironmentString", BindingFlags.InvokeMethod, null, sess, new Object[] { "MailFile", true });
                //NotesUIWorkspace.InvokeMember("OpenDatabase", BindingFlags.InvokeMethod, null, ws, new Object[] { mailServer, mailFile });
                //Object uidb = NotesUIWorkspace.InvokeMember("GetCurrentDatabase", BindingFlags.InvokeMethod, null, ws, null);
                //Object db = NotesUIWorkspace.InvokeMember("Database", BindingFlags.GetProperty, null, uidb, null);
                //Type NotesDatabase = db.GetType();

                //// compose a new memo
                //Object uidoc = NotesUIWorkspace.InvokeMember("ComposeDocument", BindingFlags.InvokeMethod, null, ws, new Object[] { mailServer, mailFile, "Memo", 0, 0, true });
                //Type NotesUIDocument = uidoc.GetType();
                //NotesUIDocument.InvokeMember("FieldSetText", BindingFlags.InvokeMethod, null, uidoc, new Object[] { "EnterSendTo", _address });
                //NotesUIDocument.InvokeMember("FieldSetText", BindingFlags.InvokeMethod, null, uidoc, new Object[] { "Subject", _subject });
                //NotesUIDocument.InvokeMember("FieldSetText", BindingFlags.InvokeMethod, null, uidoc, new Object[] { "Body", _body });




                System.Collections.Specialized.StringCollection paths = new System.Collections.Specialized.StringCollection();

                if (_attachments.Count > 0)
                {

                    for (int i = 0; i < _attachments.Count; i++)
                    {
                        //MessageBox.Show("allegato " + _attachments[i]);
                        paths.Add(_attachments[i]);


                        try
                        {
                            dynamic LNHeaderA = null;
                            dynamic LNStreamA = null;
                            dynamic bodychild = body.CreateChildEntity();

                            LNHeaderA = bodychild.CreateHeader("Content-Type");
                            LNHeaderA.SetHeaderVal("multipart/mixed");

                            System.IO.FileInfo fi = new System.IO.FileInfo(_attachments[i]);


                            LNHeaderA = bodychild.CreateHeader("Content-Disposition");
                            LNHeaderA.SetHeaderVal(string.Format("attachment; filename={0}", fi.Name));

                            LNStreamA = Notes.CreateStream();
                            LNStreamA.Open(_attachments[i]);
                            bodychild.SetContentFromBytes(LNStreamA, "application/pdf", 1730);

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Errore in allegati " + ex.Message);
                            log.LogSomething(ex.Message);

                        }

                    }


                }


                UIdoc.SAVEMESSAGEONSEND = true;



                if (_web)
                {
                    UIdoc.Save(true, false);
                    WorkSpace = Activator.CreateInstance(Type.GetTypeFromProgID("Notes.NotesUIWorkspace"));
                    WorkSpace.EditDocument(true, UIdoc);
                    
                }
                    else
                    {
                    UIdoc.Send(false);
                }


                UIdoc = null;
                WorkSpace = null;
                db = null;
                Notes = null;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore " + ex.Message);
                log.LogSomething(ex.Message);
                bOK = false;
            }


            //Default return value.
            return bOK;
        }

        
    }


    class NotificaCOA : Notifica
    {
        public override bool SendNotifyCdo(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            List<string> _bodyimages,
            bool _web = false,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = true;

            Logger log;
            log = new Logger();
            log.Setup();
            log.LogSomething("Start - SendNotifyCdo");

            try
            {


                Dictionary<string, string> cdoconfig = new Dictionary<string, string>();

                for (int i = 0; i < Properties.Settings.Default.CdoConfigCOA.Count; i++)
                {
                    cdoconfig.Add(Properties.Settings.Default.CdoConfigCOA[i].Split('=')[0].ToString(), Properties.Settings.Default.CdoConfig[i].Split('=')[1].ToString());
                }


                SendFileTo.Cdo cdo = new SendFileTo.Cdo();
                //cdo.localmailsaved = _localmailsaved;
                cdo.SMTPServerName = cdoconfig["SMTPServerName"];
                cdo.SMTPAUTHUser = cdoconfig["SMTPAUTHUser"];
                cdo.SMTPAUTHPassword = cdoconfig["SMTPAUTHPassword"];
                int.TryParse(cdoconfig["SMTPPort"], out cdo.SMTPPort);
                cdo.SMTPUseSSL = (cdoconfig["SMTPUseSSL"] == "1");
                cdo.SMTPAUTHUserEmail = cdoconfig["SMTPAUTHUserEmail"];
                cdo.SMTPAlwaysSendAs = cdoconfig["SMTPAlwaysSendAs"];

                cdo.to = _address.Replace(";", ",");
                cdo.toCC = _addressCC.Replace(";", ",");
                cdo.toBCC = _addressBCC.Replace(";", ",");
                cdo.attachments = _attachments;

                //if (_localmailsaved != "")
                //{
                //    cdo.SendSaveEML(_subject, _body, 0);
                //    System.Diagnostics.Process.Start(_localmailsaved);
                //}
                //else
                //{
                if (_web)
                {
                    cdo.SendMailWeb(_subject, _body, 0);
                }
                else
                {

                    cdo.SendMail(_subject, _body, 0);

                }
            }


            //}
            catch (Exception ex)
            {
                log.LogSomething(ex.Message);
                bOK = false;
            }


            //Default return value.
            return bOK;
        }

    }
}
