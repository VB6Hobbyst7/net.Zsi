using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace ToDoNotificheBSC
{
    class Notifica
    {

        public static Logger notifyLogger;

        //public static string _localmailsaved = "";

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

        public static bool SendNotifyMAPI(string _address,
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

                if (popup == true)
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


            try
            {
                

                Dictionary<string, string> cdoconfig = new Dictionary<string, string>();

                for (int i = 0; i < Properties.Settings.Default.CdoConfig.Count; i++)
                {
                    cdoconfig.Add(Properties.Settings.Default.CdoConfig[i].Split('=')[0].ToString(), Properties.Settings.Default.CdoConfig[i].Split('=')[1].ToString());
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
                
                bOK = false;
            }


            //Default return value.
            return bOK;
        }

        public static bool SendNotifyVBS(string _address,
            string _subject,
            string _body,
            List<string> _attachments,
            List<string> _bodyimages,
            bool _web = false,
            string _addressCC = "",
            string _addressBCC = "")
        {
            bool bOK = true;

            string VBStext = "";

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



                VBStext += "\r\n Set emailObj = CreateObject(\"CDO.Message\")";
                VBStext += string.Format("\r\n emailObj.From = \"{0}\"", cdoconfig["SMTPAUTHUser"]);
                VBStext += string.Format("\r\n emailObj.To = \"{0}\"", _address);

                VBStext += string.Format("\r\n emailObj.Subject = \"{0}\"", _subject);
                VBStext += string.Format("\r\n emailObj.HTMLBody  = \"{0}\"", _body.Replace(System.Environment.NewLine, "")); 

                foreach (var a in _attachments)
                {
                    if (System.IO.File.Exists(a.Replace("file://", "")))
                    {
                        VBStext += string.Format("\r\n emailObj.AddAttachment \"{0}\"", a);
                    }
                }
                
                VBStext += string.Format("\r\n Set emailConfig = emailObj.Configuration");

                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpserver\") = \"{0}\"", cdoconfig["SMTPServerName"]);
                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpserverport\") = {0}", cdoconfig["SMTPPort"]);
                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/sendusing\") = 2");
                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate\") = 1");
                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/smtpusessl\") = false");
                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/sendusername\") = \"{0}\"", cdoconfig["SMTPAUTHUser"]);
                VBStext += string.Format("\r\n emailConfig.Fields(\"http://schemas.microsoft.com/cdo/configuration/sendpassword\") = \"{0}\"", cdoconfig["SMTPAUTHPassword"]);
                VBStext += string.Format("\r\n emailConfig.Fields.Update");
                VBStext += string.Format("\r\n emailObj.Send");

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

                //VBStext += string.Format("\r\nMailDoc.sendto = {0}", _addressBCCLotus);
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
                MessageBox.Show("Errore " + ex.Message);
                bOK = false;
            }


            //Default return value.
            return bOK;
        }
    }


    class NotificaCOA : Notifica
    {
        public static Logger notifyLogger;

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


                notifyLogger.LogSomething(cdo.SMTPServerName);
                notifyLogger.LogSomething(cdo.SMTPAUTHUser);
                notifyLogger.LogSomething(cdo.SMTPPort.ToString());
                notifyLogger.LogSomething(cdo.SMTPAUTHUserEmail);

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

                notifyLogger.LogSomething(ex.Message);
                bOK = false;
            }


            //Default return value.
            return bOK;
        }

    }
}
