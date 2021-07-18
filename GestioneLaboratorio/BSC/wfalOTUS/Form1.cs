using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace wfalOTUS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SendMail(textBox1.Text, "alfredo.deangelo@gmail.com", "test");






        }

        public void SendMail(string Recipient, string CCD, string Subject)
        {
            dynamic Notes = null;
            object db = null;
            dynamic WorkSpace = null;
            dynamic UIdoc = null;
            dynamic LNHeader = null;
            dynamic LNStream = null;


            string userName = null;
            string MailDbName = null;
            Notes = Activator.CreateInstance(Type.GetTypeFromProgID("Notes.NotesSession"));
            userName = Notes.userName;

            textBoxLOG.Text += "\r\n" + userName;

            MailDbName = userName.Substring(0, 1) + userName.Substring(userName.Length - ((userName.Length - (userName.IndexOf(" ", 0) + 1)))) + ".nsf";
            textBoxLOG.Text += "\r\n" + MailDbName;

            db = Notes.GetDataBase(null, MailDbName);

            LNStream = Notes.CreateStream();
            Notes.ConvertMime = false;



            UIdoc = Notes.CurrentDatabase.CreateDocument();
            UIdoc.ReplaceItemValue("Form", "Memo");

            //Recipient = "test@email.com";
            //CCD = "test2@email.com";

            //UIdoc.FieldSetText("EnterSendTo", Recipient);
            //UIdoc.FieldSetText("EnterCopyTo", CCD);

            Subject = "Subject";
            //UIdoc.FieldSetText("Subject", Subject);


            //UIdoc.GotoField("Body");
            //UIdoc.INSERTTEXT("This text goes to body");
            string _body = "<p>test</p><br><p>aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa</p>";
            // Create the body to hold HTML and attachment
            dynamic body = UIdoc.CreateMIMEEntity();
            

            LNHeader = body.CreateHeader("Subject");
            LNHeader.SetHeaderVal(Subject);

            LNHeader = body.CreateHeader("To");
            LNHeader.SetHeaderVal(Recipient);

            LNStream.WriteText("<html>");
            LNStream.WriteText("<body bgcolor=\"blue\" text=\"white\">");
            LNStream.WriteText("<table border=\"2\">");
            LNStream.WriteText("<tr>");
            LNStream.WriteText("<td>Hello World!</td>");
            LNStream.WriteText("</tr>");
            LNStream.WriteText("</table>");
            LNStream.WriteText("</body>");
            LNStream.WriteText("</html>");
            body.SetContentFromText(LNStream, "text/HTML;charset=UTF-8", 1728);

            UIdoc.SAVEMESSAGEONSEND = true;
            UIdoc.Save(true, false);

            //UIdoc.PostedDate = System.DateTime.Now;
            if (checkBox1.Checked)
            {
                UIdoc.Send(false);
            }
            else
            {
                WorkSpace = Activator.CreateInstance(Type.GetTypeFromProgID("Notes.NotesUIWorkspace"));
                WorkSpace.EditDocument(true, UIdoc);
                //UIdoc = WorkSpace.currentdocument;

            }
            textBoxLOG.Text += "\r\n" + "Inviato";

            Notes.ConvertMime = true;

            UIdoc = null;
            WorkSpace = null;
            db = null;
            Notes = null;
        }
    }
}
