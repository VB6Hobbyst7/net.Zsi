using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StartProgramXSi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            System.Security.SecureString ssPwd = new System.Security.SecureString();
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.FileName = System.IO.Path.Combine(Application.StartupPath, "ToDoNOtifichebsc.exe");
            //proc.StartInfo.Arguments = "args...";
            proc.StartInfo.Domain = "ZSI";
            proc.StartInfo.UserName = "Administrator";
            string password = "Sept2010&Army";
            for (int x = 0; x < password.Length; x++)
            {
                ssPwd.AppendChar(password[x]);
            }
            proc.StartInfo.Password = ssPwd;
            proc.Start();
        }
    }
}
