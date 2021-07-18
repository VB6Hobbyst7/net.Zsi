using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ToDoNotificheBSC
{
    public partial class frmWebBrowser : Form
    {
        public string url;

        public frmWebBrowser()
        {
            InitializeComponent();
        }

        private void frmWebBrowser_Load(object sender, EventArgs e)
        {
            try
            {
                webBrowser1.Navigate(url);
            }
            catch (Exception ex)
            {

                textBox1.Text = string.Format("Errore durante l'apertura della pagina {0}\r\n{1}", url, ex.Message);
                textBox1.Visible = true;
            }

        }
    }
}
