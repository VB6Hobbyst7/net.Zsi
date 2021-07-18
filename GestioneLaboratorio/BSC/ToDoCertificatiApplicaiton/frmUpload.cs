using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ToDoNotificheBSC
{
    public partial class frmUpload : Form
    {
        public static List<Allegato> allegati = new List<Allegato>();

        public frmUpload()
        {
            InitializeComponent();
        }

        private void frmUpload_DragDrop(object sender, DragEventArgs e)
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

        private void frmUpload_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void frmUpload_Load(object sender, EventArgs e)
        {
            this.AllowDrop = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
            

            // 
            this.Close();

        }

        private void label1_DoubleClick(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", @"H:\COMUNICAZIONI");
        }
    }
}
