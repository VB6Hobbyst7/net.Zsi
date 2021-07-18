using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;



namespace ToDoNotificheBSC
{
    public partial class frmLogin : Form
    {



        string conn = "";

        public frmLogin()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (clsLogin.saveCredenziali(txtbLogin.Text, txtbPWD.Text, Properties.Settings.Default.MetodoConnectionStringUSER))
            {
                this.Close();
            }
            //}
            //if (!Properties.Settings.Default.MetodoConnectionStringUSER.Contains("{0}"))
            //{
                
            //}



            //using (SqlConnection cnUser = new SqlConnection(string.Format(Properties.Settings.Default.MetodoConnectionStringUSER, txtbLogin.Text, txtbPWD.Text)))
            //{
            //    try
            //    {
            //        cnUser.Open();

            //        Properties.Settings.Default.MetodoConnectionStringUSER = string.Format(conn, txtbLogin.Text, txtbPWD.Text);
            //        Properties.Settings.Default.CurrentUser = txtbLogin.Text;
            //        Properties.Settings.Default.Save();

            //        this.Close();

            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Accesso non riuscito!");

            //    }
            //}
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            


            //string sqlStr = Properties.Settings.Default.MetodoConnectionString;

            //string[] p = sqlStr.Split(';');

            //foreach (string px in p)
            //{
            //    string[] c = px.Split('=');

            //    conn += c[0] + "=";
                
            //    if (c[0].ToUpper() == "USER ID")
            //    {
            //        Properties.Settings.Default.CurrentUser = c[1];
            //        conn += "{0}" + ";";
            //    }
            //    else if (c[0].ToUpper() == "PASSWORD")
            //    {
            //        conn += "{1}" + ";";
            //    }

            //    else
            //    {
            //        conn += c[1] + ";";

            //    }


            //}

            //if (!Properties.Settings.Default.MetodoConnectionStringUSER.Contains("{0}"))
            //{
            //    sqlStr = Properties.Settings.Default.MetodoConnectionStringUSER;

                

            //    string[] pU = sqlStr.Split(';');

            //    foreach (string px in pU)
            //    {
            //        string[] c = px.Split('=');

            //        if (c[0].ToUpper() == "USER ID")
            //        {
            //            Properties.Settings.Default.CurrentUser = c[1];
            //        }
                    
            //    }
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
