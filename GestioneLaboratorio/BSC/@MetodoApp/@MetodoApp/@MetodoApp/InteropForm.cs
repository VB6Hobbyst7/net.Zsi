using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MetodoApp
{

    public partial class Form1 : Form
    {
        public static string UtenteDB = "";
        public static string Ditta = "";
        public static string Action = "";
        public static string Key = "";

        public Form1()
        {
            InitializeComponent();
        }

        dynamic imetodotarget ;

        private void button1_Click(object sender, EventArgs e)
        {
            ExecMetodoAction();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (imetodotarget == null) return;
            DataRow dr = null;
            string[] s = new string[] { "codconto" };
            CollectionWrapper2.MetodoHelper.Selezione(imetodotarget, txtSelezione.Text, s, ref dr);
            if (dr != null)
            {
                label6.Text = dr[0].ToString();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MetodoInterop.startlog();

            txtUtenteDB.Text = UtenteDB;
            txtDitta.Text = Ditta;
            txtAction.Text = Action;
            txtKey.Text = Key;

            if ((UtenteDB != "") && (Ditta != ""))
            {
                // verifico che Metodo sia attivo
                btnCheckMetodo_Click(null, null);

            }

            if ((Action != "") && (Key != ""))
            {
                // eeguo l'azione Metodo
                if (imetodotarget != null)
                {
                    ExecMetodoAction();
                }
            }


        }

        private void btnCheckMetodo_Click(object sender, EventArgs e)
        {
            if (btnCheckMetodo.Text == "Disconnetti")
            {
                imetodotarget = null;
                btnCheckMetodo.Text = "Connetti";
            }
            else
            {
                imetodotarget = MetodoInterop.GetObjecFromRot(txtDitta.Text, txtUtenteDB.Text);
                if (imetodotarget != null)
                {
                    btnCheckMetodo.Text = "Disconnetti";
                }
                else
                {
                    MessageBox.Show("Apri Metodo.... ora è chiuso!");


                }
            }
        }

        private void ExecMetodoAction()
        {
            string url = string.Format("metodo://MENU/{0}/{1}", txtAction.Text, txtKey.Text);

            //imetodotarget.NavigateTo(string.Format("metodo://MENU/{0}/@codice={1}", textBox3.Text, textBox5.Text ));
            //imetodotarget.NavigateTo(string.Format("metodo://MENU/{0}/@progressivo={1}", txtAction.Text, txtKey.Text));
            imetodotarget.NavigateTo(url);
            label7.Text = url;

            if (imetodotarget == null)
            {
                MessageBox.Show("Metodo non attivo!");
            }
            else
            {
                //controllo che la form sia aperta
                CollectionWrapper2.MetodoHelper.WaitForHelpContextID(imetodotarget, 3000);
            }

        }

        private void txtAction_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
