using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Confezionamento
{
    public partial class Confezionamento : Form
    {
        public Confezionamento()
        {
            InitializeComponent();
        }

        private void Confezionamento_Load(object sender, EventArgs e)
        {
            textBox1.Text = "2015";

        }

        private void btnLoadConfezionamento_Click(object sender, EventArgs e)
        {

            int esercizio = 0;

            if (int.TryParse(textBox1.Text, out esercizio) == false)
            {
                MessageBox.Show("Inserire l'esercizio!");

            }

            radGridView1.DataSource = null;
            radGridView1.Rows.Clear();

            DataTable dt = new DataTable();
            using (SqlConnection cn = new SqlConnection(Properties.Settings.Default.ZSIConnectionString))
            {
                cn.Open();

                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = cn;
                    cmd.CommandText = string.Format("SELECT * FROM ZSI_VISTA_CONFEZIONAMENTO WHERE TIPODOC IN ('OCC', 'OCE') AND ESERCIZIO = {0}", textBox1.Text);

                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }

                }


            }

            radGridView1.DataSource = dt;

            for (int i = 0; i < radGridView1.Columns.Count; i++)
            {
                if (radGridView1.Columns[i].FieldName.StartsWith("HIDE_"))
                {
                    radGridView1.Columns[i].IsVisible = false;
                }
                else
                {
                    radGridView1.Columns[i].BestFit();
                }
            }
        }

        private void radGridView1_CellDoubleClick(object sender, Telerik.WinControls.UI.GridViewCellEventArgs e)
        {
            textBoxRowEdit.Text = string.Format("{0} {1}", radGridView1.Rows[e.RowIndex].Cells["PROGRESSIVO"].Value.ToString()
                , radGridView1.Rows[e.RowIndex].Cells["CODART"].Value.ToString());


            textBox2.Text = radGridView1.Rows[e.RowIndex].Cells["NOTECLI"].Value.ToString();
        }
    }
}
