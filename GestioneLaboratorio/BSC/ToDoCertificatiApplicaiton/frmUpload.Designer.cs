namespace ToDoNotificheBSC
{
    partial class frmUpload
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.FILENAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DESCRIZIONEALLEGATO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PATHFILE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FILENAME,
            this.DESCRIZIONEALLEGATO,
            this.PATHFILE});
            this.dataGridView1.Location = new System.Drawing.Point(12, 34);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(660, 272);
            this.dataGridView1.TabIndex = 0;
            // 
            // FILENAME
            // 
            this.FILENAME.Frozen = true;
            this.FILENAME.HeaderText = "Nome File";
            this.FILENAME.Name = "FILENAME";
            this.FILENAME.Width = 250;
            // 
            // DESCRIZIONEALLEGATO
            // 
            this.DESCRIZIONEALLEGATO.Frozen = true;
            this.DESCRIZIONEALLEGATO.HeaderText = "Decrizione Allegato";
            this.DESCRIZIONEALLEGATO.Name = "DESCRIZIONEALLEGATO";
            this.DESCRIZIONEALLEGATO.Width = 250;
            // 
            // PATHFILE
            // 
            this.PATHFILE.HeaderText = "Percorso";
            this.PATHFILE.Name = "PATHFILE";
            this.PATHFILE.ReadOnly = true;
            this.PATHFILE.Width = 50;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(522, 312);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 38);
            this.button1.TabIndex = 1;
            this.button1.Text = "Allega";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(379, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Trascinate su questa tabella i file che intendete allegare alle pubblicazioni Kno" +
    "s";
            this.label1.DoubleClick += new System.EventHandler(this.label1_DoubleClick);
            // 
            // frmUpload
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 361);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.MinimumSize = new System.Drawing.Size(700, 400);
            this.Name = "frmUpload";
            this.Text = "Knos - Allega files";
            this.Load += new System.EventHandler(this.frmUpload_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.frmUpload_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.frmUpload_DragEnter);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn FILENAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn DESCRIZIONEALLEGATO;
        private System.Windows.Forms.DataGridViewTextBoxColumn PATHFILE;
    }
}