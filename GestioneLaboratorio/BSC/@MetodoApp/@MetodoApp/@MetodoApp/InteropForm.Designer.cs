namespace MetodoApp
{
    partial class Form1
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtUtenteDB = new System.Windows.Forms.TextBox();
            this.txtDitta = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtAction = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSelezione = new System.Windows.Forms.TextBox();
            this.btnCheckMetodo = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txtKey = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.lblKey = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtUtenteDB
            // 
            this.txtUtenteDB.Location = new System.Drawing.Point(126, 50);
            this.txtUtenteDB.Name = "txtUtenteDB";
            this.txtUtenteDB.Size = new System.Drawing.Size(187, 20);
            this.txtUtenteDB.TabIndex = 0;
            this.txtUtenteDB.Text = "TRM1";
            // 
            // txtDitta
            // 
            this.txtDitta.Location = new System.Drawing.Point(126, 76);
            this.txtDitta.Name = "txtDitta";
            this.txtDitta.Size = new System.Drawing.Size(187, 20);
            this.txtDitta.TabIndex = 1;
            this.txtDitta.Text = "DEMO_160000";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Login Metodo:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(32, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Ditta Metodo:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 159);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Voce di menu:";
            // 
            // txtAction
            // 
            this.txtAction.Location = new System.Drawing.Point(126, 156);
            this.txtAction.Name = "txtAction";
            this.txtAction.Size = new System.Drawing.Size(187, 20);
            this.txtAction.TabIndex = 4;
            this.txtAction.Text = "AnagraficheMag_1";
            this.txtAction.TextChanged += new System.EventHandler(this.txtAction_TextChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(208, 211);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Apri Voce Menu";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(208, 297);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(105, 23);
            this.button2.TabIndex = 9;
            this.button2.Text = "Apri Selezione";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 274);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Selezione:";
            // 
            // txtSelezione
            // 
            this.txtSelezione.Location = new System.Drawing.Point(126, 271);
            this.txtSelezione.Name = "txtSelezione";
            this.txtSelezione.Size = new System.Drawing.Size(187, 20);
            this.txtSelezione.TabIndex = 7;
            this.txtSelezione.Text = "ANACF";
            // 
            // btnCheckMetodo
            // 
            this.btnCheckMetodo.Location = new System.Drawing.Point(208, 102);
            this.btnCheckMetodo.Name = "btnCheckMetodo";
            this.btnCheckMetodo.Size = new System.Drawing.Size(105, 23);
            this.btnCheckMetodo.TabIndex = 10;
            this.btnCheckMetodo.Text = "Connetti";
            this.btnCheckMetodo.UseVisualStyleBackColor = true;
            this.btnCheckMetodo.Click += new System.EventHandler(this.btnCheckMetodo_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(32, 185);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "ID Anagrafica:";
            // 
            // txtKey
            // 
            this.txtKey.Location = new System.Drawing.Point(126, 182);
            this.txtKey.Name = "txtKey";
            this.txtKey.Size = new System.Drawing.Size(187, 20);
            this.txtKey.TabIndex = 11;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(123, 328);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(0, 13);
            this.label6.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(13, 249);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(0, 13);
            this.label7.TabIndex = 14;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(126, 213);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(77, 20);
            this.textBox6.TabIndex = 15;
            this.textBox6.Text = "MENU";
            // 
            // lblKey
            // 
            this.lblKey.AutoSize = true;
            this.lblKey.Location = new System.Drawing.Point(32, 240);
            this.lblKey.Name = "lblKey";
            this.lblKey.Size = new System.Drawing.Size(0, 13);
            this.lblKey.TabIndex = 16;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(327, 538);
            this.Controls.Add(this.lblKey);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtKey);
            this.Controls.Add(this.btnCheckMetodo);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSelezione);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtAction);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDitta);
            this.Controls.Add(this.txtUtenteDB);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtUtenteDB;
        private System.Windows.Forms.TextBox txtDitta;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtAction;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSelezione;
        private System.Windows.Forms.Button btnCheckMetodo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtKey;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label lblKey;
    }
}

