namespace ZSI_IMPORTLISTINI
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            Telerik.WinControls.UI.GridViewCheckBoxColumn gridViewCheckBoxColumn1 = new Telerik.WinControls.UI.GridViewCheckBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn1 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn2 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn3 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn4 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn5 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn6 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewDateTimeColumn gridViewDateTimeColumn1 = new Telerik.WinControls.UI.GridViewDateTimeColumn();
            Telerik.WinControls.UI.GridViewDateTimeColumn gridViewDateTimeColumn2 = new Telerik.WinControls.UI.GridViewDateTimeColumn();
            Telerik.WinControls.UI.GridViewDecimalColumn gridViewDecimalColumn1 = new Telerik.WinControls.UI.GridViewDecimalColumn();
            Telerik.WinControls.UI.GridViewDecimalColumn gridViewDecimalColumn2 = new Telerik.WinControls.UI.GridViewDecimalColumn();
            Telerik.WinControls.UI.GridViewDecimalColumn gridViewDecimalColumn3 = new Telerik.WinControls.UI.GridViewDecimalColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn7 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn8 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn9 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn10 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewDecimalColumn gridViewDecimalColumn4 = new Telerik.WinControls.UI.GridViewDecimalColumn();
            Telerik.WinControls.UI.ConditionalFormattingObject conditionalFormattingObject1 = new Telerik.WinControls.UI.ConditionalFormattingObject();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn11 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn12 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn13 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition1 = new Telerik.WinControls.UI.TableViewDefinition();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.lblFile = new System.Windows.Forms.Label();
            this.textBoxInputFilePath = new System.Windows.Forms.TextBox();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnPubblica = new System.Windows.Forms.Button();
            this.textBoxOutput = new System.Windows.Forms.TextBox();
            this.dGExcel = new System.Windows.Forms.DataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnVerifica = new System.Windows.Forms.Button();
            this.lblProgress = new System.Windows.Forms.Label();
            this.btnStorico = new System.Windows.Forms.Button();
            this.cmbFoglioDiLavoro = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnVerificaPrezzi = new System.Windows.Forms.Button();
            this.chkEuro = new System.Windows.Forms.CheckBox();
            this.radGridViewPP = new Telerik.WinControls.UI.RadGridView();
            this.btnCaricaPrezzi = new System.Windows.Forms.Button();
            this.btnAggiornaPrezzi = new System.Windows.Forms.Button();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.dGExcel)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radGridViewPP)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGridViewPP.MasterTemplate)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // lblFile
            // 
            this.lblFile.AutoSize = true;
            this.lblFile.Location = new System.Drawing.Point(6, 6);
            this.lblFile.Name = "lblFile";
            this.lblFile.Size = new System.Drawing.Size(23, 13);
            this.lblFile.TabIndex = 21;
            this.lblFile.Text = "File";
            // 
            // textBoxInputFilePath
            // 
            this.textBoxInputFilePath.Location = new System.Drawing.Point(35, 9);
            this.textBoxInputFilePath.Name = "textBoxInputFilePath";
            this.textBoxInputFilePath.Size = new System.Drawing.Size(449, 20);
            this.textBoxInputFilePath.TabIndex = 22;
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(490, 7);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(58, 23);
            this.btnOpenFile.TabIndex = 26;
            this.btnOpenFile.Text = "Cerca";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // btnPubblica
            // 
            this.btnPubblica.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPubblica.Location = new System.Drawing.Point(1058, 693);
            this.btnPubblica.Name = "btnPubblica";
            this.btnPubblica.Size = new System.Drawing.Size(106, 46);
            this.btnPubblica.TabIndex = 23;
            this.btnPubblica.Text = "Aggiorna provvigioni";
            this.btnPubblica.UseVisualStyleBackColor = true;
            this.btnPubblica.Click += new System.EventHandler(this.btnPubblica_Click);
            // 
            // textBoxOutput
            // 
            this.textBoxOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textBoxOutput.Location = new System.Drawing.Point(6, 686);
            this.textBoxOutput.Multiline = true;
            this.textBoxOutput.Name = "textBoxOutput";
            this.textBoxOutput.ReadOnly = true;
            this.textBoxOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxOutput.Size = new System.Drawing.Size(270, 53);
            this.textBoxOutput.TabIndex = 25;
            // 
            // dGExcel
            // 
            this.dGExcel.AllowUserToAddRows = false;
            this.dGExcel.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Silver;
            this.dGExcel.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dGExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dGExcel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dGExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dGExcel.DefaultCellStyle = dataGridViewCellStyle2;
            this.dGExcel.Location = new System.Drawing.Point(9, 36);
            this.dGExcel.Name = "dGExcel";
            this.dGExcel.ReadOnly = true;
            this.dGExcel.Size = new System.Drawing.Size(1155, 603);
            this.dGExcel.TabIndex = 27;
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(9, 645);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(1155, 22);
            this.progressBar1.TabIndex = 24;
            // 
            // btnVerifica
            // 
            this.btnVerifica.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnVerifica.Location = new System.Drawing.Point(946, 693);
            this.btnVerifica.Name = "btnVerifica";
            this.btnVerifica.Size = new System.Drawing.Size(106, 46);
            this.btnVerifica.TabIndex = 28;
            this.btnVerifica.Text = "Verifica";
            this.btnVerifica.UseVisualStyleBackColor = true;
            this.btnVerifica.Click += new System.EventHandler(this.btnVerifica_Click);
            // 
            // lblProgress
            // 
            this.lblProgress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(332, 732);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(0, 13);
            this.lblProgress.TabIndex = 29;
            // 
            // btnStorico
            // 
            this.btnStorico.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStorico.Location = new System.Drawing.Point(834, 693);
            this.btnStorico.Name = "btnStorico";
            this.btnStorico.Size = new System.Drawing.Size(106, 46);
            this.btnStorico.TabIndex = 30;
            this.btnStorico.Text = "Storico";
            this.btnStorico.UseVisualStyleBackColor = true;
            this.btnStorico.Click += new System.EventHandler(this.btnStorico_Click);
            // 
            // cmbFoglioDiLavoro
            // 
            this.cmbFoglioDiLavoro.FormattingEnabled = true;
            this.cmbFoglioDiLavoro.Location = new System.Drawing.Point(745, 9);
            this.cmbFoglioDiLavoro.Name = "cmbFoglioDiLavoro";
            this.cmbFoglioDiLavoro.Size = new System.Drawing.Size(150, 21);
            this.cmbFoglioDiLavoro.TabIndex = 32;
            this.cmbFoglioDiLavoro.SelectedIndexChanged += new System.EventHandler(this.cmbFoglioDiLavoro_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(554, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(193, 13);
            this.label1.TabIndex = 33;
            this.label1.Text = "Fogli di lavoro trovati nel file selezionato";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1185, 771);
            this.tabControl1.TabIndex = 34;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.cmbFoglioDiLavoro);
            this.tabPage1.Controls.Add(this.btnStorico);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnOpenFile);
            this.tabPage1.Controls.Add(this.btnVerifica);
            this.tabPage1.Controls.Add(this.textBoxInputFilePath);
            this.tabPage1.Controls.Add(this.btnPubblica);
            this.tabPage1.Controls.Add(this.lblFile);
            this.tabPage1.Controls.Add(this.textBoxOutput);
            this.tabPage1.Controls.Add(this.dGExcel);
            this.tabPage1.Controls.Add(this.progressBar1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1177, 745);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Listini e provvigioni";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.Silver;
            this.tabPage2.Controls.Add(this.btnVerificaPrezzi);
            this.tabPage2.Controls.Add(this.chkEuro);
            this.tabPage2.Controls.Add(this.radGridViewPP);
            this.tabPage2.Controls.Add(this.btnCaricaPrezzi);
            this.tabPage2.Controls.Add(this.btnAggiornaPrezzi);
            this.tabPage2.Controls.Add(this.progressBar2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1177, 745);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Prezzi particolari";
            // 
            // btnVerificaPrezzi
            // 
            this.btnVerificaPrezzi.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnVerificaPrezzi.Location = new System.Drawing.Point(945, 691);
            this.btnVerificaPrezzi.Name = "btnVerificaPrezzi";
            this.btnVerificaPrezzi.Size = new System.Drawing.Size(106, 46);
            this.btnVerificaPrezzi.TabIndex = 34;
            this.btnVerificaPrezzi.Text = "Verifica";
            this.btnVerificaPrezzi.UseVisualStyleBackColor = true;
            this.btnVerificaPrezzi.Click += new System.EventHandler(this.btnVerificaPrezzi_Click);
            // 
            // chkEuro
            // 
            this.chkEuro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkEuro.AutoSize = true;
            this.chkEuro.Checked = true;
            this.chkEuro.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkEuro.Location = new System.Drawing.Point(150, 691);
            this.chkEuro.Name = "chkEuro";
            this.chkEuro.Size = new System.Drawing.Size(267, 17);
            this.chkEuro.TabIndex = 33;
            this.chkEuro.Text = "Prezzo maggiorazione = prezzo maggiorazione euro";
            this.chkEuro.UseVisualStyleBackColor = true;
            // 
            // radGridViewPP
            // 
            this.radGridViewPP.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.radGridViewPP.Location = new System.Drawing.Point(8, 6);
            // 
            // 
            // 
            this.radGridViewPP.MasterTemplate.AllowAddNewRow = false;
            gridViewCheckBoxColumn1.FieldName = "Modificato";
            gridViewCheckBoxColumn1.HeaderText = "Modificato";
            gridViewCheckBoxColumn1.Name = "Modificato";
            gridViewCheckBoxColumn1.Width = 77;
            gridViewTextBoxColumn1.FieldName = "T";
            gridViewTextBoxColumn1.HeaderText = "T";
            gridViewTextBoxColumn1.Name = "T";
            gridViewTextBoxColumn2.FieldName = "CODCLIFOR";
            gridViewTextBoxColumn2.HeaderText = "COD CLI/FOR";
            gridViewTextBoxColumn2.Name = "CODCLIFOR";
            gridViewTextBoxColumn2.ReadOnly = true;
            gridViewTextBoxColumn3.FieldName = "DSCCONTO";
            gridViewTextBoxColumn3.HeaderText = "RAGIONE SOCIALE";
            gridViewTextBoxColumn3.Name = "DSCCONTO";
            gridViewTextBoxColumn3.ReadOnly = true;
            gridViewTextBoxColumn4.FieldName = "CODART";
            gridViewTextBoxColumn4.HeaderText = "CODART";
            gridViewTextBoxColumn4.Name = "CODART";
            gridViewTextBoxColumn4.ReadOnly = true;
            gridViewTextBoxColumn5.FieldName = "DSCARTICOLO";
            gridViewTextBoxColumn5.HeaderText = "DESCR. ARTICOLO";
            gridViewTextBoxColumn5.Name = "DSCARTICOLO";
            gridViewTextBoxColumn5.ReadOnly = true;
            gridViewTextBoxColumn6.FieldName = "UM";
            gridViewTextBoxColumn6.HeaderText = "UM";
            gridViewTextBoxColumn6.Name = "UM";
            gridViewTextBoxColumn6.ReadOnly = true;
            gridViewDateTimeColumn1.FieldName = "INIZIOVALIDITA";
            gridViewDateTimeColumn1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            gridViewDateTimeColumn1.HeaderText = "INIZIO VALIDITA\'";
            gridViewDateTimeColumn1.Name = "INIZIOVALIDITA";
            gridViewDateTimeColumn2.FieldName = "FINEVALIDITA";
            gridViewDateTimeColumn2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            gridViewDateTimeColumn2.HeaderText = "FINE VALIDITA\'";
            gridViewDateTimeColumn2.Name = "FINEVALIDITA";
            gridViewDecimalColumn1.DecimalPlaces = 3;
            gridViewDecimalColumn1.FieldName = "QTAMINIMA";
            gridViewDecimalColumn1.HeaderText = "QTA\' MINIMA";
            gridViewDecimalColumn1.Name = "QTAMINIMA";
            gridViewDecimalColumn2.DecimalPlaces = 4;
            gridViewDecimalColumn2.FieldName = "PREZZO_MAGG";
            gridViewDecimalColumn2.HeaderText = "PREZZO_MAGG";
            gridViewDecimalColumn2.Name = "PREZZO_MAGG";
            gridViewDecimalColumn3.DecimalPlaces = 4;
            gridViewDecimalColumn3.FieldName = "PREZZO_MAGGEURO";
            gridViewDecimalColumn3.HeaderText = "PREZZO EURO";
            gridViewDecimalColumn3.Name = "PREZZO_MAGGEURO";
            gridViewTextBoxColumn7.FieldName = "NRLISTINO";
            gridViewTextBoxColumn7.HeaderText = "NRLISTINO";
            gridViewTextBoxColumn7.Name = "NRLISTINO";
            gridViewTextBoxColumn7.ReadOnly = true;
            gridViewTextBoxColumn8.FieldName = "DSCLISTINO";
            gridViewTextBoxColumn8.HeaderText = "DSCLISTINO";
            gridViewTextBoxColumn8.Name = "DSCLISTINO";
            gridViewTextBoxColumn8.ReadOnly = true;
            gridViewTextBoxColumn9.FieldName = "COD_IMBALLO";
            gridViewTextBoxColumn9.HeaderText = "COD IMBALLO";
            gridViewTextBoxColumn9.Name = "COD_IMBALLO";
            gridViewTextBoxColumn9.ReadOnly = true;
            gridViewTextBoxColumn10.FieldName = "DSCIMBALLO";
            gridViewTextBoxColumn10.HeaderText = "DSC. IMBALLO";
            gridViewTextBoxColumn10.Name = "DSCIMBALLO";
            gridViewTextBoxColumn10.ReadOnly = true;
            conditionalFormattingObject1.CellBackColor = System.Drawing.Color.Empty;
            conditionalFormattingObject1.CellForeColor = System.Drawing.Color.Empty;
            conditionalFormattingObject1.Name = "NewCondition";
            conditionalFormattingObject1.RowBackColor = System.Drawing.Color.Red;
            conditionalFormattingObject1.RowForeColor = System.Drawing.Color.Empty;
            conditionalFormattingObject1.TValue1 = "0";
            conditionalFormattingObject1.TValue2 = "0";
            gridViewDecimalColumn4.ConditionalFormattingObjectList.Add(conditionalFormattingObject1);
            gridViewDecimalColumn4.FieldName = "QTA_COLLI";
            gridViewDecimalColumn4.HeaderText = "QTA_COLLI";
            gridViewDecimalColumn4.Minimum = new decimal(new int[] {
            0,
            0,
            0,
            0});
            gridViewDecimalColumn4.Name = "QTA_COLLI";
            gridViewTextBoxColumn11.FieldName = "SQLUPDATE";
            gridViewTextBoxColumn11.HeaderText = "SQLUPDATE";
            gridViewTextBoxColumn11.Name = "SQLUPDATE";
            gridViewTextBoxColumn11.ReadOnly = true;
            gridViewTextBoxColumn12.FieldName = "SQLUPDATE_TESTATE";
            gridViewTextBoxColumn12.HeaderText = "SQLUPDATE_TESTATE";
            gridViewTextBoxColumn12.Name = "SQLUPDATE_TESTATE";
            gridViewTextBoxColumn13.FieldName = "NR";
            gridViewTextBoxColumn13.HeaderText = "NR";
            gridViewTextBoxColumn13.Name = "NR";
            this.radGridViewPP.MasterTemplate.Columns.AddRange(new Telerik.WinControls.UI.GridViewDataColumn[] {
            gridViewCheckBoxColumn1,
            gridViewTextBoxColumn1,
            gridViewTextBoxColumn2,
            gridViewTextBoxColumn3,
            gridViewTextBoxColumn4,
            gridViewTextBoxColumn5,
            gridViewTextBoxColumn6,
            gridViewDateTimeColumn1,
            gridViewDateTimeColumn2,
            gridViewDecimalColumn1,
            gridViewDecimalColumn2,
            gridViewDecimalColumn3,
            gridViewTextBoxColumn7,
            gridViewTextBoxColumn8,
            gridViewTextBoxColumn9,
            gridViewTextBoxColumn10,
            gridViewDecimalColumn4,
            gridViewTextBoxColumn11,
            gridViewTextBoxColumn12,
            gridViewTextBoxColumn13});
            this.radGridViewPP.MasterTemplate.EnableFiltering = true;
            this.radGridViewPP.MasterTemplate.ViewDefinition = tableViewDefinition1;
            this.radGridViewPP.Name = "radGridViewPP";
            this.radGridViewPP.Size = new System.Drawing.Size(1155, 651);
            this.radGridViewPP.TabIndex = 32;
            this.radGridViewPP.Text = "radGridView1";
            this.radGridViewPP.CellEndEdit += new Telerik.WinControls.UI.GridViewCellEventHandler(this.radGridViewPP_CellEndEdit);
            // 
            // btnCaricaPrezzi
            // 
            this.btnCaricaPrezzi.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnCaricaPrezzi.Location = new System.Drawing.Point(8, 691);
            this.btnCaricaPrezzi.Name = "btnCaricaPrezzi";
            this.btnCaricaPrezzi.Size = new System.Drawing.Size(106, 46);
            this.btnCaricaPrezzi.TabIndex = 31;
            this.btnCaricaPrezzi.Text = "Carica prezzi";
            this.btnCaricaPrezzi.UseVisualStyleBackColor = true;
            this.btnCaricaPrezzi.Click += new System.EventHandler(this.btnCaricaPrezzi_Click);
            // 
            // btnAggiornaPrezzi
            // 
            this.btnAggiornaPrezzi.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAggiornaPrezzi.Location = new System.Drawing.Point(1057, 691);
            this.btnAggiornaPrezzi.Name = "btnAggiornaPrezzi";
            this.btnAggiornaPrezzi.Size = new System.Drawing.Size(106, 46);
            this.btnAggiornaPrezzi.TabIndex = 28;
            this.btnAggiornaPrezzi.Text = "Aggiorna prezzi";
            this.btnAggiornaPrezzi.UseVisualStyleBackColor = true;
            this.btnAggiornaPrezzi.Click += new System.EventHandler(this.btnAggiornaPrezzi_Click);
            // 
            // progressBar2
            // 
            this.progressBar2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar2.Location = new System.Drawing.Point(8, 663);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(1155, 22);
            this.progressBar2.TabIndex = 29;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1185, 771);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.lblProgress);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Aggiornamento listini e provvigioni";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dGExcel)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radGridViewPP.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radGridViewPP)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Label lblFile;
        private System.Windows.Forms.TextBox textBoxInputFilePath;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnPubblica;
        private System.Windows.Forms.TextBox textBoxOutput;
        private System.Windows.Forms.DataGridView dGExcel;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnVerifica;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Button btnStorico;
        private System.Windows.Forms.ComboBox cmbFoglioDiLavoro;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnCaricaPrezzi;
        private System.Windows.Forms.Button btnAggiornaPrezzi;
        private System.Windows.Forms.ProgressBar progressBar2;
        private Telerik.WinControls.UI.RadGridView radGridViewPP;
        private System.Windows.Forms.CheckBox chkEuro;
        private System.Windows.Forms.Button btnVerificaPrezzi;
    }
}

