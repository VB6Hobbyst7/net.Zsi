namespace SignRTFPDF
{
    partial class frmFirma
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmFirma));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtStringSignIMG = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btnSurvey = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.panelBrowser = new System.Windows.Forms.Panel();
            this.btnCloseWebBrowser = new System.Windows.Forms.Button();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.lvCCSost = new System.Windows.Forms.ListView();
            this.cboSmartCards = new System.Windows.Forms.ComboBox();
            this.cboTipoDispositivo = new System.Windows.Forms.ComboBox();
            this.btnDettagliCertificato = new System.Windows.Forms.Button();
            this.cboCertificates = new System.Windows.Forms.ComboBox();
            this.btnGetCertificates = new System.Windows.Forms.Button();
            this.btnPDLStatus = new System.Windows.Forms.Button();
            this.btnSendMail = new System.Windows.Forms.Button();
            this.lvFileFirma = new System.Windows.Forms.ListView();
            this.button2 = new System.Windows.Forms.Button();
            this.btnSettings = new System.Windows.Forms.Button();
            this.btnKnoSLogin = new System.Windows.Forms.Button();
            this.listViewAttr = new System.Windows.Forms.ListView();
            this.txtKnosUrl = new System.Windows.Forms.TextBox();
            this.btnFirmaCapoCommessa = new System.Windows.Forms.Button();
            this.dataGridViewCertificati = new System.Windows.Forms.DataGridView();
            this.label6 = new System.Windows.Forms.Label();
            this.txtIdPDL = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtKnoSPassword = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtKnoSUser = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboSmartCardCert = new System.Windows.Forms.ComboBox();
            this.menuZoom = new System.Windows.Forms.ContextMenu();
            this.miPercent800 = new System.Windows.Forms.MenuItem();
            this.miPercent600 = new System.Windows.Forms.MenuItem();
            this.miPercent400 = new System.Windows.Forms.MenuItem();
            this.miPercent200 = new System.Windows.Forms.MenuItem();
            this.miPercent150 = new System.Windows.Forms.MenuItem();
            this.miPercent100 = new System.Windows.Forms.MenuItem();
            this.miPercent75 = new System.Windows.Forms.MenuItem();
            this.miPercent50 = new System.Windows.Forms.MenuItem();
            this.miPercent25 = new System.Windows.Forms.MenuItem();
            this.miPercent10 = new System.Windows.Forms.MenuItem();
            this.menuItem10 = new System.Windows.Forms.MenuItem();
            this.miBestFit = new System.Windows.Forms.MenuItem();
            this.miFullPage = new System.Windows.Forms.MenuItem();
            this.ilToolbar = new System.Windows.Forms.ImageList(this.components);
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.tabControl1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.panelBrowser.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCertificati)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtStringSignIMG
            // 
            this.txtStringSignIMG.Location = new System.Drawing.Point(16, 574);
            this.txtStringSignIMG.Multiline = true;
            this.txtStringSignIMG.Name = "txtStringSignIMG";
            this.txtStringSignIMG.Size = new System.Drawing.Size(787, 26);
            this.txtStringSignIMG.TabIndex = 10;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(7, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(975, 672);
            this.tabControl1.TabIndex = 15;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.btnSurvey);
            this.tabPage3.Controls.Add(this.label10);
            this.tabPage3.Controls.Add(this.label9);
            this.tabPage3.Controls.Add(this.label8);
            this.tabPage3.Controls.Add(this.panelBrowser);
            this.tabPage3.Controls.Add(this.lvCCSost);
            this.tabPage3.Controls.Add(this.cboSmartCards);
            this.tabPage3.Controls.Add(this.cboTipoDispositivo);
            this.tabPage3.Controls.Add(this.btnDettagliCertificato);
            this.tabPage3.Controls.Add(this.cboCertificates);
            this.tabPage3.Controls.Add(this.btnGetCertificates);
            this.tabPage3.Controls.Add(this.btnPDLStatus);
            this.tabPage3.Controls.Add(this.btnSendMail);
            this.tabPage3.Controls.Add(this.lvFileFirma);
            this.tabPage3.Controls.Add(this.button2);
            this.tabPage3.Controls.Add(this.btnSettings);
            this.tabPage3.Controls.Add(this.btnKnoSLogin);
            this.tabPage3.Controls.Add(this.listViewAttr);
            this.tabPage3.Controls.Add(this.txtKnosUrl);
            this.tabPage3.Controls.Add(this.btnFirmaCapoCommessa);
            this.tabPage3.Controls.Add(this.dataGridViewCertificati);
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.txtIdPDL);
            this.tabPage3.Controls.Add(this.label4);
            this.tabPage3.Controls.Add(this.txtKnoSPassword);
            this.tabPage3.Controls.Add(this.label2);
            this.tabPage3.Controls.Add(this.txtKnoSUser);
            this.tabPage3.Controls.Add(this.label1);
            this.tabPage3.Controls.Add(this.cboSmartCardCert);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(967, 646);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Pubblicazione";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btnSurvey
            // 
            this.btnSurvey.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSurvey.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSurvey.ImageKey = "survey.png";
            this.btnSurvey.ImageList = this.imageList1;
            this.btnSurvey.Location = new System.Drawing.Point(652, 577);
            this.btnSurvey.Name = "btnSurvey";
            this.btnSurvey.Size = new System.Drawing.Size(81, 63);
            this.btnSurvey.TabIndex = 37;
            this.btnSurvey.Text = "Survey";
            this.btnSurvey.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnSurvey.UseVisualStyleBackColor = true;
            this.btnSurvey.Click += new System.EventHandler(this.btnSurvey_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "attributo.png");
            this.imageList1.Images.SetKeyName(1, "document_pdf.png");
            this.imageList1.Images.SetKeyName(2, "signed_pdf.png");
            this.imageList1.Images.SetKeyName(3, "Folder-icon.png");
            this.imageList1.Images.SetKeyName(4, "postage-stamp-icon.png");
            this.imageList1.Images.SetKeyName(5, "signature.png");
            this.imageList1.Images.SetKeyName(6, "stamp.png");
            this.imageList1.Images.SetKeyName(7, "busta bianca2 .jpg");
            this.imageList1.Images.SetKeyName(8, "1381772609_info_orange32.png");
            this.imageList1.Images.SetKeyName(9, "1381772289_agt_reload32.png");
            this.imageList1.Images.SetKeyName(10, "pen32.png");
            this.imageList1.Images.SetKeyName(11, "survey.png");
            // 
            // label10
            // 
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(737, 66);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(130, 13);
            this.label10.TabIndex = 36;
            this.label10.Text = "Capo Commessa Sostituto";
            // 
            // label9
            // 
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(460, 66);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(186, 13);
            this.label9.TabIndex = 35;
            this.label9.Text = "Elenco file firma certificato selezionato";
            this.label9.Visible = false;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(22, 66);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(120, 13);
            this.label8.TabIndex = 34;
            this.label8.Text = "Dettagli Piano di Lavoro";
            // 
            // panelBrowser
            // 
            this.panelBrowser.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelBrowser.BackColor = System.Drawing.Color.LightGray;
            this.panelBrowser.Controls.Add(this.btnCloseWebBrowser);
            this.panelBrowser.Controls.Add(this.webBrowser1);
            this.panelBrowser.Location = new System.Drawing.Point(1, 0);
            this.panelBrowser.Name = "panelBrowser";
            this.panelBrowser.Size = new System.Drawing.Size(10, 635);
            this.panelBrowser.TabIndex = 32;
            this.panelBrowser.Visible = false;
            // 
            // btnCloseWebBrowser
            // 
            this.btnCloseWebBrowser.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCloseWebBrowser.BackColor = System.Drawing.Color.Lime;
            this.btnCloseWebBrowser.Location = new System.Drawing.Point(-132, 588);
            this.btnCloseWebBrowser.Name = "btnCloseWebBrowser";
            this.btnCloseWebBrowser.Size = new System.Drawing.Size(136, 34);
            this.btnCloseWebBrowser.TabIndex = 25;
            this.btnCloseWebBrowser.Text = "Chiudi Browser";
            this.btnCloseWebBrowser.UseVisualStyleBackColor = false;
            this.btnCloseWebBrowser.Click += new System.EventHandler(this.btnCloseWebBrowser_Click);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser1.Location = new System.Drawing.Point(3, 3);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(20, 568);
            this.webBrowser1.TabIndex = 24;
            this.webBrowser1.Visible = false;
            // 
            // lvCCSost
            // 
            this.lvCCSost.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lvCCSost.Location = new System.Drawing.Point(733, 82);
            this.lvCCSost.Name = "lvCCSost";
            this.lvCCSost.Size = new System.Drawing.Size(215, 230);
            this.lvCCSost.TabIndex = 33;
            this.lvCCSost.UseCompatibleStateImageBehavior = false;
            this.lvCCSost.View = System.Windows.Forms.View.List;
            this.lvCCSost.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.lvSost_MouseDoubleClick);
            // 
            // cboSmartCards
            // 
            this.cboSmartCards.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboSmartCards.Font = new System.Drawing.Font("Tahoma", 9F);
            this.cboSmartCards.FormattingEnabled = true;
            this.cboSmartCards.Location = new System.Drawing.Point(303, 598);
            this.cboSmartCards.Name = "cboSmartCards";
            this.cboSmartCards.Size = new System.Drawing.Size(203, 22);
            this.cboSmartCards.TabIndex = 30;
            this.cboSmartCards.TabStop = false;
            // 
            // cboTipoDispositivo
            // 
            this.cboTipoDispositivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboTipoDispositivo.Font = new System.Drawing.Font("Tahoma", 9F);
            this.cboTipoDispositivo.FormattingEnabled = true;
            this.cboTipoDispositivo.Items.AddRange(new object[] {
            "Locali",
            "SmartCard"});
            this.cboTipoDispositivo.Location = new System.Drawing.Point(303, 579);
            this.cboTipoDispositivo.Name = "cboTipoDispositivo";
            this.cboTipoDispositivo.Size = new System.Drawing.Size(203, 22);
            this.cboTipoDispositivo.TabIndex = 29;
            this.cboTipoDispositivo.TabStop = false;
            this.cboTipoDispositivo.SelectedIndexChanged += new System.EventHandler(this.cboTipoDispositivo_SelectedIndexChanged);
            // 
            // btnDettagliCertificato
            // 
            this.btnDettagliCertificato.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDettagliCertificato.Font = new System.Drawing.Font("Tahoma", 9F);
            this.btnDettagliCertificato.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnDettagliCertificato.ImageKey = "1381772609_info_orange32.png";
            this.btnDettagliCertificato.ImageList = this.imageList1;
            this.btnDettagliCertificato.Location = new System.Drawing.Point(512, 577);
            this.btnDettagliCertificato.Name = "btnDettagliCertificato";
            this.btnDettagliCertificato.Size = new System.Drawing.Size(58, 63);
            this.btnDettagliCertificato.TabIndex = 28;
            this.btnDettagliCertificato.Text = "Info";
            this.btnDettagliCertificato.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnDettagliCertificato.UseVisualStyleBackColor = true;
            this.btnDettagliCertificato.Click += new System.EventHandler(this.btnDettagliCertificato_Click);
            // 
            // cboCertificates
            // 
            this.cboCertificates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboCertificates.Font = new System.Drawing.Font("Tahoma", 9F);
            this.cboCertificates.FormattingEnabled = true;
            this.cboCertificates.Location = new System.Drawing.Point(166, 621);
            this.cboCertificates.Name = "cboCertificates";
            this.cboCertificates.Size = new System.Drawing.Size(340, 22);
            this.cboCertificates.TabIndex = 27;
            this.cboCertificates.TabStop = false;
            // 
            // btnGetCertificates
            // 
            this.btnGetCertificates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnGetCertificates.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGetCertificates.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGetCertificates.ImageIndex = 2;
            this.btnGetCertificates.Location = new System.Drawing.Point(165, 577);
            this.btnGetCertificates.Name = "btnGetCertificates";
            this.btnGetCertificates.Size = new System.Drawing.Size(132, 42);
            this.btnGetCertificates.TabIndex = 26;
            this.btnGetCertificates.TabStop = false;
            this.btnGetCertificates.Text = "Carica Certificati";
            this.btnGetCertificates.UseVisualStyleBackColor = true;
            this.btnGetCertificates.Click += new System.EventHandler(this.btnGetCertificates_Click);
            // 
            // btnPDLStatus
            // 
            this.btnPDLStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnPDLStatus.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnPDLStatus.ImageKey = "signed_pdf.png";
            this.btnPDLStatus.ImageList = this.imageList1;
            this.btnPDLStatus.Location = new System.Drawing.Point(17, 575);
            this.btnPDLStatus.Name = "btnPDLStatus";
            this.btnPDLStatus.Size = new System.Drawing.Size(142, 63);
            this.btnPDLStatus.TabIndex = 22;
            this.btnPDLStatus.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnPDLStatus.UseVisualStyleBackColor = true;
            this.btnPDLStatus.Click += new System.EventHandler(this.btnPDLStatus_Click);
            // 
            // btnSendMail
            // 
            this.btnSendMail.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSendMail.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSendMail.ImageKey = "busta bianca2 .jpg";
            this.btnSendMail.ImageList = this.imageList1;
            this.btnSendMail.Location = new System.Drawing.Point(739, 577);
            this.btnSendMail.Name = "btnSendMail";
            this.btnSendMail.Size = new System.Drawing.Size(81, 63);
            this.btnSendMail.TabIndex = 21;
            this.btnSendMail.Text = "Notifica";
            this.btnSendMail.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnSendMail.UseVisualStyleBackColor = true;
            this.btnSendMail.Click += new System.EventHandler(this.btnSendMail_Click);
            // 
            // lvFileFirma
            // 
            this.lvFileFirma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lvFileFirma.Location = new System.Drawing.Point(462, 82);
            this.lvFileFirma.Name = "lvFileFirma";
            this.lvFileFirma.Size = new System.Drawing.Size(265, 230);
            this.lvFileFirma.TabIndex = 20;
            this.lvFileFirma.UseCompatibleStateImageBehavior = false;
            this.lvFileFirma.View = System.Windows.Forms.View.List;
            this.lvFileFirma.Visible = false;
            // 
            // button2
            // 
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.ImageKey = "1381772289_agt_reload32.png";
            this.button2.ImageList = this.imageList1;
            this.button2.Location = new System.Drawing.Point(746, 9);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(93, 37);
            this.button2.TabIndex = 19;
            this.button2.Text = "Ricarica";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_2);
            // 
            // btnSettings
            // 
            this.btnSettings.Location = new System.Drawing.Point(861, 6);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(84, 25);
            this.btnSettings.TabIndex = 18;
            this.btnSettings.Text = "Settings";
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Visible = false;
            this.btnSettings.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // btnKnoSLogin
            // 
            this.btnKnoSLogin.Image = global::SignRTFPDF.Properties.Resources.LOGO_KNOS;
            this.btnKnoSLogin.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnKnoSLogin.Location = new System.Drawing.Point(452, 6);
            this.btnKnoSLogin.Name = "btnKnoSLogin";
            this.btnKnoSLogin.Size = new System.Drawing.Size(80, 57);
            this.btnKnoSLogin.TabIndex = 17;
            this.btnKnoSLogin.Text = "Login";
            this.btnKnoSLogin.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnKnoSLogin.UseVisualStyleBackColor = true;
            this.btnKnoSLogin.Click += new System.EventHandler(this.button2_Click);
            // 
            // listViewAttr
            // 
            this.listViewAttr.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewAttr.Location = new System.Drawing.Point(9, 82);
            this.listViewAttr.Name = "listViewAttr";
            this.listViewAttr.Size = new System.Drawing.Size(718, 230);
            this.listViewAttr.TabIndex = 15;
            this.listViewAttr.UseCompatibleStateImageBehavior = false;
            this.listViewAttr.View = System.Windows.Forms.View.Details;
            // 
            // txtKnosUrl
            // 
            this.txtKnosUrl.Location = new System.Drawing.Point(85, 9);
            this.txtKnosUrl.Name = "txtKnosUrl";
            this.txtKnosUrl.Size = new System.Drawing.Size(361, 20);
            this.txtKnosUrl.TabIndex = 14;
            // 
            // btnFirmaCapoCommessa
            // 
            this.btnFirmaCapoCommessa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFirmaCapoCommessa.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnFirmaCapoCommessa.ImageKey = "signature.png";
            this.btnFirmaCapoCommessa.ImageList = this.imageList1;
            this.btnFirmaCapoCommessa.Location = new System.Drawing.Point(826, 577);
            this.btnFirmaCapoCommessa.Name = "btnFirmaCapoCommessa";
            this.btnFirmaCapoCommessa.Size = new System.Drawing.Size(129, 63);
            this.btnFirmaCapoCommessa.TabIndex = 13;
            this.btnFirmaCapoCommessa.Text = "Firma Capocommessa";
            this.btnFirmaCapoCommessa.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnFirmaCapoCommessa.UseVisualStyleBackColor = true;
            this.btnFirmaCapoCommessa.Click += new System.EventHandler(this.btnFirmaCapoCommessa_Click);
            // 
            // dataGridViewCertificati
            // 
            this.dataGridViewCertificati.AllowUserToAddRows = false;
            this.dataGridViewCertificati.AllowUserToDeleteRows = false;
            this.dataGridViewCertificati.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewCertificati.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridViewCertificati.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridViewCertificati.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewCertificati.Location = new System.Drawing.Point(10, 318);
            this.dataGridViewCertificati.MultiSelect = false;
            this.dataGridViewCertificati.Name = "dataGridViewCertificati";
            this.dataGridViewCertificati.ReadOnly = true;
            this.dataGridViewCertificati.Size = new System.Drawing.Size(938, 241);
            this.dataGridViewCertificati.TabIndex = 13;
            this.dataGridViewCertificati.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewCertificati_CellContentClick);
            this.dataGridViewCertificati.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridViewCertificati_CellMouseDoubleClick);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(13, 17);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(54, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Sito KnoS";
            // 
            // txtIdPDL
            // 
            this.txtIdPDL.Location = new System.Drawing.Point(658, 9);
            this.txtIdPDL.Name = "txtIdPDL";
            this.txtIdPDL.Size = new System.Drawing.Size(82, 20);
            this.txtIdPDL.TabIndex = 11;
            this.txtIdPDL.Leave += new System.EventHandler(this.txtIdPDL_Leave);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(541, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Piano di lavoro";
            // 
            // txtKnoSPassword
            // 
            this.txtKnoSPassword.Location = new System.Drawing.Point(318, 37);
            this.txtKnoSPassword.Name = "txtKnoSPassword";
            this.txtKnoSPassword.PasswordChar = '*';
            this.txtKnoSPassword.Size = new System.Drawing.Size(128, 20);
            this.txtKnoSPassword.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(259, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Password";
            // 
            // txtKnoSUser
            // 
            this.txtKnoSUser.Location = new System.Drawing.Point(85, 37);
            this.txtKnoSUser.Name = "txtKnoSUser";
            this.txtKnoSUser.Size = new System.Drawing.Size(128, 20);
            this.txtKnoSUser.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Utente KnoS";
            // 
            // cboSmartCardCert
            // 
            this.cboSmartCardCert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboSmartCardCert.Font = new System.Drawing.Font("Tahoma", 9F);
            this.cboSmartCardCert.FormattingEnabled = true;
            this.cboSmartCardCert.Location = new System.Drawing.Point(167, 621);
            this.cboSmartCardCert.Name = "cboSmartCardCert";
            this.cboSmartCardCert.Size = new System.Drawing.Size(338, 22);
            this.cboSmartCardCert.TabIndex = 31;
            this.cboSmartCardCert.TabStop = false;
            // 
            // menuZoom
            // 
            this.menuZoom.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.miPercent800,
            this.miPercent600,
            this.miPercent400,
            this.miPercent200,
            this.miPercent150,
            this.miPercent100,
            this.miPercent75,
            this.miPercent50,
            this.miPercent25,
            this.miPercent10,
            this.menuItem10,
            this.miBestFit,
            this.miFullPage});
            // 
            // miPercent800
            // 
            this.miPercent800.Index = 0;
            this.miPercent800.Text = "800%";
            // 
            // miPercent600
            // 
            this.miPercent600.Index = 1;
            this.miPercent600.Text = "600%";
            // 
            // miPercent400
            // 
            this.miPercent400.Index = 2;
            this.miPercent400.Text = "400%";
            // 
            // miPercent200
            // 
            this.miPercent200.Index = 3;
            this.miPercent200.Text = "200%";
            // 
            // miPercent150
            // 
            this.miPercent150.Index = 4;
            this.miPercent150.Text = "150%";
            // 
            // miPercent100
            // 
            this.miPercent100.Index = 5;
            this.miPercent100.Text = "100%";
            // 
            // miPercent75
            // 
            this.miPercent75.Index = 6;
            this.miPercent75.Text = "75%";
            // 
            // miPercent50
            // 
            this.miPercent50.Index = 7;
            this.miPercent50.Text = "50%";
            // 
            // miPercent25
            // 
            this.miPercent25.Index = 8;
            this.miPercent25.Text = "25%";
            // 
            // miPercent10
            // 
            this.miPercent10.Index = 9;
            this.miPercent10.Text = "10%";
            // 
            // menuItem10
            // 
            this.menuItem10.Index = 10;
            this.menuItem10.Text = "-";
            // 
            // miBestFit
            // 
            this.miBestFit.Index = 11;
            this.miBestFit.Text = "Best fit";
            // 
            // miFullPage
            // 
            this.miFullPage.Index = 12;
            this.miFullPage.Text = "Full Page";
            // 
            // ilToolbar
            // 
            this.ilToolbar.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ilToolbar.ImageStream")));
            this.ilToolbar.TransparentColor = System.Drawing.Color.Lime;
            this.ilToolbar.Images.SetKeyName(0, "");
            this.ilToolbar.Images.SetKeyName(1, "");
            this.ilToolbar.Images.SetKeyName(2, "");
            this.ilToolbar.Images.SetKeyName(3, "");
            this.ilToolbar.Images.SetKeyName(4, "");
            this.ilToolbar.Images.SetKeyName(5, "");
            this.ilToolbar.Images.SetKeyName(6, "");
            this.ilToolbar.Images.SetKeyName(7, "");
            this.ilToolbar.Images.SetKeyName(8, "");
            this.ilToolbar.Images.SetKeyName(9, "");
            this.ilToolbar.Images.SetKeyName(10, "");
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 700);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(984, 22);
            this.statusStrip1.TabIndex = 19;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never;
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(150, 16);
            this.toolStripProgressBar1.Visible = false;
            // 
            // frmFirma
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 722);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.txtStringSignIMG);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(1000, 760);
            this.Name = "frmFirma";
            this.Text = "Firma Documenti";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmFirma_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.panelBrowser.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewCertificati)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        //private WinWordControl.WinWordControl winWordControl1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtStringSignIMG;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.ContextMenu menuZoom;
        private System.Windows.Forms.MenuItem miPercent800;
        private System.Windows.Forms.MenuItem miPercent600;
        private System.Windows.Forms.MenuItem miPercent400;
        private System.Windows.Forms.MenuItem miPercent200;
        private System.Windows.Forms.MenuItem miPercent150;
        private System.Windows.Forms.MenuItem miPercent100;
        private System.Windows.Forms.MenuItem miPercent75;
        private System.Windows.Forms.MenuItem miPercent50;
        private System.Windows.Forms.MenuItem miPercent25;
        private System.Windows.Forms.MenuItem miPercent10;
        private System.Windows.Forms.MenuItem menuItem10;
        private System.Windows.Forms.MenuItem miBestFit;
        private System.Windows.Forms.MenuItem miFullPage;
        private System.Windows.Forms.ImageList ilToolbar;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TextBox txtKnoSPassword;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtKnoSUser;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtIdPDL;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnFirmaCapoCommessa;
        private System.Windows.Forms.DataGridView dataGridViewCertificati;
        private System.Windows.Forms.TextBox txtKnosUrl;
        private System.Windows.Forms.ListView listViewAttr;
        private System.Windows.Forms.Button btnKnoSLogin;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ListView lvFileFirma;
        private System.Windows.Forms.Button btnSendMail;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.Button btnPDLStatus;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ComboBox cboTipoDispositivo;
        private System.Windows.Forms.Button btnDettagliCertificato;
        private System.Windows.Forms.ComboBox cboCertificates;
        private System.Windows.Forms.Button btnGetCertificates;
        private System.Windows.Forms.ComboBox cboSmartCards;
        private System.Windows.Forms.ComboBox cboSmartCardCert;
        private System.Windows.Forms.Panel panelBrowser;
        private System.Windows.Forms.Button btnCloseWebBrowser;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ListView lvCCSost;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
        private System.Windows.Forms.Button btnSurvey;
    }
}

