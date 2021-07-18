namespace ToDoNotificheBSC
{
    partial class ItaCalendarObject
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Telerik.WinControls.UI.AppointmentMappingInfo appointmentMappingInfo1 = new Telerik.WinControls.UI.AppointmentMappingInfo();
            Telerik.WinControls.UI.ResourceMappingInfo resourceMappingInfo1 = new Telerik.WinControls.UI.ResourceMappingInfo();
            Telerik.WinControls.UI.SchedulerDailyPrintStyle schedulerDailyPrintStyle1 = new Telerik.WinControls.UI.SchedulerDailyPrintStyle();
            Telerik.WinControls.UI.ListViewDataItem listViewDataItem1 = new Telerik.WinControls.UI.ListViewDataItem("CISTERNA");
            Telerik.WinControls.UI.ListViewDataItem listViewDataItem2 = new Telerik.WinControls.UI.ListViewDataItem("NON CISTERNA");
            Telerik.WinControls.UI.ListViewDataItem listViewDataItem3 = new Telerik.WinControls.UI.ListViewDataItem("ITALIA");
            Telerik.WinControls.UI.ListViewDataItem listViewDataItem4 = new Telerik.WinControls.UI.ListViewDataItem("ESTERO");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ItaCalendarObject));
            Telerik.WinControls.UI.RadPrintWatermark radPrintWatermark1 = new Telerik.WinControls.UI.RadPrintWatermark();
            Telerik.WinControls.UI.RadPrintWatermark radPrintWatermark2 = new Telerik.WinControls.UI.RadPrintWatermark();
            this.schedulerBindingDataSource1 = new Telerik.WinControls.UI.SchedulerBindingDataSource();
            this.zSIDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ZSIDataSet = new ToDoNotificheBSC.ZSIDataSet();
            this.radScheduler1 = new Telerik.WinControls.UI.RadScheduler();
            this.chkSearch = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.radCheckedListBox1 = new Telerik.WinControls.UI.RadCheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.resourcesBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.radCalendar1 = new Telerik.WinControls.UI.RadCalendar();
            this.radSchedulerNavigator1 = new Telerik.WinControls.UI.RadSchedulerNavigator();
            this.btnCOA = new System.Windows.Forms.Button();
            this.btnPrint = new System.Windows.Forms.Button();
            this.rulerTO = new System.Windows.Forms.NumericUpDown();
            this.rulerFROM = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.imageList16 = new System.Windows.Forms.ImageList(this.components);
            this.resourcesBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.appointmentsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.appointmentsTableAdapter = new ToDoNotificheBSC.ZSIDataSetTableAdapters.AppointmentsTableAdapter();
            this.zsiDataSet1 = new ToDoNotificheBSC.ZSIDataSet();
            this.resourcesTableAdapter = new ToDoNotificheBSC.ZSIDataSetTableAdapters.ResourcesTableAdapter();
            this.radPrintDocument1 = new Telerik.WinControls.UI.RadPrintDocument();
            this.radPrintDocument2 = new Telerik.WinControls.UI.RadPrintDocument();
            this.chkPrevisionali = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1.EventProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1.ResourceProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.zSIDataSetBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ZSIDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radScheduler1)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radCheckedListBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesBindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radCalendar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radSchedulerNavigator1)).BeginInit();
            this.radSchedulerNavigator1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rulerTO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rulerFROM)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.appointmentsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.zsiDataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // schedulerBindingDataSource1
            // 
            // 
            // 
            // 
            this.schedulerBindingDataSource1.EventProvider.AllowNew = true;
            this.schedulerBindingDataSource1.EventProvider.DataMember = "Appointments";
            this.schedulerBindingDataSource1.EventProvider.DataSource = this.zSIDataSetBindingSource;
            appointmentMappingInfo1.BackgroundId = "BackgroundId";
            appointmentMappingInfo1.Description = "Description";
            appointmentMappingInfo1.End = "End";
            appointmentMappingInfo1.Location = "Location";
            appointmentMappingInfo1.ResourceId = "ResourceID";
            appointmentMappingInfo1.Start = "Start";
            appointmentMappingInfo1.Summary = "Summary";
            this.schedulerBindingDataSource1.EventProvider.Mapping = appointmentMappingInfo1;
            // 
            // 
            // 
            this.schedulerBindingDataSource1.ResourceProvider.DataMember = "Resources";
            this.schedulerBindingDataSource1.ResourceProvider.DataSource = this.zSIDataSetBindingSource;
            resourceMappingInfo1.Id = "ID";
            resourceMappingInfo1.Image = "Image";
            resourceMappingInfo1.Name = "Name";
            this.schedulerBindingDataSource1.ResourceProvider.Mapping = resourceMappingInfo1;
            // 
            // zSIDataSetBindingSource
            // 
            this.zSIDataSetBindingSource.DataSource = this.ZSIDataSet;
            this.zSIDataSetBindingSource.Position = 0;
            this.zSIDataSetBindingSource.CurrentChanged += new System.EventHandler(this.zSIDataSetBindingSource_CurrentChanged);
            // 
            // ZSIDataSet
            // 
            this.ZSIDataSet.DataSetName = "ZSIDataSet";
            this.ZSIDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // radScheduler1
            // 
            this.radScheduler1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.radScheduler1.AppointmentTitleFormat = "{2} {3}";
            this.radScheduler1.Culture = new System.Globalization.CultureInfo("it-IT");
            this.radScheduler1.DataSource = this.schedulerBindingDataSource1;
            this.radScheduler1.Location = new System.Drawing.Point(2, 76);
            this.radScheduler1.Name = "radScheduler1";
            schedulerDailyPrintStyle1.AppointmentFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            schedulerDailyPrintStyle1.DateEndRange = new System.DateTime(2016, 12, 3, 0, 0, 0, 0);
            schedulerDailyPrintStyle1.DateHeadingFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            schedulerDailyPrintStyle1.DateStartRange = new System.DateTime(2016, 11, 28, 0, 0, 0, 0);
            schedulerDailyPrintStyle1.PageHeadingFont = new System.Drawing.Font("Microsoft Sans Serif", 22F, System.Drawing.FontStyle.Bold);
            this.radScheduler1.PrintStyle = schedulerDailyPrintStyle1;
            this.radScheduler1.Size = new System.Drawing.Size(686, 478);
            this.radScheduler1.TabIndex = 0;
            this.radScheduler1.Text = "radScheduler1";
            this.radScheduler1.AppointmentAdded += new System.EventHandler<Telerik.WinControls.UI.AppointmentAddedEventArgs>(this.radScheduler1_AppointmentAdded);
            ((Telerik.WinControls.UI.DayViewAppointmentsTable)(this.radScheduler1.GetChildAt(0).GetChildAt(0).GetChildAt(2).GetChildAt(0).GetChildAt(2).GetChildAt(3).GetChildAt(0))).Font = new System.Drawing.Font("Segoe UI", 7F);
            // 
            // chkSearch
            // 
            this.chkSearch.AutoSize = true;
            this.chkSearch.Checked = true;
            this.chkSearch.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSearch.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSearch.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkSearch.ImageKey = "filter-add-icon.png";
            this.chkSearch.Location = new System.Drawing.Point(3, 3);
            this.chkSearch.Name = "chkSearch";
            this.chkSearch.Size = new System.Drawing.Size(133, 17);
            this.chkSearch.TabIndex = 3;
            this.chkSearch.Text = "Attiva/Disattiva Filtro";
            this.chkSearch.UseVisualStyleBackColor = true;
            this.chkSearch.CheckedChanged += new System.EventHandler(this.chkSearch_CheckedChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtSearch);
            this.panel1.Controls.Add(this.radCheckedListBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(0, 26);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(168, 148);
            this.panel1.TabIndex = 4;
            // 
            // txtSearch
            // 
            this.txtSearch.Location = new System.Drawing.Point(7, 21);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(148, 18);
            this.txtSearch.TabIndex = 1;
            this.txtSearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSearch_KeyDown);
            // 
            // radCheckedListBox1
            // 
            listViewDataItem1.CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
            listViewDataItem1.Text = "CISTERNA";
            listViewDataItem2.CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
            listViewDataItem2.Text = "NON CISTERNA";
            listViewDataItem3.BackColor = System.Drawing.SystemColors.ActiveCaption;
            listViewDataItem3.CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
            listViewDataItem3.Text = "ITALIA";
            listViewDataItem4.BackColor = System.Drawing.Color.GreenYellow;
            listViewDataItem4.CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
            listViewDataItem4.Text = "ESTERO";
            this.radCheckedListBox1.Items.AddRange(new Telerik.WinControls.UI.ListViewDataItem[] {
            listViewDataItem1,
            listViewDataItem2,
            listViewDataItem3,
            listViewDataItem4});
            this.radCheckedListBox1.Location = new System.Drawing.Point(7, 41);
            this.radCheckedListBox1.Name = "radCheckedListBox1";
            this.radCheckedListBox1.Size = new System.Drawing.Size(149, 98);
            this.radCheckedListBox1.TabIndex = 3;
            this.radCheckedListBox1.Text = "radCheckedListBox1";
            this.radCheckedListBox1.SelectedItemChanged += new System.EventHandler(this.radCheckedListBox1_SelectedItemChanged);
            this.radCheckedListBox1.ItemCheckedChanged += new Telerik.WinControls.UI.ListViewItemEventHandler(this.radCheckedListBox1_ItemCheckedChanged);
            this.radCheckedListBox1.Click += new System.EventHandler(this.radCheckedListBox1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "testo da cercare....";
            // 
            // resourcesBindingSource1
            // 
            this.resourcesBindingSource1.DataMember = "Resources";
            this.resourcesBindingSource1.DataSource = this.zSIDataSetBindingSource;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.chkPrevisionali);
            this.splitContainer1.Panel1.Controls.Add(this.chkSearch);
            this.splitContainer1.Panel1.Controls.Add(this.radCalendar1);
            this.splitContainer1.Panel1.Controls.Add(this.panel1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.radSchedulerNavigator1);
            this.splitContainer1.Panel2.Controls.Add(this.radScheduler1);
            this.splitContainer1.Size = new System.Drawing.Size(863, 557);
            this.splitContainer1.SplitterDistance = 168;
            this.splitContainer1.TabIndex = 2;
            // 
            // radCalendar1
            // 
            this.radCalendar1.AllowMultipleView = true;
            this.radCalendar1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.radCalendar1.Location = new System.Drawing.Point(3, 202);
            this.radCalendar1.MultiViewRows = 3;
            this.radCalendar1.Name = "radCalendar1";
            this.radCalendar1.Size = new System.Drawing.Size(168, 352);
            this.radCalendar1.TabIndex = 0;
            this.radCalendar1.Text = "radCalendar1";
            this.radCalendar1.SelectionChanged += new System.EventHandler(this.radCalendar1_SelectionChanged);
            // 
            // radSchedulerNavigator1
            // 
            this.radSchedulerNavigator1.AssociatedScheduler = this.radScheduler1;
            this.radSchedulerNavigator1.Controls.Add(this.btnCOA);
            this.radSchedulerNavigator1.Controls.Add(this.btnPrint);
            this.radSchedulerNavigator1.Controls.Add(this.rulerTO);
            this.radSchedulerNavigator1.Controls.Add(this.rulerFROM);
            this.radSchedulerNavigator1.Controls.Add(this.label2);
            this.radSchedulerNavigator1.Controls.Add(this.btnUpdate);
            this.radSchedulerNavigator1.DateFormat = "dd/MM/yyyy";
            this.radSchedulerNavigator1.Dock = System.Windows.Forms.DockStyle.Top;
            this.radSchedulerNavigator1.Location = new System.Drawing.Point(0, 0);
            this.radSchedulerNavigator1.Name = "radSchedulerNavigator1";
            this.radSchedulerNavigator1.NavigationStepType = Telerik.WinControls.UI.NavigationStepTypes.Day;
            // 
            // 
            // 
            this.radSchedulerNavigator1.RootElement.StretchVertically = false;
            this.radSchedulerNavigator1.Size = new System.Drawing.Size(691, 77);
            this.radSchedulerNavigator1.TabIndex = 1;
            this.radSchedulerNavigator1.Text = "radSchedulerNavigator1";
            // 
            // btnCOA
            // 
            this.btnCOA.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCOA.Location = new System.Drawing.Point(461, 44);
            this.btnCOA.Name = "btnCOA";
            this.btnCOA.Size = new System.Drawing.Size(48, 23);
            this.btnCOA.TabIndex = 7;
            this.btnCOA.Text = "COA";
            this.btnCOA.UseVisualStyleBackColor = true;
            this.btnCOA.Click += new System.EventHandler(this.btnCOA_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.btnPrint.Location = new System.Drawing.Point(400, 44);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(55, 23);
            this.btnPrint.TabIndex = 6;
            this.btnPrint.Text = "Stampa";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // rulerTO
            // 
            this.rulerTO.Location = new System.Drawing.Point(359, 47);
            this.rulerTO.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.rulerTO.Minimum = new decimal(new int[] {
            12,
            0,
            0,
            0});
            this.rulerTO.Name = "rulerTO";
            this.rulerTO.Size = new System.Drawing.Size(35, 18);
            this.rulerTO.TabIndex = 5;
            this.rulerTO.Value = new decimal(new int[] {
            12,
            0,
            0,
            0});
            this.rulerTO.ValueChanged += new System.EventHandler(this.rulerTO_ValueChanged);
            // 
            // rulerFROM
            // 
            this.rulerFROM.Location = new System.Drawing.Point(320, 47);
            this.rulerFROM.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.rulerFROM.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.rulerFROM.Name = "rulerFROM";
            this.rulerFROM.Size = new System.Drawing.Size(35, 18);
            this.rulerFROM.TabIndex = 4;
            this.rulerFROM.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.rulerFROM.ValueChanged += new System.EventHandler(this.rulerFROM_ValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(247, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 19);
            this.label2.TabIndex = 3;
            this.label2.Text = "Dalle/Alle";
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.Color.SteelBlue;
            this.btnUpdate.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpdate.ForeColor = System.Drawing.Color.White;
            this.btnUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnUpdate.ImageKey = "table-save-icon.png";
            this.btnUpdate.ImageList = this.imageList16;
            this.btnUpdate.Location = new System.Drawing.Point(515, 44);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(125, 25);
            this.btnUpdate.TabIndex = 2;
            this.btnUpdate.Text = "Salva Modifiche";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // imageList16
            // 
            this.imageList16.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList16.ImageStream")));
            this.imageList16.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList16.Images.SetKeyName(0, "filter-add-icon.png");
            this.imageList16.Images.SetKeyName(1, "filter-delete-icon.png");
            this.imageList16.Images.SetKeyName(2, "Log-Out-icon.png");
            this.imageList16.Images.SetKeyName(3, "Programming-Save-icon.png");
            this.imageList16.Images.SetKeyName(4, "table-save-icon.png");
            // 
            // resourcesBindingSource
            // 
            this.resourcesBindingSource.DataMember = "Resources";
            this.resourcesBindingSource.DataSource = this.zSIDataSetBindingSource;
            // 
            // appointmentsBindingSource
            // 
            this.appointmentsBindingSource.DataSource = this.zSIDataSetBindingSource;
            this.appointmentsBindingSource.Position = 0;
            this.appointmentsBindingSource.CurrentChanged += new System.EventHandler(this.appointmentsBindingSource_CurrentChanged);
            // 
            // appointmentsTableAdapter
            // 
            this.appointmentsTableAdapter.ClearBeforeFill = true;
            // 
            // zsiDataSet1
            // 
            this.zsiDataSet1.DataSetName = "ZSIDataSet";
            this.zsiDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // resourcesTableAdapter
            // 
            this.resourcesTableAdapter.ClearBeforeFill = true;
            // 
            // radPrintDocument1
            // 
            this.radPrintDocument1.FooterFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radPrintDocument1.HeaderFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radPrintDocument1.Watermark = radPrintWatermark1;
            // 
            // radPrintDocument2
            // 
            this.radPrintDocument2.FooterFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radPrintDocument2.HeaderFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radPrintDocument2.Watermark = radPrintWatermark2;
            // 
            // chkPrevisionali
            // 
            this.chkPrevisionali.AutoSize = true;
            this.chkPrevisionali.Location = new System.Drawing.Point(7, 179);
            this.chkPrevisionali.Name = "chkPrevisionali";
            this.chkPrevisionali.Size = new System.Drawing.Size(97, 16);
            this.chkPrevisionali.TabIndex = 5;
            this.chkPrevisionali.Text = "Mostra Previsioni";
            this.chkPrevisionali.UseVisualStyleBackColor = true;
            // 
            // ItaCalendarObject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "ItaCalendarObject";
            this.Size = new System.Drawing.Size(863, 557);
            this.Load += new System.EventHandler(this.ItaCalendarObject_Load);
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1.EventProvider)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1.ResourceProvider)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.zSIDataSetBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ZSIDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radScheduler1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radCheckedListBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesBindingSource1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.radCalendar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radSchedulerNavigator1)).EndInit();
            this.radSchedulerNavigator1.ResumeLayout(false);
            this.radSchedulerNavigator1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rulerTO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rulerFROM)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.resourcesBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.appointmentsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.zsiDataSet1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Telerik.WinControls.UI.SchedulerBindingDataSource schedulerBindingDataSource1;
        private Telerik.WinControls.UI.RadScheduler radScheduler1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private Telerik.WinControls.UI.RadCalendar radCalendar1;
        private System.Windows.Forms.BindingSource appointmentsBindingSource;
        private System.Windows.Forms.BindingSource zSIDataSetBindingSource;
        private ToDoNotificheBSC.ZSIDataSetTableAdapters.AppointmentsTableAdapter appointmentsTableAdapter;
        private Telerik.WinControls.UI.RadSchedulerNavigator radSchedulerNavigator1;
        public ToDoNotificheBSC.ZSIDataSet ZSIDataSet;
        private ToDoNotificheBSC.ZSIDataSet zsiDataSet1;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.CheckBox chkSearch;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Label label1;
        private Telerik.WinControls.UI.RadCheckedListBox radCheckedListBox1;
        private System.Windows.Forms.BindingSource resourcesBindingSource;
        private ZSIDataSetTableAdapters.ResourcesTableAdapter resourcesTableAdapter;
        private System.Windows.Forms.BindingSource resourcesBindingSource1;
        private System.Windows.Forms.ImageList imageList16;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown rulerTO;
        private System.Windows.Forms.NumericUpDown rulerFROM;
        private System.Windows.Forms.Button btnPrint;
        private Telerik.WinControls.UI.RadPrintDocument radPrintDocument1;
        private Telerik.WinControls.UI.RadPrintDocument radPrintDocument2;
        private System.Windows.Forms.Button btnCOA;
        private System.Windows.Forms.CheckBox chkPrevisionali;
    }
}
