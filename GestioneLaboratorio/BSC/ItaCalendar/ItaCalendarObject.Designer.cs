namespace ItaCalendar
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
            this.schedulerBindingDataSource1 = new Telerik.WinControls.UI.SchedulerBindingDataSource();
            this.radScheduler1 = new Telerik.WinControls.UI.RadScheduler();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.radCalendar1 = new Telerik.WinControls.UI.RadCalendar();
            this.radSchedulerNavigator1 = new Telerik.WinControls.UI.RadSchedulerNavigator();
            this.appointmentsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.zSIDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.zSIDataSet = new ItaCalendar.ZSIDataSet();
            this.appointmentsTableAdapter = new ItaCalendar.ZSIDataSetTableAdapters.AppointmentsTableAdapter();
            this.zsiDataSet1 = new ItaCalendar.ZSIDataSet();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1.EventProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerBindingDataSource1.ResourceProvider)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radScheduler1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.radCalendar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.radSchedulerNavigator1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.appointmentsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.zSIDataSetBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.zSIDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.zsiDataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // schedulerBindingDataSource1
            // 
            // 
            // 
            // 
            this.schedulerBindingDataSource1.EventProvider.DataMember = "Appointments";
            this.schedulerBindingDataSource1.EventProvider.DataSource = this.zSIDataSetBindingSource;
            appointmentMappingInfo1.Description = "Description";
            appointmentMappingInfo1.End = "End";
            appointmentMappingInfo1.Location = "Location";
            appointmentMappingInfo1.Start = "Start";
            appointmentMappingInfo1.Summary = "Start";
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
            // radScheduler1
            // 
            this.radScheduler1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
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
            this.radScheduler1.Size = new System.Drawing.Size(686, 481);
            this.radScheduler1.TabIndex = 0;
            this.radScheduler1.Text = "radScheduler1";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.radCalendar1);
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
            this.radCalendar1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.radCalendar1.Location = new System.Drawing.Point(0, 0);
            this.radCalendar1.MultiViewRows = 3;
            this.radCalendar1.Name = "radCalendar1";
            this.radCalendar1.Size = new System.Drawing.Size(168, 557);
            this.radCalendar1.TabIndex = 0;
            this.radCalendar1.Text = "radCalendar1";
            this.radCalendar1.SelectionChanged += new System.EventHandler(this.radCalendar1_SelectionChanged);
            // 
            // radSchedulerNavigator1
            // 
            this.radSchedulerNavigator1.AssociatedScheduler = this.radScheduler1;
            this.radSchedulerNavigator1.DateFormat = "yyyy/MM/dd";
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
            // appointmentsBindingSource
            // 
            this.appointmentsBindingSource.DataMember = "Appointments";
            this.appointmentsBindingSource.DataSource = this.zSIDataSetBindingSource;
            // 
            // zSIDataSetBindingSource
            // 
            this.zSIDataSetBindingSource.DataSource = this.zSIDataSet;
            this.zSIDataSetBindingSource.Position = 0;
            // 
            // zSIDataSet
            // 
            this.zSIDataSet.DataSetName = "ZSIDataSet";
            this.zSIDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
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
            ((System.ComponentModel.ISupportInitialize)(this.radScheduler1)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.radCalendar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.radSchedulerNavigator1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.appointmentsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.zSIDataSetBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.zSIDataSet)).EndInit();
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
        private ZSIDataSetTableAdapters.AppointmentsTableAdapter appointmentsTableAdapter;
        private Telerik.WinControls.UI.RadSchedulerNavigator radSchedulerNavigator1;
        public ZSIDataSet zSIDataSet;
        private ZSIDataSet zsiDataSet1;
    }
}
