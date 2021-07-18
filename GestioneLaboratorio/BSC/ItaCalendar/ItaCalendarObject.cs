using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Telerik.WinControls.UI.Data;
using Telerik.WinControls.UI;
using ItaCalendar.ZSIDataSetTableAdapters;

namespace ItaCalendar
{
    public partial class ItaCalendarObject : UserControl
    {
        public string cnstring = "";


        
        public ItaCalendarObject()
        {
            InitializeComponent();
        }

        private void ItaCalendarObject_Load(object sender, EventArgs e)
        {
           

            SchedulerDayView dayView = radScheduler1.GetDayView();
            dayView.RulerStartScale = 7;
            dayView.RulerEndScale = 19;
            dayView.RangeFactor = ScaleRange.HalfHour;

            LoadItems();
        }


        public void LoadItems()
        {
            //    SchedulerBindingDataSource dataSource = new SchedulerBindingDataSource();
            //    AppointmentMappingInfo appointmentMappingInfo = new AppointmentMappingInfo();
            //    appointmentMappingInfo.Start = "Start";
            //    appointmentMappingInfo.End = "End";
            //    appointmentMappingInfo.Summary = "Subject";
            //    appointmentMappingInfo.Description = "Description";
            //    appointmentMappingInfo.Location = "Location";
            //    appointmentMappingInfo.UniqueId = "Id";
            //    SchedulerMapping idMapping = appointmentMappingInfo.FindByDataSourceProperty("Id");
            //    idMapping.ConvertToDataSource = new Telerik.WinControls.UI.Data.ConvertCallback(this.ConvertIdToDataSource);
            //    idMapping.ConvertToScheduler = new Telerik.WinControls.UI.Data.ConvertCallback(this.ConvertIdToScheduler);
            //    dataSource.EventProvider.Mapping = appointmentMappingInfo;
            //    dataSource.EventProvider.DataSource = this.appointments;
            //    this.radScheduler1.DataSource = dataSource;
            //radScheduler1.AppointmentTitleFormat = "{2} {3}";

            SchedulerDayView dayView = radScheduler1.GetDayView();
            dayView.RulerStartScale = 7;
            dayView.RulerEndScale = 19;
            dayView.RangeFactor = ScaleRange.HalfHour;
            dayView.WorkWeekStart = DayOfWeek.Monday;
            dayView.WorkWeekEnd = DayOfWeek.Friday;
            dayView.DayCount = 1;

            AppointmentsTableAdapter appointmentsAdapter = new AppointmentsTableAdapter();
            appointmentsAdapter.Fill(this.zSIDataSet.Appointments);
            ResourcesTableAdapter resourcesAdapter = new ResourcesTableAdapter();
            resourcesAdapter.Fill(this.zSIDataSet.Resources);
            AppointmentsResourcesTableAdapter appointmentsResourcesAdapter = new AppointmentsResourcesTableAdapter();
            appointmentsResourcesAdapter.Fill(this.zSIDataSet.AppointmentsResources);

            schedulerBindingDataSource1.Rebind();

        }

        private void radCalendar1_SelectionChanged(object sender, EventArgs e)
        {
            radScheduler1.ActiveView.StartDate = radCalendar1.SelectedDate;

        }
    }
}
