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
using ToDoNotificheBSC.ZSIDataSetTableAdapters;
using System.IO;

namespace ToDoNotificheBSC
{
    public partial class ItaCalendarObject : UserControl
    {
        SchedulerDailyPrintStyle dailyStyle = new SchedulerDailyPrintStyle();
        SchedulerWeeklyPrintStyle weekStyle = new SchedulerWeeklyPrintStyle();

        public string cnstring = "";

        public int chkViewNONCISTERNA = 0;
        public int chkViewCISTERNA = 0;

        public int chkViewITALIA = 0;
        public int chkViewESTERO = 0;

        static bool lFilter = false;


        SchedulerDayView dayView;
        SchedulerWeekView weekView;
        AppointmentsTableAdapter appointmentsAdapter = new AppointmentsTableAdapter();
        ResourcesTableAdapter resourcesAdapter = new ResourcesTableAdapter();
        AppointmentsResourcesTableAdapter appointmentsResourcesAdapter = new AppointmentsResourcesTableAdapter();

        public ItaCalendarObject()
        {
            InitializeComponent();
        }

        private void ItaCalendarObject_Load(object sender, EventArgs e)
        {

            dailyStyle.CellElementFormatting += new PrintSchedulerCellEventHandler(dailyStyle_CellElementFormatting);

            dayView = radScheduler1.GetDayView();
            weekView = radScheduler1.GetWeekView();

            rulerFROM.Value = 7;
            rulerTO.Value = 20;
//            dayView.RulerStartScale = int.Parse(rulerFROM.Value.ToString());
//            dayView.RulerEndScale = int.Parse(rulerTO.Value.ToString());
            dayView.RangeFactor = ScaleRange.HalfHour;
            dayView.WorkWeekStart = DayOfWeek.Monday;
            dayView.WorkWeekEnd = DayOfWeek.Friday;
            dayView.DayCount = 1;


            LoadItems();
        }

        public void loadfilter()
        {

            lFilter = true;

            try
            {
                for (int i = 0; i < radCheckedListBox1.Items.Count; i++)
                {
                    radCheckedListBox1.Items[i].CheckState = Telerik.WinControls.Enumerations.ToggleState.Off;
                }

                if (chkViewITALIA == 1)
                {
                    radCheckedListBox1.Items[2].CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
                }

                if (chkViewESTERO == 1)
                {
                    radCheckedListBox1.Items[3].CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
                }

                if (chkViewCISTERNA == 1)
                {
                    radCheckedListBox1.Items[0].CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
                }

                if (chkViewNONCISTERNA == 1)
                {
                    radCheckedListBox1.Items[1].CheckState = Telerik.WinControls.Enumerations.ToggleState.On;
                }
            }
            catch (Exception ex)
            {

                lFilter = false;
            }
            finally
            {
                lFilter = false;
            }
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




            DateTime d1 = System.DateTime.Today;
            DateTime d2 = System.DateTime.Today;

            if (radCalendar1.SelectedDates.Count > 0)
            {
                d1 = radCalendar1.SelectedDates.First();
                d2 = radCalendar1.SelectedDates.Last();
            }


            //ZSIDataSet.Appointments.Select(string.Format("Start >= '{0}' and start < '{1}'", "20170401", "21000101"));

            
            

            resourcesAdapter.Fill(ZSIDataSet.Resources);

            //if (radCheckedListBox1.Items.Count == 0)
            //{
            //    for (int i = 0; i < ZSIDataSet.Resources.Rows.Count; i++)
            //    {
            //        ListViewDataItem l = new ListViewDataItem();
            //        l.Key = ZSIDataSet.Resources.Rows[i][0].ToString();
            //        l.Text = ZSIDataSet.Resources.Rows[i][1].ToString();
            //        radCheckedListBox1.Items.Add(l); // ZSIDataSet.Resources.Rows[i][0].ToString(), ZSIDataSet.Resources.Rows[i][1].ToString());
            //    }
            //    //        radCheckedListBox1.DataSource = ZSIDataSet.Resources;
            //    //radCheckedListBox1.DataMember = "ID";
            //    //radCheckedListBox1.DisplayMember = "Name";
            //}




            short cisterna = 0;
            if (radCheckedListBox1.Items[0].CheckState == Telerik.WinControls.Enumerations.ToggleState.On)
            {
                cisterna = 1;
            }

            short imballato = 0;
            if (radCheckedListBox1.Items[1].CheckState == Telerik.WinControls.Enumerations.ToggleState.On)
            {
                imballato = 1;
            }

            short italia = 0;
            if (radCheckedListBox1.Items[2].CheckState == Telerik.WinControls.Enumerations.ToggleState.On)
            {
                italia = 1;
            }

            short estero = 0;
            if (radCheckedListBox1.Items[3].CheckState == Telerik.WinControls.Enumerations.ToggleState.On)
            {
                estero = 1;
            }

            string tipoimballo = "%";

            if ((cisterna == 1) && (imballato == 0))
            {
                tipoimballo = "C";
            }
            if ((cisterna == 0) && (imballato == 1))
            {
                tipoimballo = "I";
            }

            if ((cisterna == 0) && (imballato == 0))
            {
                tipoimballo = "-";
            }

            string tipoordine = "%";
            if ((italia == 1) && (estero == 0))
            {
                tipoordine = "I";
            }

            if ((italia == 0) && (estero == 1))
            {
                tipoordine = "E";
            }

            if ((italia == 0) && (estero == 0))
            {
                tipoordine = "-";
            }

            short previsionali = 0;

            if (chkPrevisionali.Checked)
                previsionali = 1;
             

            appointmentsAdapter.Fill(ZSIDataSet.Appointments, d1.AddDays(-30), d2.AddDays(30), txtSearch.Text, tipoimballo, tipoordine, previsionali);

            appointmentsResourcesAdapter.Fill(ZSIDataSet.AppointmentsResources);


            schedulerBindingDataSource1.Rebind();


            //radScheduler1.DataSource = schedulerBindingDataSource1;
            //    radScheduler1.Refresh();

            //radScheduler1.SchedulerElement.Refresh();
            radScheduler1.AppointmentTitleFormat = "{2} {3}";

        }




        private void radCalendar1_SelectionChanged(object sender, EventArgs e)
        {
            radScheduler1.ActiveView.StartDate = radCalendar1.SelectedDate;

            LoadItems();

        }

        private void appointmentsBindingSource_CurrentChanged(object sender, EventArgs e)
        {
            
        }

        private void zSIDataSetBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {

            resourcesAdapter.Adapter.AcceptChangesDuringUpdate = false;


            ZSIDataSet.AppointmentsResourcesDataTable deletedRelationRecords =
                            ZSIDataSet.AppointmentsResources.GetChanges(DataRowState.Deleted)
                            as ZSIDataSet.AppointmentsResourcesDataTable;
            ZSIDataSet.AppointmentsResourcesDataTable newRelationRecords =
                            ZSIDataSet.AppointmentsResources.GetChanges(DataRowState.Added)
                            as ZSIDataSet.AppointmentsResourcesDataTable;
            ZSIDataSet.AppointmentsResourcesDataTable modifiedRelationRecords =
                            ZSIDataSet.Appointments.GetChanges(DataRowState.Modified)
                            as ZSIDataSet.AppointmentsResourcesDataTable;
            ZSIDataSet.AppointmentsDataTable newAppointmentRecords =
                ZSIDataSet.Appointments.GetChanges(DataRowState.Added) as ZSIDataSet.AppointmentsDataTable;
            ZSIDataSet.AppointmentsDataTable deletedAppointmentRecords =
                ZSIDataSet.Appointments.GetChanges(DataRowState.Deleted) as ZSIDataSet.AppointmentsDataTable;
            ZSIDataSet.AppointmentsDataTable modifiedAppointmentRecords =
                ZSIDataSet.Appointments.GetChanges(DataRowState.Modified) as ZSIDataSet.AppointmentsDataTable;


            try
            {
                if (newAppointmentRecords != null)
                {
                    Dictionary<int, int> newAppointmentIds = new Dictionary<int, int>();
                    Dictionary<object, int> oldAppointmentIds = new Dictionary<object, int>();
                    for (int i = 0; i < newAppointmentRecords.Count; i++)
                    {
                        oldAppointmentIds.Add(newAppointmentRecords[i], newAppointmentRecords[i].ID);
                    }
                    appointmentsTableAdapter.Update(newAppointmentRecords);
                    for (int i = 0; i < newAppointmentRecords.Count; i++)
                    {
                        newAppointmentIds.Add(oldAppointmentIds[newAppointmentRecords[i]], newAppointmentRecords[i].ID);
                    }
                    if (newRelationRecords != null)
                    {
                        for (int i = 0; i < newRelationRecords.Count; i++)
                        {
                            newRelationRecords[i].AppointmentID = newAppointmentIds[newRelationRecords[i].AppointmentID];
                        }
                    }
                }
                if (deletedRelationRecords != null)
                {

                    appointmentsResourcesAdapter.Update(deletedRelationRecords);
                }
                if (deletedAppointmentRecords != null)
                {
                    appointmentsAdapter.Update(deletedAppointmentRecords);
                }
                if (modifiedAppointmentRecords != null)
                {
                    appointmentsAdapter.Update(modifiedAppointmentRecords);
                }
                if (newRelationRecords != null)
                {
                    appointmentsResourcesAdapter.Update(newRelationRecords);
                }
                if (modifiedRelationRecords != null)
                {
                    appointmentsResourcesAdapter.Update(modifiedRelationRecords);
                }
                this.zsiDataSet1.AcceptChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Impossibile aggiornare :\n{0}\n{1}", ex.Message, "Ricaricare i posizionamenti e riprovare"));
            }
            finally
            {
                if (deletedRelationRecords != null)
                {
                    deletedRelationRecords.Dispose();
                }
                if (newRelationRecords != null)
                {
                    newRelationRecords.Dispose();
                }
                if (modifiedRelationRecords != null)
                {
                    modifiedRelationRecords.Dispose();
                }
            }

            LoadItems();

            //lblStatus.Text = "Updated scheduler at " + DateTime.Now.ToString();
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //enter key is down
                //if (txtSearch.Text.Trim() == "")
                //{
                //}
                //else
                //{
                    LoadItems();
                //}


            }
        }

        private void chkSearch_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSearch.Checked)
            {
                //panel1.Visible = true;
            }
            else {
                txtSearch.Text = "";
                //panel1.Visible = false;
                LoadItems();
            }
        }

        private void radCheckedListBox1_SelectedItemChanged(object sender, EventArgs e)
        {
            
        }

        private void radCheckedListBox1_Click(object sender, EventArgs e)
        {
        }

        private void radCheckedListBox1_ItemCheckedChanged(object sender, ListViewItemEventArgs e)
        {
            if (lFilter)
            {
            }
            else
            {
                LoadItems();
            }

        }

        private void rulerTO_ValueChanged(object sender, EventArgs e)
        {
            dayView.RulerEndScale =  int.Parse(rulerTO.Value.ToString());
        }

        private void rulerFROM_ValueChanged(object sender, EventArgs e)
        {
            dayView.RulerStartScale = int.Parse(rulerFROM.Value.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            string reportTitle = string.Format("Cliente: {0}", txtSearch.Text);

            switch (radScheduler1.ActiveView.ViewType)
            {
                case SchedulerViewType.Day:
                    this.radScheduler1.PrintStyle = dailyStyle;
                    dailyStyle.DateStartRange = radCalendar1.SelectedDate;
                    dailyStyle.DateEndRange = radCalendar1.SelectedDate.AddDays(7);

                    dailyStyle.TimeStartRange = TimeSpan.FromHours(double.Parse(rulerFROM.Value.ToString()));
                    dailyStyle.TimeEndRange = TimeSpan.FromHours(double.Parse(rulerTO.Value.ToString()));

                    dailyStyle.DateHeadingFont = new Font("Segoe UI", 8, FontStyle.Bold);
                    dailyStyle.AppointmentFont = new Font("Arial", 10, FontStyle.Regular);
                    dailyStyle.PageHeadingFont = new Font("Segoe UI", 10, FontStyle.Bold);

                    dailyStyle.HoursColumnWidth = 1;

                    dailyStyle.ShowNotesArea = false;
                    dailyStyle.ShowLinedNotesArea = false;

                    dailyStyle.DrawPageTitle = true;
                    dailyStyle.DrawPageTitleCalendar = false;

                    dailyStyle.HeadingAreaHeight = 30;
                    dailyStyle.ShowTimezone = false;
                    break;

                case SchedulerViewType.Week:
                case SchedulerViewType.WorkWeek:
                    this.radScheduler1.PrintStyle = weekStyle;
                    weekStyle.DateStartRange = radCalendar1.SelectedDate;
                    weekStyle.DateEndRange = radCalendar1.SelectedDate.AddDays(7);

                    weekStyle.TimeStartRange = TimeSpan.FromHours(double.Parse(rulerFROM.Value.ToString()));
                    weekStyle.TimeEndRange = TimeSpan.FromHours(double.Parse(rulerTO.Value.ToString()));

                    weekStyle.DateHeadingFont = new Font("Segoe UI", 8, FontStyle.Bold);
                    weekStyle.AppointmentFont = new Font("Arial", 8, FontStyle.Regular);
                    weekStyle.PageHeadingFont = new Font("Segoe UI", 10, FontStyle.Bold);

                    weekStyle.DaysLayout = WeeklyStyleLayout.LeftToRight;

                    weekStyle.ShowNotesArea = false;
                    weekStyle.ShowLinedNotesArea = false;
                    
                    weekStyle.DrawPageTitle = true;
                    weekStyle.DrawPageTitleCalendar = false;

                    weekStyle.HeadingAreaHeight = 30;
                    weekStyle.ShowTimezone = false;

                    weekStyle.ExcludeNonWorkingDays = (radScheduler1.ActiveView.ViewType == SchedulerViewType.WorkWeek);
                    break;

            }
                            

            //radPrintDocument1.MiddleHeader = String.Format("Posizionamenti del {0}", string.Format("{0:MM/dd/yyyy}", radCalendar1.SelectedDate));
            radPrintDocument1.Margins.Left = 20;
            radPrintDocument1.Margins.Right = 20;
            radPrintDocument1.Margins.Top = 20;
            radPrintDocument1.Margins.Bottom = 20;




            radScheduler1.PrintPreview(radPrintDocument1);




        }

        

        void dailyStyle_CellElementFormatting(object sender, PrintSchedulerCellEventArgs e)
        {

            //e.CellElement.Font = new Font("Arial", 4, FontStyle.Regular);

            //if (e.CellElement.Date.Hour >= 12 && e.CellElement.Date.Hour < 13)
            //{
            //    e.CellElement.DrawFill = true;
            //    e.CellElement.BackColor = Color.OrangeRed;
            //}
            
        }

        private void radScheduler1_AppointmentAdded(object sender, AppointmentAddedEventArgs e)
        {
            e.Appointment.ToolTipText = e.Appointment.Description;
        }

        private void btnCOA_Click(object sender, EventArgs e)
        {

            Cursor.Current = Cursors.WaitCursor;

            bool bOK = true;
            string modulo = "COA";


            try
            {

                string xslxmodello = Path.Combine(Application.StartupPath, "XLSXModelli", string.Format("{0}.xlsm", modulo));
                string outfile = Path.Combine(Application.StartupPath, string.Format("{0}.xlsm", modulo));

                string tmpfile = Path.Combine(Path.GetTempPath(), Path.GetTempFileName() + ".xlsm");

                //toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", tmpfile);

                ExportExcel.excelFile = tmpfile;

                //DataRowView r = ((DataRowView)comboBoxSpedizionieri.SelectedItem);


                ExportExcel.creaModuloCOA(radCalendar1.SelectedDate, xslxmodello, outfile, 0, 9);

                //toolStripStatusLabelLOG.Text = string.Format("Export Excel {0}", outfile);

                //toolStripStatusLabelLOG.Text = string.Format("{0}", "");


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Cursor.Current = Cursors.Default;
                //toolStripStatusLabelLOG.Text = string.Format("{0}", "");

            }
            finally
            {
                Cursor.Current = Cursors.Default;
                //toolStripStatusLabelLOG.Text = string.Format("{0}", "");
            }

        }
    }
}
