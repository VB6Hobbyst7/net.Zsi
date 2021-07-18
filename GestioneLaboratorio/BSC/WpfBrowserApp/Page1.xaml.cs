using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ToDoNotificheBSC;
using System.Windows.Forms.Integration;

namespace WpfBrowserApp
{
    /// <summary>
    /// Logica di interazione per Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {

        private readonly frmMagazzino mainForm = new frmMagazzino();


        public Page1()
        {
            InitializeComponent();

            //Create a Windows Forms Host to host a form
            WindowsFormsHost windowsFormsHost = new WindowsFormsHost();

            stackPanel.Width = mainForm.Width;
            stackPanel.Height = mainForm.Height;
            windowsFormsHost.Width = mainForm.Width;
            windowsFormsHost.Height = mainForm.Height;

            mainForm.TopLevel = false;

            windowsFormsHost.Child = mainForm;

            stackPanel.Children.Add(windowsFormsHost);
        }
    }
}
