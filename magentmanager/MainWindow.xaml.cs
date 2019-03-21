using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace magentmanager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // When button is clicked, pop file open dialog.
        // Using the FileInfo to the Database.
        private async void btnNewRequest_Clicked(object sender, RoutedEventArgs e)
        {
            FileInfo[] TargetExcelFiles = await NewTask.TaskNewFileOpen();
            //Debug.Print("FileName is : " + TargetExcelFiles.FullName);
            sbarProgress.IsIndeterminate = true;
            sbarTextBox.Content = "Status: Processing New Request.";
            if (TargetExcelFiles[0].Name != "null")
                await NewTask.TaskExcelToDatabase(TargetExcelFiles);

            sbarProgress.Value = 100; sbarTextBox.Content = "Status: Ready.";
            sbarProgress.IsIndeterminate = false; sbarProgress.Value = 0;
        }
    }
}
