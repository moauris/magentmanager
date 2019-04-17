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
using ControlMAgent.NewTask;
using ControlMAgent.CCM;

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
            //dgCCM.DataContext = CCMReader.FillDataset();
        }

        // When button is clicked, pop file open dialog.
        // Using the FileInfo to the Database.
        private async void btnNewRequest_Clicked(object sender, RoutedEventArgs e)
        {
            //Disable the button while running, after that enable
            Button ThisButton = (Button)sender;
            ThisButton.IsEnabled = false;
            FileInfo[] TargetExcelFiles = await NewMARequest.TaskNewFileOpen();
            sbarProgress.Maximum = 1000;
            var updatePbar = new Progress<double>(RunningProgress);
            var ProgressHandler = updatePbar as IProgress<double>;
            
            if (TargetExcelFiles[0].Name != "null")
                await Task.Run(() => NewMARequest.TaskExcelToDatabase
                    (TargetExcelFiles, ProgressHandler));
            //If the use Task type instead of void, progress bar will not update.

            //sbarProgress.Value = 100; sbarTextBox.Content = "Status: Ready.";
            //sbarProgress.IsIndeterminate = false; sbarProgress.Value = 0;


            ThisButton.IsEnabled = true;
        }

        public void RunningProgress(double dProgress)
        {
            sbarProgress.Value = dProgress;
        }
    }
}
