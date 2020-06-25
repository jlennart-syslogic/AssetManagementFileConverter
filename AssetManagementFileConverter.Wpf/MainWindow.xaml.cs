using MPSAssestManagementFileConverter.BusinessLogic;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
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

namespace MPSAssetManagementFileConverter.Wpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            companyTextBox.Text = ReadSetting("Company");
            csvHeadersTextBox.Text = ReadSetting("OutputHeader");
        }

        private void Button_Click_Select_File(object sender, RoutedEventArgs e)
        {
            var filename = OpenDialog();
            inputFileNameTextBlock.Text = filename;
            outputFileNameTextBox.Text = System.IO.Path.ChangeExtension(filename, ".csv");
        }


        private void OpenCSV(string outputFileLocation)
        {
            try
            {
                Process.Start("Excel.exe", outputFileLocation);
            }
            catch (Exception)
            {
                try
                {
                    Process.Start("notepad.exe", outputFileLocation);
                } 
                catch (Exception)
                {
                    messageTextBlock.Text = "Unable to open file.";
                }
            }
        }

        private void CreateCSV()
        {
            var filename = inputFileNameTextBlock.Text;            
            var companyName = companyTextBox.Text;
            var outputHeader = csvHeadersTextBox.Text;
            var outputFileName = outputFileNameTextBox.Text;
            var sb = new StringBuilder();
            
            if (string.IsNullOrWhiteSpace(filename))
            {
                sb.AppendLine("Input File is required.");                
            }

            if (string.IsNullOrWhiteSpace(outputFileName))
            {
                sb.AppendLine("Output File is required.");
            }

            if (string.IsNullOrWhiteSpace(companyName))
            {
                sb.AppendLine("Company is required");
            }

            if (string.IsNullOrWhiteSpace(outputHeader))
            {
                sb.AppendLine("Output Header is required.");
            }

            if(sb.Length > 0)
            {
                messageTextBlock.Text = sb.ToString();
                openCsvButton.Visibility = Visibility.Hidden;
                return;
            }

            try
            {
                var converter = new FileConverter();
                converter.ConvertFile(filename, outputFileName, companyName, outputHeader);
               
                messageTextBlock.Text = $"Conversion Successful!";
                openCsvButton.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                messageTextBlock.Text = $"Conversion Failed - {ex.Message}";
            }
        }

        private string OpenDialog()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog
            {

                // Set filter for file extension and default file extension 
                DefaultExt = ".xlsx",
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result.Value == true)
            {
                // Open document 
                string filename = dlg.FileName;
                //textBox1.Text = filename;
                return filename;
            }
            return null;
        }

        private static string ReadSetting(string key)
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                string result = appSettings[key] ?? "Not Found";
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Error reading app settings");
                return null;
            }
        }

        private void Button_Click_Create_CSV(object sender, RoutedEventArgs e)
        {
            CreateCSV();
        }

        private void Button_Click_Open_CSV(object sender, RoutedEventArgs e)
        {
            var outputFileName = outputFileNameTextBox.Text;
            OpenCSV(outputFileName);
        }
    }
}
