using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ExcelReaderConsole;
using GUI.Models;

namespace GUI
{
    /// <summary>
    /// Interaction logic for AppSettings.xaml
    /// </summary>
    public partial class AppSettings : Window
    {
        private ExcelReaderConsole.AppSettings appSettings = ExcelReaderConsole.AppSettings.Instance;
        public AppSettings()
        {   
            InitializeComponent();
            DataContext = new AppSettingsViewModel()
            {
                InputDirectory = appSettings.InputDirectoryPath,
                OutputDirectory = appSettings.OutputDirectoryPath,
                ExcelTemplatePath = appSettings.ExcelTemplateFilePath
            };
        }

        private DirectoryInfo CreateIfNotExists(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            if (!di.Exists)
            {
                di.Create();
            }
            return di;
        }

        private void inputBrowsButton_Click(object sender, RoutedEventArgs args)
        {
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            System.Windows.Forms.DialogResult result = folderDialog.ShowDialog(this.GetIWin32Window());
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                this.inputDirectoryTextBox.Text = folderDialog.SelectedPath;
            }
        }

        private void outputBrowsButton_Click(object sender, RoutedEventArgs args)
        {
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            System.Windows.Forms.DialogResult result = folderDialog.ShowDialog(this.GetIWin32Window());
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                this.outputDirectoryTextBox.Text = folderDialog.SelectedPath;
            }
        }

        private void excelTemplateBrowsButton_Click(object sender, RoutedEventArgs args)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel template file (*.excel)|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            bool? dialogResult = openFileDialog.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value == true)
            {
                this.excelTemplateTextBox.Text = openFileDialog.FileName;
            }
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(outputDirectoryTextBox.Text))
                {
                    appSettings.OutputDirectoryPath = CreateIfNotExists(outputDirectoryTextBox.Text)?.FullName ?? string.Empty;
                }
                if (!string.IsNullOrEmpty(inputDirectoryTextBox.Text))
                {
                    appSettings.InputDirectoryPath = CreateIfNotExists(inputDirectoryTextBox.Text)?.FullName ?? string.Empty;
                }
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("Can't save settings.");
                return;
            }

            appSettings.SaveAppSettings();
            DialogResult = true;
            Close();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
