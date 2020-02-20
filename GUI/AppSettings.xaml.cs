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
        private AppSettingsViewModel model;
        public AppSettings()
        {   
            InitializeComponent();
            model = new AppSettingsViewModel()
            {
                InputDirectory = appSettings.InputDirectoryPath,
                OutputDirectory = appSettings.OutputDirectoryPath
            };
            DataContext = model;
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
                model.InputDirectory = folderDialog.SelectedPath;
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
                model.OutputDirectory = folderDialog.SelectedPath;
            }
        }


        private bool CheckDirectory(string directoryPath, out string fullPath)
        {
            fullPath = string.Empty;
            DirectoryInfo directoryInfo = new DirectoryInfo(directoryPath);
            if (string.IsNullOrEmpty(directoryPath))
            {
                return false;
            }
            if (!directoryInfo.Exists)
            {
                var result = MessageBox.Show($"Директория ({directoryPath}) не существует, создать?",
                                             "Внимание!", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    DirectoryInfo createdDirInfo = CreateIfNotExists(directoryInfo.FullName);
                    if (createdDirInfo != null)
                    {
                        fullPath = createdDirInfo.FullName;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            fullPath = directoryInfo.FullName;
            return true;
        }
        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            bool isOk = true;
            try
            {
                string fullPath;
                if (isOk &= CheckDirectory(outputDirectoryTextBox.Text, out fullPath))
                {
                    appSettings.OutputDirectoryPath = fullPath;
                    model.OutputDirectory = appSettings.OutputDirectoryPath;
                }

                if (isOk &= CheckDirectory(inputDirectoryTextBox.Text, out fullPath))
                {
                    appSettings.InputDirectoryPath = fullPath;
                    model.InputDirectory = appSettings.InputDirectoryPath;
                }
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("Can't save settings.");
                return;
            }

            if (isOk)
            {
                appSettings.SaveAppSettings();
                DialogResult = true;
                Close();
            }
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
