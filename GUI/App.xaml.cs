using System.IO;
using System.Windows;
using ExcelReaderConsole;

namespace GUI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private readonly ExcelReaderConsole.AppSettings appSettings;
        private readonly FileManager fileManager;
        private readonly CardBuilder cardBuilder;

        public App()
        {
            appSettings = ExcelReaderConsole.AppSettings.Instance;
            appSettings.LoadAppSettings();
            fileManager = FileManager.Instance;
            cardBuilder = CardBuilder.Instance;

            DirectoryInfo outputDirectory = new DirectoryInfo(appSettings.GetOutputDirectoryPath());
            if (!outputDirectory.Exists)
            {
                outputDirectory.Create();
            }
            cardBuilder.SetOutputDirectory(outputDirectory.FullName);
        }
    }
}
