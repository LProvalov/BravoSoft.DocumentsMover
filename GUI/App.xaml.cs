using System.IO;
using System.Windows;
using System.Windows.Threading;
using ExcelReaderConsole;

namespace GUI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        DocumentManager documentManager = DocumentManager.Instance;

        public App()
        {
        }

        private void App_OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            ExceptionHandling ehWindow = new ExceptionHandling(e.Exception);
            ehWindow.ShowDialog();
            e.Handled = true;
        }
    }
}
