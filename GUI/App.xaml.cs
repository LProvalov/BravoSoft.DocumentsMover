using System.IO;
using System.Windows;
using System.Windows.Threading;
using ExcelReaderConsole;
using GUI.Models;

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
            ExceptionHandling exceptionHandlingWindow = ExceptionHandling.Instance;
            exceptionHandlingWindow.Model.AddException(e.Exception);
            exceptionHandlingWindow.Show();
            e.Handled = true;
        }
    }
}
