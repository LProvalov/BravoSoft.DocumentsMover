using ExcelReaderConsole.Models;
using System;
using System.IO;
using System.Text;

namespace ExcelReaderConsole
{
    public class ConsoleApplication
    {
        private readonly DocumentManager documentManager = DocumentManager.Instance;

        public ConsoleApplication()
        {
            documentManager.StatusStringChanged += StatusStringChanged;
            documentManager.DocumentProcessingErrorOccured += DocumentProcessingErrorOccured;
        }

        private void DocumentProcessingErrorOccured(Document document, StringBuilder obj)
        {
            Console.WriteLine($"{document.Identifier}:");
            Console.WriteLine(obj.ToString());
        }

        private void StatusStringChanged(object sender, EventArgs e)
        {
            string message = (e as DocumentManager.StatusStringChangedArgs)?.Message ?? string.Empty;
            Console.WriteLine($"Status: {message}");
        }

        public void Run()
        {
            documentManager.ReadDataFromTemplate();
            documentManager.ProcessDocuments();
        }
    }
}
