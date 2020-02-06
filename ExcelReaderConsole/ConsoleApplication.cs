using ExcelReaderConsole.Models;
using System;
using ExcelReaderConsole.Logger;

namespace ExcelReaderConsole
{
    public class ConsoleApplication
    {
        private readonly DocumentManager documentManager = DocumentManager.Instance;

        public ConsoleApplication()
        {
            documentManager.StatusStringChanged += StatusStringChanged;
            documentManager.DocumentProcessed += DocumentProcessedHandler;
        }

        private void DocumentProcessedHandler(Document document)
        {
            Console.WriteLine($"Document {document.Identifier} processed.");
            var messages = documentManager.loggerManager.GetDocumentMessages(document);
            if (messages != null)
            {
                Console.WriteLine("------------------    LOGS    ------------------------");
                foreach (var message in messages)
                {
                    if (message is LogMessage)
                    {
                        LogMessage logM = (message as LogMessage);
                        Console.WriteLine($"[{logM.Time.ToShortTimeString()}] {logM.Message}");
                    }

                    if (message is ErrorMessage)
                    {
                        ErrorMessage logM = (message as ErrorMessage);
                        Console.WriteLine($"[{logM.Time.ToShortTimeString()}] [Error] {logM.Message}");
                    }
                }

                Console.WriteLine("------------------  END LOGS  ------------------------");
                Console.WriteLine();
            }
        }

        private void StatusStringChanged(object sender, EventArgs e)
        {
            string message = (e as DocumentManager.StatusStringChangedArgs)?.Message ?? string.Empty;
            Console.WriteLine($"Status: {message}");
        }

        public void Run()
        {
            documentManager.ReadDataFromTemplate();
            documentManager.ProcessDocumentsAsync().Wait();
        }
    }
}
