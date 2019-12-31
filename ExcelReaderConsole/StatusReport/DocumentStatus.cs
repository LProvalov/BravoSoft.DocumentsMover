using ExcelReaderConsole.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole.StatusReport
{
    public class DocumentStatus
    {
        private Document document;
        public DocumentStatus(Document document)
        {
            this.document = document;
        }

        public bool TextFileExist { get; set; } = false;
        public bool ScanFileExist { get; set; } = false;
        public string StatusMessage { get; set; } = string.Empty;

        public void ConsolePrint()
        {
            if (!ScanFileExist || !TextFileExist)
            {
                Console.WriteLine(StatusMessage);
            }
        }
    }
}
