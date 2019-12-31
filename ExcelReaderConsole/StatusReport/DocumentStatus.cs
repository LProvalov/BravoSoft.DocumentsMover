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
        private readonly StringBuilder stringBuilder;
        public DocumentStatus(Document document)
        {
            this.document = document;
            stringBuilder = new StringBuilder();
        }

        public bool TextFileExist { get; set; } = false;
        public bool ScanFileExist { get; set; } = false;
        public bool TextFileWasMoved { get; set; } = false;
        public bool ScanFileWasMoved { get; set; } = false;

        public string StatusMessage => stringBuilder.ToString();

        public void StatusMessageAppendLine(string line)
        {
            stringBuilder.AppendLine(line);
        }
    }
}
