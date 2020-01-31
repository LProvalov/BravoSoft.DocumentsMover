using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole.Logger
{
    public class ErrorMessage : MessageBase
    {
        public ErrorMessage(Document document, string message) : base(document, message)
        {
            Type = MessageType.log;
        }
    }
}
