using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole.Logger
{
    public abstract class MessageBase
    {
        public enum MessageType
        {
            error = 0,
            log,
            warning,
        }

        protected MessageBase(Document document, string message)
        {
            Message = message;
            this.document = document;
            Time = DateTime.Now;
        }

        public DateTime Time { get; private set; }
        protected Document document;

        public string DocumentId
        {
            get { return document.Identifier; }
        }
        public MessageType Type { get; protected set; }
        public string Message { get; protected set; }

    }
}
