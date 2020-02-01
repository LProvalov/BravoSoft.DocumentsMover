using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole.Logger
{
    public class WarningMessage : MessageBase
    {
        public enum PlaceType
        {
            TextFile = 0,
            Scan,
            TextPdf,
            Attachments
        }
        private PlaceType placeFlag;
        public WarningMessage(Document document, string message, PlaceType placeFlag) : base(document, message)
        {
            Type = MessageType.warning;
            this.placeFlag = placeFlag;
        }

        public PlaceType Place
        {
            get => placeFlag;
        }
    }
}
