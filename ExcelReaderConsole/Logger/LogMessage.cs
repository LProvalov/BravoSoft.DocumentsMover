using ExcelReaderConsole.Models;

namespace ExcelReaderConsole.Logger
{
    public class LogMessage : MessageBase
    {
        public LogMessage(Document document, string message) : base(document, message)
        {
            Type = MessageType.log;
        }
    }
}
