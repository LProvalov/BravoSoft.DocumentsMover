using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole.Logger
{
    public class LoggerManager
    {
        private object _lockObject = new object();
        private static LoggerManager _instance;

        public static LoggerManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new LoggerManager();
                }

                return _instance;
            }
        }

        private Dictionary<string, List<MessageBase>> logs;
        private LoggerManager()
        {
            logs = new Dictionary<string, List<MessageBase>>();
        }

        public void Add(MessageBase message)
        {
            lock (_lockObject)
            {
                List<MessageBase> documentMessages;
                if (!logs.ContainsKey(message.DocumentId))
                {
                    documentMessages = new List<MessageBase>();
                    logs.Add(message.DocumentId, documentMessages);
                }
                else
                {
                    documentMessages = logs[message.DocumentId];
                }

                documentMessages.Add(message);
            }
        }

        public IEnumerable<MessageBase> GetDocumentMessages(Document document)
        {
            lock (_lockObject)
            {
                if (logs.ContainsKey(document.Identifier))
                {
                    return logs[document.Identifier];
                }
                else
                {
                    return null;
                }
            }
        }

        public MessageBase GetLastDocumentMessage(Document document)
        {
            lock (_lockObject)
            {
                if (logs.ContainsKey(document.Identifier))
                {
                    return logs[document.Identifier].Last();
                }
                else
                {
                    return null;
                }
            }
        }

        public void ClearDocumentMessages(Document document)
        {
            lock (_lockObject)
            {
                if (logs.ContainsKey(document.Identifier))
                {
                    logs[document.Identifier].Clear();
                }
            }
        }

        public void Clear()
        {
            lock (_lockObject)
            {
                logs.Clear();
            }
        }

    }
}
