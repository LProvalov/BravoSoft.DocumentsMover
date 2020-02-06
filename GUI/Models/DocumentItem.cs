using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Media;
using ExcelReaderConsole.Models;

namespace GUI.Models
{
    public class DocumentItem : INotifyPropertyChanged
    {
        public enum DocumentStatus
        {
            Normal,
            WarningOccured,
            ErrorOccured
        }

        public enum DocumentState
        {
            Loaded = 0,
            Processed,
        }
        public DocumentItem(Document document, IEnumerable<DocumentAttribute> attributes)
        {
            Identifier = document.Identifier;
            TextFileName = document.TextFileName;
            ScanFileName = document.ScanFileName;
            TextPdfFileName = document.TextPdfFileName;
            AttachmentsFilesNames = document.AttachmentsFilesNames;

            attributeValues = new Dictionary<string, string>();
            foreach (var attribute in attributes)
            {
                if (document.GetValue(attribute.Identifier) != null)
                {
                    attributeValues.Add(attribute.Identifier,
                        document.GetValue(attribute.Identifier).Value.ToString());
                }
            }

            Status = DocumentStatus.Normal;
        }

        public string Identifier { get; set; }
        public string TextFileName { get; set; }
        public string ScanFileName { get; set; }
        public string TextPdfFileName { get; set; }
        public string AttachmentsFilesNames { get; set; }

        private Dictionary<string, string> attributeValues;
        /*
         * Get attribute value by attributeId
         */
        public string this[string attributeId]
        {
            get
            {
                if (attributeValues.ContainsKey(attributeId))
                {
                    return attributeValues[attributeId];
                }
                return string.Empty;
            }
        }

        public DocumentStatus Status { get; set; }
        private DocumentState state;

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public DocumentState State {
            get { return state; }
            set
            {
                state = value; 
                OnPropertyChanged("StatusBrush");
            }
        }

        public SolidColorBrush StatusBrush
        {
            get
            {
                if (State == DocumentState.Processed)
                {
                    return new SolidColorBrush(Colors.LightGreen);
                }
                return new SolidColorBrush(Colors.LightGray);
            }
        }
    }
}
