using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using ExcelReaderConsole.Logger;

namespace GUI.Models
{
    public class DocumentWindowModel : INotifyPropertyChanged
    {
        public DocumentWindowModel()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string windowHandler = string.Empty;
        public string WindowHandler 
        { 
            get => windowHandler;
            set
            {
                windowHandler = value;
                OnPropertyChanged("WindowHandler");
            }
        }

        private IEnumerable<LogMessageItem> logMessages;
        public IEnumerable<LogMessageItem> LogMessages
        {
            get { return logMessages; }
            set
            {
                logMessages = value;
                OnPropertyChanged("LogMessages");
            }
        }

        private IEnumerable<LogMessageItem> textMessages;
        public IEnumerable<LogMessageItem> TextMessages
        {
            get { return textMessages; }
            set
            {
                textMessages = value;
                OnPropertyChanged("TextMessages");
            }
        }

        private IEnumerable<LogMessageItem> scanMessages;
        public IEnumerable<LogMessageItem> ScanMessages
        {
            get { return scanMessages; }
            set
            {
                scanMessages = value;
                OnPropertyChanged("ScanMessages");
            }
        }

        private IEnumerable<LogMessageItem> textPdfMessages;
        public IEnumerable<LogMessageItem> TextPdfMessages
        {
            get { return textPdfMessages; }
            set
            {
                textPdfMessages = value;
                OnPropertyChanged("TextPdfMessages");
            }
        }

        private IEnumerable<LogMessageItem> attachMessages;
        public IEnumerable<LogMessageItem> AttachMessages
        {
            get { return attachMessages; }
            set
            {
                attachMessages = value;
                OnPropertyChanged("AttachMessages");
            }
        }
    }
}
