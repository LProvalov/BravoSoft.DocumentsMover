using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace GUI.Models
{
    public class LogMessageItem : INotifyPropertyChanged
    {
        public enum LogStatus
        {
            Normal,
            Warning,
            Error
        }

        public LogMessageItem(string message, string time, LogStatus status = LogStatus.Normal)
        {
            this.message = message;
            this.time = time;
            this.status = status;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string time;
        public string Time
        {
            get => time;
            set
            {
                time = value;
                OnPropertyChanged("Time");
            }
        }

        private LogStatus status;
        public LogStatus Status
        {
            get => status;
            set
            {
                status = value;
                OnPropertyChanged("Status");
            }
        }

        private string message;
        public string Message
        {
            get => message;
            set
            {
                message = value;
                OnPropertyChanged("Message");
            }
        }
    }
}
