using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace GUI.Models
{
    public class ExceptionHandlingWindowModel : INotifyPropertyChanged
    {
        private readonly object _syncObject = new object();
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private readonly IList<string> messages = new List<string>();
        private readonly IList<Exception> exceptions = new List<Exception>();

        public IEnumerable<string> Messages
        {
            get { return messages; }
        }

        public IEnumerable<Exception> Exceptions
        {
            get { return exceptions; }
        }

        public void AddMessage(string msg)
        {
            lock (_syncObject)
            {
                messages.Add(msg);
                OnPropertyChanged("Messages");
            }
        }

        public void AddException(Exception ex)
        {
            lock (_syncObject)
            {
                exceptions.Add(ex);
                OnPropertyChanged("Exceptions");
            }
        }
    }
}
