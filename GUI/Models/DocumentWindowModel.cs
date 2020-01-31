using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace GUI.Models
{
    public class DocumentWindowModel : INotifyPropertyChanged
    {
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
    }
}
