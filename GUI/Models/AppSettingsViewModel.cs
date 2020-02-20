using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace GUI.Models
{
    public class AppSettingsViewModel : INotifyPropertyChanged
    {
        private String inputDirectory;

        public String InputDirectory
        {
            get { return this.inputDirectory; }
            set
            {
                this.inputDirectory = value;
                OnPropertyChanged("InputDirectory");
                OnPropertyChanged("IsExistsInputDirectory");
            }
        }

        private String outputDirectory;

        public bool IsExistsInputDirectory { get { return Directory.Exists(inputDirectory); } }
        public bool IsExistsOutputDirectory
        {
            get { return Directory.Exists(outputDirectory); }
        }

        public String OutputDirectory
        {
            get { return this.outputDirectory; }
            set
            {
                this.outputDirectory = value;
                OnPropertyChanged("OutputDirectory");
                OnPropertyChanged("IsExistsOutputDirectory");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
