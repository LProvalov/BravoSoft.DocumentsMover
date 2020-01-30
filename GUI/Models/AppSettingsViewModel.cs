using System;
using System.Collections.Generic;
using System.ComponentModel;
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
            }
        }

        private String outputDirectory;

        public String OutputDirectory
        {
            get { return this.outputDirectory; }
            set
            {
                this.outputDirectory = value;
                OnPropertyChanged("OutputDirectory");
            }
        }

        private String excelTemplatePath;

        public String ExcelTemplatePath
        {
            get { return this.excelTemplatePath; }
            set
            {
                this.excelTemplatePath = value;
                OnPropertyChanged("ExcelTemplatePath");
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
