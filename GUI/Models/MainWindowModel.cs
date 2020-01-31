using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ExcelReaderConsole.Models;
using GUI.Extensions;

namespace GUI.Models
{
    public class MainWindowModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private ConcurrentDictionary<string, DocumentItem> documents = new ConcurrentDictionary<string, DocumentItem>();
        public IList<DocumentItem> Documents => documents.Select(i => i.Value).ToList();

        public void SetDocuments(ConcurrentDictionary<string, DocumentItem> newDocuments)
        {
            documents = newDocuments;
            OnPropertyChanged("Documents");
        }

        public void SetDocumentProcessed(string documentId, bool isSuccessfully)
        {
            if (documents.ContainsKey(documentId))
            {
                documents[documentId].ProcessedSuccessfully = isSuccessfully;
                OnPropertyChanged("Documents");
            }
        }
       
    }
}
