using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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
                documents[documentId].State = DocumentItem.DocumentState.Processed;
            }
        }

        public string statusMessage;

        public string StatusMessage
        {
            get => statusMessage;
            set
            {
                statusMessage = value;
                OnPropertyChanged("StatusMessage");
            }
        }

        private ImageSource imageRunSource = null;
        public ImageSource ImageRunSource { get { return imageRunSource; } }

        public void SetImageRunSource(ImageSource imageSource)
        {
            if (imageSource != null)
            {
                imageRunSource = imageSource;
                OnPropertyChanged("ImageRunSource");
            }
        }
    }
}
