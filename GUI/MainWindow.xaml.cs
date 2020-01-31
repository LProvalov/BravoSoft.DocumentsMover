using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelReaderConsole;
using ExcelReaderConsole.Models;
using GUI.Models;

namespace GUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly DocumentManager documentManager;

        public readonly MainWindowModel model;

        public MainWindow()
        {
            InitializeComponent();
            model = new MainWindowModel();
            DataContext = model;
            documentManager = DocumentManager.Instance;
            documentManager.StatusChanged += DocumentManagerStatusChangedHandler;
            documentManager.OverwriteDocumentFiles += OverwriteDocumentFiles;
            documentManager.OverwriteDocumentCards += OverwriteDocumentCards;
            documentManager.DocumentProcessed += DocumentProcessed;
        }

        private void DocumentProcessed(Document obj)
        {
            model.SetDocumentProcessed(obj.Identifier, true);
        }

        private bool OverwriteDocumentCards()
        {
            var result = MessageBox.Show($"Существуют катрочки для данного документа, перезаписать?",
                "Внимание!",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                return true;
            }
            return false;
        }

        private bool OverwriteDocumentFiles(Document document)
        {
            var result = MessageBox.Show($"Часть копируемых файлов документа ({document.Identifier}) уже существуют, перезаписать?", 
                "Внимание!",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                return true;
            }
            return false;
        }

        private void InitColumnToListView(IEnumerable<DocumentAttribute> attributes)
        {
            gridView.Columns.Clear();
            gridView.Columns.Add(new GridViewColumn()
            {
                Header = "Идентификатор",
                DisplayMemberBinding = new Binding($"Identifier")
            });
            gridView.Columns.Add(new GridViewColumn()
            {
                Header = "Текстовый файл",
                DisplayMemberBinding = new Binding($"TextFileName")
            });
            gridView.Columns.Add(new GridViewColumn()
            {
                Header = "Загрузить как скан копию",
                DisplayMemberBinding = new Binding($"ScanFileName")
            });
            gridView.Columns.Add(new GridViewColumn()
            {
                Header = "Текстовый PDF",
                DisplayMemberBinding = new Binding($"TextPdfFileName")
            });
            gridView.Columns.Add(new GridViewColumn()
            {
                Header = "Вложения",
                DisplayMemberBinding = new Binding($"AttachmentsFilesNames")
            });
            foreach (var attribute in attributes)
            {
                gridView.Columns.Add(new GridViewColumn()
                {
                    Header = attribute.Name,
                    DisplayMemberBinding = new Binding($"[{attribute.Identifier}]")
                });
            }
        }

        private string selectExcelTemplateFile()
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel template file (*.excel)|*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            bool? dialogResult = openFileDialog.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value == true)
            {
                return openFileDialog.FileName;
            }
            return null;
        }

        private void MenuItemLoadTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            string templateFile = selectExcelTemplateFile();
            if (!string.IsNullOrEmpty(templateFile))
            {
                ExcelReaderConsole.AppSettings.Instance.ExcelTemplateFilePath = templateFile;
            }
            else
            {
                return;
            }
            documentManager.ReadDataFromTemplate();
            RunEllipse.Fill = Brushes.GreenYellow;
            RunMenuItem.IsEnabled = true;

            var attributes = documentManager.DocumentsStorage.GetUsedDocumentAttributes();
            InitColumnToListView(attributes);
            
            var documentItems = new ConcurrentDictionary<string, DocumentItem>();
            foreach (var document in documentManager.DocumentsStorage.GetDocuments())
            {
                documentItems.AddOrUpdate(document.Identifier, 
                    (id) => new DocumentItem(document, attributes),
                    ((s, item) => new DocumentItem(document, attributes)));
            }

            model.SetDocuments(documentItems);
        }

        private void MenuItemSettings_OnClick(object sender, RoutedEventArgs e)
        {
            AppSettings appSettings = new AppSettings();
            appSettings.Left = this.Left + this.Width / 2 - appSettings.Width / 2;
            appSettings.Top = this.Top + this.Height / 2 - appSettings.Height / 2;
            bool? dialogResult = appSettings.ShowDialog();
        }

        private void MenuItemExit_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MenuItemRun_OnClick(object sender, RoutedEventArgs e)
        {
            documentManager.ProcessDocuments();
            MessageBox.Show("Document processing is done.", "Processing result", MessageBoxButton.OK);
        }

        private void DocumentManagerStatusChangedHandler(DocumentManager.State oldState,
            DocumentManager.State newState)
        {
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            var messageBoxResult = MessageBox.Show("Are you sure want to close application?", "Exit",
                MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (messageBoxResult == MessageBoxResult.Yes)
            {
                base.OnClosing(e);
            }
            else
            {
                e.Cancel = true;
            }
        }
        
        private void listView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var index = listView.SelectedIndex;
            var selectedItem = listView.Items[index] as DocumentItem;
            var document = documentManager.DocumentsStorage.GetDocument(selectedItem.Identifier);
            
            // TODO: prepare documentWindowModel and put it to documentWindow constructor
            DocumentWindowModel dwModel = new DocumentWindowModel();
            dwModel.WindowHandler = document.Identifier;
            DocumentWindow documentWindow = new DocumentWindow(dwModel);

            documentWindow.Left = this.Left + this.Width / 2 - documentWindow.Width / 2;
            documentWindow.Top = this.Top + this.Height / 2 - documentWindow.Height / 2;
            documentWindow.ShowDialog();
        }
    }
}
