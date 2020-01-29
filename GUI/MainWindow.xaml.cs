using System;
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
            documentManager = DocumentManager.Instance;
            documentManager.StatusChanged += DocumentManagerStatusChangedHandler;
        }
        
        private void InitColumnToTemplateGridItem(IEnumerable<string> newColumnNames)
        {
            TemplateDataGrid.Columns.Clear();
            foreach (string name in newColumnNames)
            {
                TemplateDataGrid.Columns.Add(new DataGridTextColumn()
                {
                    Binding = new Binding($"TemplateDataGridRows[{name}]"),
                    Header = name
                });
            }
        }

        private void MenuItemLoadTemplate_OnClick(object sender, RoutedEventArgs e)
        {
            documentManager.ReadDataFromTemplate();
            RunEllipse.Fill = Brushes.GreenYellow;
            RunMenuItem.IsEnabled = true;

            var attributes= documentManager.DocumentsStorage.GetUsedDocumentAttributes();
            
            model.TemplateDataGridModel.TemplateDataGridColumns.Add(TemplateDataGridModel.AdditionalAttributes.TextFileAttribute);
            model.TemplateDataGridModel.TemplateDataGridColumns.Add(TemplateDataGridModel.AdditionalAttributes.ScanCopyAttribute);
            model.TemplateDataGridModel.TemplateDataGridColumns.Add(TemplateDataGridModel.AdditionalAttributes.TextPDFAttribute);
            model.TemplateDataGridModel.TemplateDataGridColumns.Add(TemplateDataGridModel.AdditionalAttributes.AttachmentsAttribute);
            model.TemplateDataGridModel.TemplateDataGridColumns.AddRange(attributes.Select(item => item.Name));
            InitColumnToTemplateGridItem(model.TemplateDataGridModel.TemplateDataGridColumns);

            var documents = documentManager.DocumentsStorage.GetDocuments();
            foreach (var document in documents)
            {
                var tgi = new TemplateGridItem();
                model.TemplateDataGridModel.TemplateGridItems.Add(tgi);
                if (!string.IsNullOrEmpty(document.TextFileName?.Trim()))
                {
                    tgi.TemplateDataGridRows.Add(TemplateDataGridModel.AdditionalAttributes.TextFileAttribute,
                        document.TextFileName);
                    tgi.TextFileExists = document.TextFileInfo?.Exists ?? false;
                }

                if (!string.IsNullOrEmpty(document.ScanFileName?.Trim()))
                {
                    tgi.TemplateDataGridRows.Add(TemplateDataGridModel.AdditionalAttributes.ScanCopyAttribute,
                        document.ScanFileName);
                    tgi.ScanFileExists = document.ScanFileInfo?.Exists ?? false;
                }

                if (!string.IsNullOrEmpty(document.TextPdfFileName?.Trim()))
                {
                    tgi.TemplateDataGridRows.Add(TemplateDataGridModel.AdditionalAttributes.TextPDFAttribute,
                        document.TextPdfFileName);
                    tgi.TextPDFFileExists = document.TextPdfFileInfo?.Exists ?? false;
                }

                if (!string.IsNullOrEmpty(document.AttachmentsFilesNames?.Trim()))
                {
                    tgi.TemplateDataGridRows.Add(TemplateDataGridModel.AdditionalAttributes.AttachmentsAttribute,
                        document.AttachmentsFilesNames);
                    tgi.AttachmentFiles = true;
                    if (document.AttachmentsFilesInfos != null)
                    {
                        foreach (var fileInfo in document.AttachmentsFilesInfos)
                        {
                            if ((fileInfo?.Exists ?? false) != true)
                            {
                                tgi.AttachmentFiles = false;
                                break;
                            }
                        }
                    }
                    else
                    {
                        tgi.AttachmentFiles = false;
                    }
                }

                foreach (var attribute in attributes)
                {
                    var documentAttrValue = document.GetValue(attribute.Identifier).Value;
                    tgi.TemplateDataGridRows.Add(attribute.Name, documentAttrValue);
                }
            }
            TemplateDataGrid.ItemsSource = model.TemplateDataGridModel.TemplateGridItems;
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
            if (newState == DocumentManager.State.TemplateLoaded)
            {
                TemplateDataGridTextMessage.Visibility = Visibility.Hidden;
            }
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
    }
}
