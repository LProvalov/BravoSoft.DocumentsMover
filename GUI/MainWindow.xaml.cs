using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.VisualStyles;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using ExcelReaderConsole;
using ExcelReaderConsole.Logger;
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
        private delegate void NoArgDelegate();
        private Dictionary<string, ImageSource> runIconsList = new Dictionary<string, ImageSource>();
        private bool templateWasLoaded = false;

        public MainWindow()
        {
            InitializeComponent();
            runIconsList.Add("start", new BitmapImage(new Uri("pack://application:,,,/Icons/start.png")));
            runIconsList.Add("stop", new BitmapImage(new Uri("pack://application:,,,/Icons/stop.png")));
            runIconsList.Add("stop_yellow", new BitmapImage(new Uri("pack://application:,,,/Icons/stop_yellow.png")));

            model = new MainWindowModel();
            model.SetImageRunSource(runIconsList["stop"]);
            model.statusMessage = "Ready to work";
            DataContext = model;
            documentManager = DocumentManager.Instance;
            documentManager.StatusChanged += DocumentManagerStatusChangedHandler;
            documentManager.OverwriteDocumentFiles += OverwriteDocumentFiles;
            documentManager.OverwriteDocumentCards += OverwriteDocumentCards;
            documentManager.DocumentProcessed += DocumentProcessed;
            documentManager.OverwriteFile += OverwriteFile;
            documentManager.ExceptionOccured += DocumentManagerExceptionOccured;
        }

        private void DocumentManagerExceptionOccured(Exception ex)
        {
            Dispatcher?.BeginInvoke(DispatcherPriority.Input, new NoArgDelegate(() => {
                ExceptionHandling.Instance.Model.AddException(ex);
                ExceptionHandling.Instance.Show();
            }));
        }

        private bool OverwriteFile(string fullFileName)
        {
            var result = MessageBox.Show($"{fullFileName} уже существует, перезаписать?",
                "Внимание!",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                return true;
            }
            return false;
        }

        private void DocumentProcessed(Document obj)
        {
            model.SetDocumentProcessed(obj.Identifier, true);
            model.StatusMessage = $"{documentManager.ProcentProcessed:P} completed";
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
            gridView.Dispatcher?.BeginInvoke(new NoArgDelegate(() =>
            {
                gridView.Columns.Clear();
                var factory = new FrameworkElementFactory(typeof(Ellipse));
                factory.SetValue(Ellipse.FillProperty, new Binding("StatusBrush"));
                factory.SetValue(Ellipse.WidthProperty, 10.0);
                factory.SetValue(Ellipse.HeightProperty, 10.0);

                gridView.Columns.Add(new GridViewColumn()
                {
                    Header = "Состояние",
                    CellTemplate = new DataTemplate
                    {
                        VisualTree = factory
                    }
                });
                var identifierGridView = new GridViewColumn()
                {
                    Header = "Идентификатор",
                    DisplayMemberBinding = new Binding($"Identifier")
                };
                gridView.Columns.Add(identifierGridView);
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

            }), DispatcherPriority.Input);
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

        private void LoadTemplate()
        {
            templateWasLoaded = false;
            model.SetImageRunSource(runIconsList["stop"]);
            RunMenuItem.IsEnabled = false;
            textBoxListViewLable.Text = "Loading pattern...";
            textBoxListViewLable.Visibility = Visibility.Visible;

            documentManager.ReadDataFromTemplateAsync((task) =>
            {
                var attributes = documentManager.DocumentsStorage.GetUsedDocumentAttributes();
                InitColumnToListView(attributes);

                var documentItems = new ConcurrentDictionary<string, DocumentItem>();
                foreach (var document in documentManager.DocumentsStorage.GetDocuments())
                {
                    var newDocumentItem = new DocumentItem(document, attributes);
                    int validateStatus = documentManager.ValidateDocument(document);
                    if (0 != (validateStatus & (int)DocumentManager.ValidateStatus.Warning))
                    {
                        newDocumentItem.Status = DocumentItem.DocumentStatus.WarningOccured;
                    }

                    if (0 != (validateStatus & (int)DocumentManager.ValidateStatus.Error))
                    {
                        newDocumentItem.Status = DocumentItem.DocumentStatus.ErrorOccured;
                    }

                    newDocumentItem.State = DocumentItem.DocumentState.Loaded;
                    documentItems.AddOrUpdate(document.Identifier,
                                              (id) => newDocumentItem,
                                              ((s, item) => newDocumentItem));
                }

                gridView.Dispatcher?.BeginInvoke(new NoArgDelegate(() =>
                {
                    textBoxListViewLable.Visibility = Visibility.Hidden;
                    model.SetDocuments(documentItems);
                    model.SetImageRunSource(runIconsList["start"]);
                    RunMenuItem.IsEnabled = true;
                    var view = (CollectionView)CollectionViewSource.GetDefaultView(listView.ItemsSource);
                    view.SortDescriptions.Add(new SortDescription("Identifier", ListSortDirection.Ascending));

                }), DispatcherPriority.Input);
                templateWasLoaded = true;
            });
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
            LoadTemplate();
        }

        private void MenuItemSettings_OnClick(object sender, RoutedEventArgs e)
        {
            AppSettings appSettings = new AppSettings();
            appSettings.Left = this.Left + this.Width / 2 - appSettings.Width / 2;
            appSettings.Top = this.Top + this.Height / 2 - appSettings.Height / 2;
            bool? dialogResult = appSettings.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value && templateWasLoaded)
            {
                // In this case we should reload template
                LoadTemplate();
            }
        }

        private void MenuItemExit_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MenuItemRun_OnClick(object sender, RoutedEventArgs e)
        {
            menuItemLoadTemplate.IsEnabled = false;
            model.SetImageRunSource(runIconsList["stop_yellow"]);
            RunMenuItem.IsEnabled = false;

            model.StatusMessage = $"{0:P} completed";
            documentManager.ProcessDocumentsAsync().ContinueWith((task =>
            {
                MessageBox.Show("Document processing is done.", "Processing result",
                                MessageBoxButton.OK);
                gridView.Dispatcher?.BeginInvoke(new NoArgDelegate(() =>
                {
                    model.SetImageRunSource(runIconsList["stop"]);
                    RunMenuItem.IsEnabled = false;
                    menuItemLoadTemplate.IsEnabled = true;
                }), DispatcherPriority.Input);
            }));
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
            int index = listView.SelectedIndex;
            if (index < 0 || index >= listView.Items.Count)
            {
                return;
            }

            if (listView.Items[index] is DocumentItem selectedItem)
            {
                Document document = documentManager.DocumentsStorage.GetDocument(selectedItem.Identifier);

                DocumentWindowModel dwModel = new DocumentWindowModel
                {
                    WindowHandler = document.Identifier,
                    LogMessages = documentManager.loggerManager.GetDocumentMessages(document)
                        .Where(msg => msg is LogMessage || msg is ErrorMessage).Select(msg =>
                        {
                            LogMessageItem.LogStatus ls = LogMessageItem.LogStatus.Normal;
                            if (msg.Type == MessageBase.MessageType.warning) ls = LogMessageItem.LogStatus.Warning;
                            else if (msg.Type == MessageBase.MessageType.error) ls = LogMessageItem.LogStatus.Error;
                            return new LogMessageItem(msg.Message, msg.Time.ToShortTimeString(), ls);
                        }),
                    TextMessages = documentManager.loggerManager.GetDocumentMessages(document).Where(msg =>
                        {
                            if (msg is WarningMessage)
                            {
                                return (msg as WarningMessage).Place == WarningMessage.PlaceType.TextFile;
                            }
                            return false;
                        })
                        .Select(msg =>
                        {
                            LogMessageItem.LogStatus ls = LogMessageItem.LogStatus.Normal;
                            if (msg.Type == MessageBase.MessageType.warning) ls = LogMessageItem.LogStatus.Warning;
                            else if (msg.Type == MessageBase.MessageType.error) ls = LogMessageItem.LogStatus.Error;
                            return new LogMessageItem(msg.Message, msg.Time.ToShortTimeString(), ls);
                        }),
                    ScanMessages = documentManager.loggerManager.GetDocumentMessages(document).Where(msg =>
                        {
                            if (msg is WarningMessage)
                            {
                                return (msg as WarningMessage).Place == WarningMessage.PlaceType.Scan;
                            }
                            return false;
                        })
                        .Select(msg =>
                        {
                            LogMessageItem.LogStatus ls = LogMessageItem.LogStatus.Normal;
                            if (msg.Type == MessageBase.MessageType.warning) ls = LogMessageItem.LogStatus.Warning;
                            else if (msg.Type == MessageBase.MessageType.error) ls = LogMessageItem.LogStatus.Error;
                            return new LogMessageItem(msg.Message, msg.Time.ToShortTimeString(), ls);
                        }),
                    TextPdfMessages = documentManager.loggerManager.GetDocumentMessages(document).Where(msg =>
                        {
                            if (msg is WarningMessage)
                            {
                                return (msg as WarningMessage).Place == WarningMessage.PlaceType.TextPdf;
                            }
                            return false;
                        })
                        .Select(msg =>
                        {
                            LogMessageItem.LogStatus ls = LogMessageItem.LogStatus.Normal;
                            if (msg.Type == MessageBase.MessageType.warning) ls = LogMessageItem.LogStatus.Warning;
                            else if (msg.Type == MessageBase.MessageType.error) ls = LogMessageItem.LogStatus.Error;
                            return new LogMessageItem(msg.Message, msg.Time.ToShortTimeString(), ls);
                        }),
                    AttachMessages = documentManager.loggerManager.GetDocumentMessages(document).Where(msg =>
                        {
                            if (msg is WarningMessage)
                            {
                                return (msg as WarningMessage).Place == WarningMessage.PlaceType.Attachments;
                            }
                            return false;
                        })
                        .Select(msg =>
                        {
                            LogMessageItem.LogStatus ls = LogMessageItem.LogStatus.Normal;
                            if (msg.Type == MessageBase.MessageType.warning) ls = LogMessageItem.LogStatus.Warning;
                            else if (msg.Type == MessageBase.MessageType.error) ls = LogMessageItem.LogStatus.Error;
                            return new LogMessageItem(msg.Message, msg.Time.ToShortTimeString(), ls);
                        })
                };
                DocumentWindow documentWindow = new DocumentWindow(dwModel);

                documentWindow.Left = this.Left + this.Width / 2 - documentWindow.Width / 2;
                documentWindow.Top = this.Top + this.Height / 2 - documentWindow.Height / 2;
                documentWindow.ShowDialog();
            }
        }
    }
}
