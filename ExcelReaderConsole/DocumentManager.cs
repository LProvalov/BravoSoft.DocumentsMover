using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelReaderConsole.Logger;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole
{
    public class DocumentManager
    {
        public enum ValidateStatus
        {
            Success = 0,
            Warning = 1,
            Error = 2
        }

        private static DocumentManager _instance = null;
        public static DocumentManager Instance
        {
            get { return _instance ?? (_instance = new DocumentManager()); }
        }
        public enum State
        {
            InitializationDone = 0,
            TemplateLoading,
            TemplateLoaded,
            DocumentProcessing,
            DocumentMoved,
            ErrorOccured,
        }

        public Action<Document> DocumentProcessed;
        public EventHandler StatusStringChanged;
        public Action<State, State> StatusChanged;
        public Action<Exception> ExceptionOccured;

        public Func<Document, bool> OverwriteDocumentFiles = null;
        public Func<bool> OverwriteDocumentCards = null;
        public Func<string, bool> OverwriteFile = null;

        public class StatusStringChangedArgs : EventArgs
        {
            public string Message { get; protected set; }
            public StatusStringChangedArgs(string msg)
            {
                Message = msg;
            }
        }

        private State _state;
        protected State _State
        {
            set
            {
                State old = _state;
                _state = value;
                StatusChanged(old, value);
            }
            get { return _state; }
        }
        private readonly AppSettings appSettings;
        private readonly DocumentsStorage ds;
        private readonly FileManager fileManager;
        private readonly CardBuilder cardBuilder;
        public readonly Logger.LoggerManager loggerManager;

        private float procentProcessed = 0f;
        private int countOfDocumentProcessed = 0;

        public float ProcentProcessed
        {
            get { return procentProcessed; }
        }

        private DocumentManager()
        {
            StatusChanged = (states, states1) => { };
            StatusStringChanged = (sender, args) => { };
            appSettings = AppSettings.Instance;
            appSettings.LoadAppSettings();
            ds = new DocumentsStorage();
            fileManager = FileManager.Instance;
            fileManager.FileExistOverwrite = s => OverwriteFile?.Invoke(s) ?? false;
            cardBuilder = CardBuilder.Instance;
            loggerManager = LoggerManager.Instance;
            _State = State.InitializationDone;
        }

        public void ReadDataFromTemplate()
        {
            _State = State.TemplateLoading;
            try
            {
                if (!appSettings.Validate(out string errorMsg))
                {

                    _State = State.ErrorOccured;
                    return;
                }
                string excelPath = appSettings.GetExcelTemplateFilePath();
                FileInfo excelFile = new FileInfo(excelPath);
                StatusStringChanged(this, new StatusStringChangedArgs($"Template is loading: {excelPath}"));
                if (excelFile.Exists && excelFile.Extension.Equals(".xlsx"))
                {
                    ds.Clean();
                    ExcelReader excelReader = new ExcelReader(excelPath);
                    excelReader.ExceptionOccured += exception =>
                        {
                            ExceptionOccured?.BeginInvoke(exception, null, null);
                        };
                    excelReader.ReadData(ds);
                }
                else
                {
                    throw new Exception($"Файл шаблона ({excelPath}) не существует или его раширение ({excelFile.Extension}) не .xlsx");
                }
            }
            catch (Exception ex)
            {
                _State = State.ErrorOccured;
                StatusStringChanged(this, new StatusStringChangedArgs("Error occured"));
                throw ex;
            }
            StatusStringChanged(this, new StatusStringChangedArgs("Template loaded successfully"));
            _State = State.TemplateLoaded;
        }

        public void ReadDataFromTemplateAsync(Action<Task> continueWith)
        {
            try
            {
                var readDataTask = Task.Factory.StartNew(action: () => { ReadDataFromTemplate(); });
                readDataTask.ContinueWith(continueWith, TaskContinuationOptions.NotOnFaulted);
                readDataTask.ContinueWith((task) => {
                    ExceptionOccured?.BeginInvoke(task.Exception, null, null);
                }, TaskContinuationOptions.OnlyOnFaulted);
            }
            catch(Exception ex)
            {
                ExceptionOccured?.BeginInvoke(ex, null, null);
            }
        }

        private void ProcessDocuments()
        {
            if (_State != State.TemplateLoaded)
            {
                StatusStringChanged(this,
                    new StatusStringChangedArgs("Can't process documents if template isn't loaded"));
                return;
            }
            _State = State.DocumentProcessing;

            StatusStringChanged(this, new StatusStringChangedArgs("Checking of output directories"));
            DirectoryInfo outputDirectory = new DirectoryInfo(appSettings.GetOutputDirectoryPath());
            if (!outputDirectory.Exists)
            {
                outputDirectory.Create();
            }
            cardBuilder.SetOutputDirectory(outputDirectory.FullName);

            StatusStringChanged(this, new StatusStringChangedArgs("Documents are processing..."));
            procentProcessed = 0f;
            countOfDocumentProcessed = 0;
            var documents = ds.GetDocuments();
            var documentNumber = documents.Count();
            foreach (var document in documents)
            {
                string errorMessage;
                if (!string.IsNullOrEmpty(document.TextFileName))
                {
                    loggerManager.Add(new LogMessage(document, $"Копирование текстового файла"));
                    fileManager.TryToCopyTextFile(document, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                if (!string.IsNullOrEmpty(document.ScanFileName))
                {
                    loggerManager.Add(new LogMessage(document, $"Копирование скан файла"));
                    fileManager.TryToCopyScanFile(document, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                if (!string.IsNullOrEmpty(document.TextPdfFileName))
                {
                    loggerManager.Add(new LogMessage(document, $"Копирование текстового Pdf файла"));
                    fileManager.TryToCopyTextPdfFile(document, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                if (!string.IsNullOrEmpty(document.AttachmentsFilesNames))
                {
                    loggerManager.Add(new LogMessage(document, $"Копирование файлов вложений"));
                    fileManager.TryToCopyAttachmentFiles(document, out errorMessage);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                loggerManager.Add(new LogMessage(document, $"Создание карточки документа"));
                cardBuilder.BuildCard(document.Identifier, document);
                if (document.CopiedScanFileInfo != null && document.CopiedScanFileInfo.Exists)
                {
                    loggerManager.Add(new LogMessage(document, $"Создание дополнительной карточки документа"));
                    cardBuilder.BuildAdditionalCard(document.Identifier, document);
                }

                if (document.CopiedAttachmentsFilesInfos != null && document.CopiedAttachmentsFilesInfos.Length > 0)
                {
                    loggerManager.Add(new LogMessage(document, $"Создание карточек вложений"));
                    for (int attachmentCount = 1; attachmentCount <= document.CopiedAttachmentsFilesInfos.Length; attachmentCount++)
                    {
                        loggerManager.Add(new LogMessage(document, $"Вложение{attachmentCount}"));
                        cardBuilder.BuildAdditionalCardForAttachment($"Вложение{attachmentCount}", attachmentCount - 1, document);
                    }
                }
                loggerManager.Add(new LogMessage(document, $"Обработка документа завершена"));
                countOfDocumentProcessed++;
                procentProcessed = countOfDocumentProcessed / (float) documentNumber;
                DocumentProcessed?.Invoke(document);
            }
            StatusStringChanged(this, new StatusStringChangedArgs($"Documents processing finished."));
            _State = State.DocumentMoved;
        }

        public Task ProcessDocumentsAsync()
        {
            return Task.Factory.StartNew(action: () => { ProcessDocuments(); });
        }

        public int ValidateDocument(Document document)
        {
            int returnStatus = 0;
            if (!string.IsNullOrEmpty(document.TextFileName))
            {
                if (!document.TextFileInfo.Exists)
                {
                    returnStatus |= (int)ValidateStatus.Error;
                    loggerManager.Add(new WarningMessage(document,
                        $"Текстовый файл {document.TextFileName} не существует во входящей директории.",
                        WarningMessage.PlaceType.TextFile));
                }

                string newFullFilePath = fileManager.MakeNewTextFilePath(document);
                if (File.Exists(newFullFilePath))
                {
                    returnStatus |= (int)ValidateStatus.Warning;
                    loggerManager.Add(new WarningMessage(document,
                        $"Файл документа {newFullFilePath} уже существует в исходящей директории.",
                        WarningMessage.PlaceType.TextFile));
                }
            }

            if (!string.IsNullOrEmpty(document.ScanFileName))
            {
                if (!document.ScanFileInfo.Exists)
                {
                    returnStatus |= (int)ValidateStatus.Error;
                    loggerManager.Add(new WarningMessage(document,
                        $"Скан файл {document.ScanFileName} не существует во входящей директории.",
                        WarningMessage.PlaceType.Scan));
                }

                string newFullFilePath = fileManager.MakeNewScanFilePath(document);
                if (File.Exists(newFullFilePath))
                {
                    returnStatus |= (int)ValidateStatus.Warning;
                    loggerManager.Add(new WarningMessage(document,
                        $"Файл документа {newFullFilePath} уже существует в исходящей директории.",
                        WarningMessage.PlaceType.Scan));
                }
            }

            if (!string.IsNullOrEmpty(document.TextPdfFileName))
            {
                if (!document.TextPdfFileInfo.Exists)
                {
                    returnStatus |= (int)ValidateStatus.Error;
                    loggerManager.Add(new WarningMessage(document,
                        $"Скан файл {document.TextPdfFileName} не существует во входящей директории.",
                        WarningMessage.PlaceType.TextPdf));
                }

                string newFullFilePath = fileManager.MakeNewTextPdfPath(document);
                if (File.Exists(newFullFilePath))
                {
                    returnStatus |= (int)ValidateStatus.Warning;
                    loggerManager.Add(new WarningMessage(document,
                        $"Файл документа {newFullFilePath} уже существует в исходящей директории.",
                        WarningMessage.PlaceType.TextPdf));
                }
            }

            if (!string.IsNullOrEmpty(document.AttachmentsFilesNames))
            {
                try
                {
                    foreach (FileInfo attachFileInfo in fileManager.GetAttachmentFileInfos(document))
                    {
                        if (!attachFileInfo.Exists)
                        {
                            returnStatus |= (int)ValidateStatus.Error;
                            loggerManager.Add(new WarningMessage(document,
                                $"Дополнительный файл {attachFileInfo.Name} не существует во входящей директории.",
                                WarningMessage.PlaceType.Attachments));
                        }
                    }
                }
                catch
                {
                    returnStatus |= (int)ValidateStatus.Error;
                    loggerManager.Add(new WarningMessage(document,
                        $"Произошла ошибка при обработке строки {document.AttachmentsFilesNames}",
                        WarningMessage.PlaceType.Attachments));
                }
            }
            loggerManager.Add(new LogMessage(document, "Валидация файлов документа завершена."));
            return returnStatus;
        }

        public DocumentsStorage DocumentsStorage
        {
            get { return ds; }
        }
    }
}
