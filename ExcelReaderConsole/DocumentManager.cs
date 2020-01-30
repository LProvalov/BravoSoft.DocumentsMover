using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole
{
    public class DocumentManager
    {
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

        public EventHandler StatusStringChanged;
        public Action<State, State> StatusChanged;
        public Action<Document, StringBuilder> DocumentProcessingErrorOccured = null;
        public Func<Document, bool> OverwriteDocumentFiles = null;
        public Func<bool> OverwriteDocumentCards = null;

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

        private readonly StringBuilder errorLogBuilder = new StringBuilder();
        private readonly StringBuilder statusLogBuilder = new StringBuilder();
        
        private DocumentManager()
        {
            StatusChanged = (states, states1) => {};
            StatusStringChanged = (sender, args) => {};
            appSettings = AppSettings.Instance;
            appSettings.LoadAppSettings();
            ds = new DocumentsStorage();
            fileManager = FileManager.Instance;
            cardBuilder = CardBuilder.Instance;
            _State = State.InitializationDone;
        }

        public void ReadDataFromTemplate()
        {
            _State = State.TemplateLoading;
            try
            {
                string excelPath = appSettings.GetExcelTemplateFilePath();
                FileInfo excelFile = new FileInfo(excelPath);
                StatusStringChanged(this, new StatusStringChangedArgs($"Template is loading: {excelPath}"));
                if (excelFile.Exists && excelFile.Extension.Equals(".xlsx"))
                {
                    ds.Clean();
                    ExcelReader excelReader = new ExcelReader(excelPath);
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

        public void ProcessDocuments()
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

            int copyFilesErrorCount = 0;

            StatusStringChanged(this, new StatusStringChangedArgs("Documents are processing..."));
            foreach (var document in ds.GetDocuments())
            {
                string errorMessage;
                bool copyErrorOccured = false;

                string newTextFilePath = fileManager.MakeNewTextFilePath(document);
                string newScanFilePath = fileManager.MakeNewScanFilePath(document);
                string newTextPdfFilePath = fileManager.MakeNewTextPdfPath(document);

                bool overwrite = false;
                if (File.Exists(newTextFilePath) || File.Exists(newScanFilePath) || File.Exists(newTextPdfFilePath) ||
                    fileManager.NeedToOverwriteAttachments(document))
                {
                    overwrite = OverwriteDocumentFiles?.Invoke(document) ?? false;
                }

                fileManager.TryToCopyTextFile(document, out errorMessage, overwrite);
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    errorLogBuilder.AppendLine(errorMessage);
                    copyErrorOccured = true;
                }

                fileManager.TryToCopyScanFile(document, out errorMessage, overwrite);
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    errorLogBuilder.AppendLine(errorMessage);
                    copyErrorOccured = true;
                }

                fileManager.TryToCopyTextPdfFile(document, out errorMessage, overwrite);
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    errorLogBuilder.AppendLine(errorMessage);
                    copyErrorOccured = true;
                }

                fileManager.TryToCopyAttachmentFiles(document, out errorMessage, overwrite);
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    errorLogBuilder.AppendLine(errorMessage);
                    copyErrorOccured = true;
                }

                if (copyErrorOccured) copyFilesErrorCount++;

                cardBuilder.BuildCard(document.Identifier, document);
                if (document.CopiedScanFileInfo != null && document.CopiedScanFileInfo.Exists)
                {
                    cardBuilder.BuildAdditionalCard(document.Identifier, document);
                }

                if (document.CopiedAttachmentsFilesInfos != null && document.CopiedAttachmentsFilesInfos.Length > 0)
                {
                    for (int attachmentCount = 1; attachmentCount <= document.CopiedAttachmentsFilesInfos.Length; attachmentCount++)
                    {
                        cardBuilder.BuildAdditionalCardForAttachment($"Вложение{attachmentCount}", attachmentCount - 1, document);
                    }
                }

                if (copyFilesErrorCount > 0)
                {
                    DocumentProcessingErrorOccured?.Invoke(document, errorLogBuilder);
                }
            }
            StatusStringChanged(this, new StatusStringChangedArgs($"Documents processing finished. Errors: {copyFilesErrorCount}"));
            _State = State.DocumentMoved;
        }

        public DocumentsStorage DocumentsStorage
        {
            get { return ds; }
        }
    }
}
