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
                StatusStringChanged(this, new StatusStringChangedArgs($"Template is loading: {excelPath}"));
                ExcelReader excelReader = new ExcelReader(excelPath);
                excelReader.ReadData(ds);
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

            StatusStringChanged(this, new StatusStringChangedArgs("Documents are processing..."));
            foreach (var document in ds.GetDocuments())
            {
                string errorMessage;

                fileManager.TryToCopyTextFile(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) errorLogBuilder.AppendLine(errorMessage);

                fileManager.TryToCopyScanFile(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) errorLogBuilder.AppendLine(errorMessage);

                fileManager.TryToCopyTextPdfFile(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) errorLogBuilder.AppendLine(errorMessage);

                fileManager.TryToCopyAttachmentFiles(document, out errorMessage);
                if (!string.IsNullOrEmpty(errorMessage)) errorLogBuilder.AppendLine(errorMessage);

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
            }
            StatusStringChanged(this, new StatusStringChangedArgs($"Documents processing finished. Errors: {errorLogBuilder.Length}"));
            _State = State.DocumentMoved;
        }

        public DocumentsStorage DocumentsStorage
        {
            get { return ds; }
        }
    }
}
