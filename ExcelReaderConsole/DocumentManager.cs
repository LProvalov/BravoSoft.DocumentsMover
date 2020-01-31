using System;
using System.IO;
using System.Text;
using ExcelReaderConsole.Logger;
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

        public Action<Document> DocumentProcessed;
        public EventHandler StatusStringChanged;
        public Action<State, State> StatusChanged;

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
        public readonly Logger.LoggerManager loggerManager;
        
        private DocumentManager()
        {
            StatusChanged = (states, states1) => {};
            StatusStringChanged = (sender, args) => {};
            appSettings = AppSettings.Instance;
            appSettings.LoadAppSettings();
            ds = new DocumentsStorage();
            fileManager = FileManager.Instance;
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

            StatusStringChanged(this, new StatusStringChangedArgs("Documents are processing..."));
            foreach (var document in ds.GetDocuments())
            {
                string errorMessage;

                string newTextFilePath = fileManager.MakeNewTextFilePath(document);
                string newScanFilePath = fileManager.MakeNewScanFilePath(document);
                string newTextPdfFilePath = fileManager.MakeNewTextPdfPath(document);

                bool overwrite = false;
                bool tfp = File.Exists(newTextFilePath);
                bool sfp = File.Exists(newScanFilePath);
                bool tpfp = File.Exists(newTextPdfFilePath);
                bool ato = fileManager.NeedToOverwriteAttachments(document);
                if ( tfp || sfp || tpfp || ato)
                {
                    overwrite = OverwriteDocumentFiles?.Invoke(document) ?? false;
                    loggerManager.Add(new LogMessage(document, $"TextFile exists: {tfp}"));
                    loggerManager.Add(new LogMessage(document, $"TextScanFile exists: {sfp}"));
                    loggerManager.Add(new LogMessage(document, $"TextPdfFile exists: {tpfp}"));
                    loggerManager.Add(new LogMessage(document, $"Attachments need to overwrite: {ato}"));
                    loggerManager.Add(new LogMessage(document, $"Существующие файлы будут перезаписаны: {overwrite}"));
                }

                if (!string.IsNullOrEmpty(document.TextFileName))
                {
                    fileManager.TryToCopyTextFile(document, out errorMessage, overwrite);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                if (!string.IsNullOrEmpty(document.ScanFileName))
                {
                    fileManager.TryToCopyScanFile(document, out errorMessage, overwrite);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                if (!string.IsNullOrEmpty(document.TextPdfFileName))
                {
                    fileManager.TryToCopyTextPdfFile(document, out errorMessage, overwrite);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                if (!string.IsNullOrEmpty(document.AttachmentsFilesNames))
                {
                    fileManager.TryToCopyAttachmentFiles(document, out errorMessage, overwrite);
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        loggerManager.Add(new ErrorMessage(document, errorMessage));
                    }
                }

                bool overwriteCard = false;
                overwrite |= File.Exists(cardBuilder.GetCardPath(document.Identifier));
                overwrite |= File.Exists(cardBuilder.GetAdditionalCardName(document.Identifier, document));
                if (document.CopiedAttachmentsFilesInfos != null && document.CopiedAttachmentsFilesInfos.Length > 0)
                {
                    for (int attachmentCount = 1;
                        attachmentCount <= document.CopiedAttachmentsFilesInfos.Length;
                        attachmentCount++)
                    {
                        overwrite |=
                            File.Exists(
                                cardBuilder.GetAdditionalCardNameForAttachment($"Вложение{attachmentCount}", document));
                    }
                }

                if (overwriteCard)
                {
                    overwriteCard = OverwriteDocumentCards?.Invoke() ?? false;
                }

                cardBuilder.BuildCard(document.Identifier, document, overwriteCard);
                if (document.CopiedScanFileInfo != null && document.CopiedScanFileInfo.Exists)
                {
                    cardBuilder.BuildAdditionalCard(document.Identifier, document, overwriteCard);
                }

                if (document.CopiedAttachmentsFilesInfos != null && document.CopiedAttachmentsFilesInfos.Length > 0)
                {
                    for (int attachmentCount = 1; attachmentCount <= document.CopiedAttachmentsFilesInfos.Length; attachmentCount++)
                    {
                        cardBuilder.BuildAdditionalCardForAttachment($"Вложение{attachmentCount}", attachmentCount - 1, document, overwriteCard);
                    }
                }
                DocumentProcessed?.Invoke(document);
            }
            StatusStringChanged(this, new StatusStringChangedArgs($"Documents processing finished."));
            _State = State.DocumentMoved;
        }

        public DocumentsStorage DocumentsStorage
        {
            get { return ds; }
        }
    }
}
