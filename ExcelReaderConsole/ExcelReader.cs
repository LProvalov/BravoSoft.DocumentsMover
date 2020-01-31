using System;
using System.IO;
using System.Text;
using ExcelDataReader;
using ExcelReaderConsole.Models;

namespace ExcelReaderConsole
{
    public class ExcelReader
    {
        private const int HeaderCount = 3;

        private FileInfo excelFileInfo;
        private bool isDataRead = false;
        public ExcelReader(string filePath)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new Exception("Excel file path is null or empty");
            }
            excelFileInfo = new FileInfo(filePath);
            if (!excelFileInfo.Exists || excelFileInfo.Extension != ".xlsx")
            {
                throw new Exception("File not exists or have invalid extension");
            }
        }

        private void ReadColumn(IExcelDataReader reader, int colPosition, out object value, out Type valueType)
        {
            switch (Type.GetTypeCode(reader.GetFieldType(colPosition)))
            {
                case TypeCode.String:
                    {
                        valueType = typeof(string);
                        value = reader.GetString(colPosition) as string;
                    }
                    break;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Double:
                    {
                        valueType = typeof(double);
                        value = reader.GetDouble(colPosition);
                    }
                    break;
                default:
                    {
                        valueType = typeof(object);
                        value = null;
                    }
                    break;
            }
        }

        private string ReadTextFieldFromExcelReader(IExcelDataReader reader,
            Document document, int colPosition, DocumentsStorage.FilesType filesType)
        {
            ReadColumn(reader, colPosition, out object textFromReader, out Type textFromReaderType);
            if (textFromReaderType == typeof(string) && !string.IsNullOrEmpty(textFromReader as string))
            {
                if (filesType == DocumentsStorage.FilesType.Text)
                {
                    document.TextFileName = textFromReader as string;
                }
                else if (filesType == DocumentsStorage.FilesType.ScanCopy)
                {
                    document.ScanFileName = textFromReader as string;
                    document.SetAttributeValue("7860:", new DocumentAttributeValue(textFromReader as string, typeof(string)));
                }
                else if (filesType == DocumentsStorage.FilesType.TextPdf)
                {
                    document.TextPdfFileName = textFromReader as string;
                }
                else if (filesType == DocumentsStorage.FilesType.Attachments)
                {
                    document.AttachmentsFilesNames = textFromReader as string;
                }
                return document.Identifier;
            }
            return string.Empty;
        }

        public void ReadData(DocumentsStorage documentStorage)
        {
            using (var stream = excelFileInfo.OpenRead())
            {
                var configuration = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.Default
                };
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream, configuration))
                {
                    do
                    {
                        int rowCount = reader.RowCount;
                        int workingRow = 0;

                        // + Work with header + 
                        int attributePositionStart = 5;
                        int fieldCount = reader.FieldCount - attributePositionStart;                        
                        documentStorage.Init(fieldCount);

                        reader.Read();
                        for (int attributeNumber = 0; attributeNumber < fieldCount; attributeNumber++)
                        {
                            documentStorage.SetAttributeName(attributeNumber, reader.GetString(attributeNumber + attributePositionStart));
                        }
                        workingRow++;

                        reader.Read(); reader.Read();
                        for (int attributeNumber = 0; attributeNumber < fieldCount; attributeNumber++)
                        {
                            documentStorage.SetAttributeIdentifier(attributeNumber, reader.GetString(attributeNumber + attributePositionStart));
                        }
                        // - Work with header -

                        int documentCount = 1;
                        while (reader.Read())
                        {
                            workingRow++;

                            string documentIdentifier = null;
                            ReadColumn(reader, 0, out object textFromReader, out Type textFromReaderType);
                            if (textFromReader != null && !string.IsNullOrWhiteSpace(textFromReader.ToString()))
                            {
                                documentIdentifier = textFromReader.ToString();
                            }
                            Document document = documentStorage.CreateDocument(documentIdentifier);
                            ReadTextFieldFromExcelReader(reader, document, 1, DocumentsStorage.FilesType.Text);

                            ReadTextFieldFromExcelReader(reader, document, 2, DocumentsStorage.FilesType.ScanCopy);
                            
                            ReadTextFieldFromExcelReader(reader, document, 3, DocumentsStorage.FilesType.TextPdf);
                            
                            ReadTextFieldFromExcelReader(reader, document, 4, DocumentsStorage.FilesType.Attachments);

                            bool shouldToBeAddedInDocumentStorage = false;

                            for (int attributeNumber = 0; attributeNumber < fieldCount; attributeNumber++)
                            {
                                object value;
                                Type valueType;
                                string attributeId = documentStorage.GetAttributeIdentifier(attributeNumber);
                                int colPosition = attributeNumber + attributePositionStart;
                                ReadColumn(reader, colPosition, out value, out valueType);
                                if (value != null)
                                {
                                    document.SetAttributeValue(attributeId, new DocumentAttributeValue(value, valueType));
                                    shouldToBeAddedInDocumentStorage = true;
                                    //DocumentAttributeValue documentValue = new DocumentAttributeValue(value, valueType);
                                    //if (string.IsNullOrEmpty(documentId))
                                    //{
                                    //    documentId = documentStorage.AddDocument(string.Format("{0:D5}", documentCount++));
                                    //}
                                    //documentStorage.SetAttributeValue(document.Identifier, attributeId, documentValue);
                                }
                            }

                            if (shouldToBeAddedInDocumentStorage)
                            {
                                documentStorage.AddDocument(document);
                            }
                        }
                    } while (reader.NextResult());
                }
            }
        }
    }
}
