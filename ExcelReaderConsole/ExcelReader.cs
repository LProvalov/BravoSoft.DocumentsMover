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
                        value = reader.GetDouble(colPosition).ToString();
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
                        int attributePositionStart = 2;
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
                            string documentId = null; 

                            object textFileName;
                            Type textFileNameType;
                            ReadColumn(reader, 1, out textFileName, out textFileNameType);
                            if (textFileNameType == typeof(string) && !string.IsNullOrEmpty(textFileName as string))
                            {
                                if (string.IsNullOrEmpty(documentId))
                                {
                                    documentId = documentStorage.AddDocument(string.Format("{0:D5}", documentCount++));
                                }
                                documentStorage.SetTextFileName(documentId, textFileName as string);
                            }

                            object scanFileName;
                            Type scanFileNameType;
                            ReadColumn(reader, 2, out scanFileName, out scanFileNameType);
                            if (scanFileNameType == typeof(string) && !string.IsNullOrEmpty(scanFileName as string))
                            {
                                if (string.IsNullOrEmpty(documentId))
                                {
                                    documentId = documentStorage.AddDocument(string.Format("{0:D5}", documentCount++));
                                }
                                documentStorage.SetScanFileName(documentId, scanFileName as string);
                            }

                            for (int attributeNumber = 0; attributeNumber < fieldCount; attributeNumber++)
                            {
                                object value;
                                Type valueType;
                                string attributeId = documentStorage.GetAttributeIdentifier(attributeNumber);
                                int colPosition = attributeNumber + attributePositionStart;
                                ReadColumn(reader, colPosition, out value, out valueType);
                                if (value != null)
                                {
                                    DocumentAttributeValue documentValue = new DocumentAttributeValue(value, valueType);
                                    if (string.IsNullOrEmpty(documentId))
                                    {
                                        documentId = documentStorage.AddDocument(string.Format("{0:D5}", documentCount++));
                                    }
                                    documentStorage.SetAttributeValue(documentId, attributeId, documentValue);
                                }
                            }
                        }
                    } while (reader.NextResult());
                }
            }
        }
    }
}
