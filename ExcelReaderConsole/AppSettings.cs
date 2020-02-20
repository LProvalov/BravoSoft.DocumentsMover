using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace ExcelReaderConsole
{
    public class AppSettings
    {
        private static readonly string APP_SETTINGS_PATH = "AppConfig.xml";
        private AppSettingsModel model;
        private XmlSerializer serializer;
        private bool isLoaded;

        private AppSettings()
        {
            model = new AppSettingsModel();
            serializer = new XmlSerializer(typeof(AppSettingsModel));
            isLoaded = false;
        }

        private bool DeserializeAppSettingsModel(string settingsPath)
        {
            FileInfo settingsFile = new FileInfo(settingsPath);
            if (!settingsFile.Exists) return false;
            StreamReader sr = null;
            try
            {
                using (sr = new StreamReader(settingsPath))
                {
                    model = serializer.Deserialize(sr) as AppSettingsModel;
                }
            }
            finally
            {
                if (sr != null) sr.Close();
            }
            return true;
        }
        private void SerializeAppSettingsModel(string settingsPath)
        {
            StreamWriter streamWriter = null;
            try
            {
                streamWriter = new StreamWriter(settingsPath, false);
                serializer.Serialize(streamWriter, model);
            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (streamWriter != null) streamWriter.Close();
            }
        }

        public void LoadAppSettings()
        {
            DeserializeAppSettingsModel(APP_SETTINGS_PATH);
            isLoaded = true;
        }
        public void SaveAppSettings()
        {
            SerializeAppSettingsModel(APP_SETTINGS_PATH);
        }

        private static AppSettings instance = null;
        public static AppSettings Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new AppSettings();
                }
                return instance;
            }
        }

        private void CheckAppSettingsLoaded()
        {
            if (!isLoaded) throw new Exception("AppSettings does not loaded.");
        }

        public string GetCardFileExtension()
        {
            CheckAppSettingsLoaded();
            return model.CardFileExtension;
        }

        public string ExcelTemplateFilePath
        {
            get
            {
                CheckAppSettingsLoaded();
                return model.ExcelTemplateFilePath;
            }
            set
            {
                model.ExcelTemplateFilePath = value;
            }
        }

        private string GetDirectoryPath(string path, string defaultFolder)
        {
            if (!string.IsNullOrEmpty(path))
            {
                return path;
            }
            else
            {
                string defaultPath = Path.Combine(Directory.GetCurrentDirectory(), defaultFolder);
                var dir = Utility.CheckDirectoryAndCreate(defaultPath);
                return dir.FullName;
            }
        }

        public string InputDirectoryPath
        {
            get
            {
                CheckAppSettingsLoaded();
                return model.InputDirectoryPath;
            }
            set
            {
                model.InputDirectoryPath = value;
            }
        }

        public string OutputDirectoryPath
        {
            get
            {
                CheckAppSettingsLoaded();
                return model.OutputDirectoryPath;
            }
            set
            {
                model.OutputDirectoryPath = value;
            }
        }

        public string GetEncoding()
        {
            CheckAppSettingsLoaded();
            return model.Encoding;
        }

        public bool Validate(out string errorMsg)
        {
            bool valid = true;
            errorMsg = string.Empty;
            StringBuilder sb = new StringBuilder();
            if (string.IsNullOrEmpty(this.InputDirectoryPath?.Trim()))
            {
                valid = false;
                sb.AppendLine("Input directory is null or empty.");
            }

            if (string.IsNullOrEmpty(this.OutputDirectoryPath?.Trim()))
            {
                valid = false;
                sb.AppendLine("Output directory is null or empty.");
            }

            if (string.IsNullOrEmpty(this.ExcelTemplateFilePath?.Trim()))
            {
                valid = false;
                sb.AppendLine("Template file path is null or empty.");
            }

            if (!valid) errorMsg = sb.ToString();
            return valid;
        }
    }
}
