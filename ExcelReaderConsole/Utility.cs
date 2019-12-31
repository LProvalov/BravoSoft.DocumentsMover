using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReaderConsole
{
    public class Utility
    {
        public static string GenerateStringIdentifier(int seed)
        {
            return Guid.NewGuid().ToString();
        }

        public static int GetSessionSeed()
        {
            return 1;
        }

        public static DirectoryInfo CheckDirectoryAndCreate(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
            {
                dir.Create();
            }
            return dir;
        }
    }
}
