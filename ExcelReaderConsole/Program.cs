using System;
using System.Text;

namespace ExcelReaderConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.Unicode;
            try
            {
                Application app = new Application();
                app.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occured:");
                ConsolePrintException(ex);
            }
            Console.WriteLine("End. Application will be close...");
            Console.ReadLine();
        }

        private static void ConsolePrintException(Exception ex)
        {
            if (ex != null)
            {
                Console.WriteLine(ex.Message);
                ConsolePrintException(ex.InnerException);
            }
        }
    }
}
