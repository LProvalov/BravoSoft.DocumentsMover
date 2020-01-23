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
                ConsoleApplication app = new ConsoleApplication();
                app.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occured:");
                ConsolePrintException(ex);
            }
            Console.WriteLine("End. ConsoleApplication will be close...");
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
