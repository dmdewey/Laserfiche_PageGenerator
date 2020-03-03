using System;
using System.IO;

namespace Laserfiche_PageGenerator
{
    public sealed class Logger
    {
        public static void MessageToFile(string message)
        {
            var monthFolder = AppDomain.CurrentDomain.BaseDirectory + "Log_" + DateTime.Now.Month + "_" + DateTime.Now.Year;
            if (!Directory.Exists(monthFolder))
            {
                Directory.CreateDirectory(monthFolder);
            }
            var fileName = monthFolder + @"\Log_" + DateTime.Now.Month + "_" + DateTime.Now.Day + "_" + DateTime.Now.Year + ".txt";
            if (!File.Exists(fileName))
            {
                File.Create(fileName).Close();
            }
            using (var streamWriter = File.AppendText(fileName))
            {
                try
                {
                    streamWriter.WriteLine($"{DateTime.Now:G}: {message}.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Message: {ex.Message} - InnerException: {ex.InnerException}");
                }
            }
        }
    }
}
