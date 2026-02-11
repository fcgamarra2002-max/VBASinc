using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace VBASinc.Diagnostics
{
    public static class AddInLogger
    {
        private static readonly object SyncRoot = new object();
        private static readonly string LogDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBASinc");
        private static readonly string LogFilePath = Path.Combine(LogDirectory, "VBASinc.log");

        public static void Log(string message)
        {
            Write(message, null);
        }

        public static void Log(string message, Exception ex)
        {
            Write(message, ex);
        }

        private static void Write(string message, Exception? ex)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            if (ex != null)
            {
                line = $"{line} | {ex.GetType().Name}: {ex.Message}";
            }

            try
            {
                lock (SyncRoot)
                {
                    Directory.CreateDirectory(LogDirectory);
                    File.AppendAllText(LogFilePath, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
            }

            Debug.WriteLine(line);
            if (ex != null)
            {
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}
