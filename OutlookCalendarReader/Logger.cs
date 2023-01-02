using System;
using System.IO;
using System.Reflection;

namespace OutlookCalendarReader
{
    internal class Logger
    {
        private static readonly object LockObject = new();

        public static void Log(string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var formattedMessage = $"[{timestamp}] {message}";

            lock (LockObject)
            {
                Console.WriteLine(formattedMessage);
                var runDir = Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location);
                var logFile = Path.Join(runDir, "log.log");
                File.AppendAllText(logFile, formattedMessage + "\n");
            }
        }
    }
}
