using System;
using System.IO;
using System.Text;

namespace KoreanIMEFixer.Logging
{
    public static class LogService
    {
        private static readonly string LogDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KoreanIMEFixer");
        private static readonly string LogPath = Path.Combine(LogDir, "app.log");

        static LogService()
        {
            try { if (!Directory.Exists(LogDir)) Directory.CreateDirectory(LogDir); } catch { }
        }

        public static void Write(string message)
        {
            try
            {
                var line = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") + " : " + message;
                File.AppendAllText(LogPath, line + Environment.NewLine, Encoding.UTF8);
                Console.WriteLine(line);
            }
            catch { }
        }

        public static string ReadAll()
        {
            try { return File.Exists(LogPath) ? File.ReadAllText(LogPath, Encoding.UTF8) : string.Empty; }
            catch { return string.Empty; }
        }
    }
}
