using System;
using System.IO;

namespace Gryzak.Services
{
    /// <summary>
    /// Poziomy ważności logów
    /// </summary>
    public enum LogSeverity
    {
        Debug = 0,      // Najniższy poziom - szczegółowe informacje diagnostyczne
        Info = 1,       // Informacje ogólne
        Warning = 2,   // Ostrzeżenia - coś może być nieprawidłowe, ale aplikacja działa
        Error = 3,      // Błędy - coś poszło nie tak, ale aplikacja może kontynuować
        Critical = 4    // Błędy krytyczne - poważne problemy wymagające uwagi
    }

    /// <summary>
    /// Klasa do logowania z poziomami ważności
    /// </summary>
    public static class Logger
    {
        private static LogSeverity _minimumSeverity = LogSeverity.Debug;

        /// <summary>
        /// Ustawia minimalny poziom ważności logów (logi o niższym poziomie będą ignorowane)
        /// </summary>
        public static void SetMinimumSeverity(LogSeverity severity)
        {
            _minimumSeverity = severity;
        }

        /// <summary>
        /// Loguje wiadomość z określonym poziomem ważności
        /// </summary>
        public static void Log(LogSeverity severity, string message, string? source = null)
        {
            if (severity < _minimumSeverity)
            {
                return; // Ignoruj logi o zbyt niskim poziomie
            }

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string severityString = GetSeverityString(severity);
            string sourceString = !string.IsNullOrWhiteSpace(source) ? $"[{source}]" : "";
            
            string formattedMessage = $"[{timestamp}] [{severityString}] {sourceString} {message}";
            
            Console.WriteLine(formattedMessage);
        }

        /// <summary>
        /// Loguje wiadomość z poziomem Debug
        /// </summary>
        public static void Debug(string message, string? source = null)
        {
            Log(LogSeverity.Debug, message, source);
        }

        /// <summary>
        /// Loguje wiadomość z poziomem Info
        /// </summary>
        public static void Info(string message, string? source = null)
        {
            Log(LogSeverity.Info, message, source);
        }

        /// <summary>
        /// Loguje wiadomość z poziomem Warning
        /// </summary>
        public static void Warning(string message, string? source = null)
        {
            Log(LogSeverity.Warning, message, source);
        }

        /// <summary>
        /// Loguje wiadomość z poziomem Error
        /// </summary>
        public static void Error(string message, string? source = null)
        {
            Log(LogSeverity.Error, message, source);
        }

        /// <summary>
        /// Loguje wiadomość z poziomem Critical
        /// </summary>
        public static void Critical(string message, string? source = null)
        {
            Log(LogSeverity.Critical, message, source);
        }

        /// <summary>
        /// Loguje wyjątek z poziomem Error
        /// </summary>
        public static void Error(Exception exception, string? source = null, string? additionalMessage = null)
        {
            string message = additionalMessage != null 
                ? $"{additionalMessage}: {exception.Message}" 
                : exception.Message;
            
            string fullMessage = $"{message}\n{exception.StackTrace}";
            Log(LogSeverity.Error, fullMessage, source);
        }

        /// <summary>
        /// Loguje wyjątek z poziomem Critical
        /// </summary>
        public static void Critical(Exception exception, string? source = null, string? additionalMessage = null)
        {
            string message = additionalMessage != null 
                ? $"{additionalMessage}: {exception.Message}" 
                : exception.Message;
            
            string fullMessage = $"{message}\n{exception.StackTrace}";
            Log(LogSeverity.Critical, fullMessage, source);
        }

        /// <summary>
        /// Konwertuje poziom ważności na string
        /// </summary>
        private static string GetSeverityString(LogSeverity severity)
        {
            return severity switch
            {
                LogSeverity.Debug => "DEBUG",
                LogSeverity.Info => "INFO ",
                LogSeverity.Warning => "WARN ",
                LogSeverity.Error => "ERROR",
                LogSeverity.Critical => "CRIT ",
                _ => "UNKWN"
            };
        }
    }
}
