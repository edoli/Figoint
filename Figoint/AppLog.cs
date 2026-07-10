using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Figoint
{
    public enum AppLogLevel
    {
        Debug,
        Info,
        Warn,
        Error
    }

    public sealed class AppLogEntry
    {
        public DateTimeOffset Timestamp { get; set; }
        public AppLogLevel Level { get; set; }
        public string Operation { get; set; }
        public string Message { get; set; }
        public Exception Exception { get; set; }
        public IReadOnlyDictionary<string, string> Properties { get; set; }
    }

    public interface IAppLogSink
    {
        void Write(AppLogEntry entry);
    }

    public static class AppLog
    {
        private const int LogRetentionDays = 14;
        private static readonly object SyncRoot = new object();
        private static readonly List<IAppLogSink> Sinks = new List<IAppLogSink>();

        static AppLog()
        {
            Sinks.Add(new FileAppLogSink(GetLogDirectory));
#if DEBUG
            Sinks.Add(new DebugOutputAppLogSink());
#endif
        }

        public static string LogDirectory
        {
            get { return GetLogDirectory(); }
        }

        public static void RegisterSink(IAppLogSink sink)
        {
            if (sink == null)
            {
                throw new ArgumentNullException(nameof(sink));
            }

            lock (SyncRoot)
            {
                Sinks.Add(sink);
            }
        }

        public static void Debug(string message, string operation = null, IReadOnlyDictionary<string, string> properties = null)
        {
            Write(AppLogLevel.Debug, message, null, operation, properties);
        }

        public static void Info(string message, string operation = null, IReadOnlyDictionary<string, string> properties = null)
        {
            Write(AppLogLevel.Info, message, null, operation, properties);
        }

        public static void Warn(string message, Exception exception = null, string operation = null, IReadOnlyDictionary<string, string> properties = null)
        {
            Write(AppLogLevel.Warn, message, exception, operation, properties);
        }

        public static void Error(string message, Exception exception, string operation = null, IReadOnlyDictionary<string, string> properties = null)
        {
            Write(AppLogLevel.Error, message, exception, operation, properties);
        }

        public static void CleanupOldLogs()
        {
            SafeLog(() =>
            {
                var directory = GetLogDirectory();
                if (!Directory.Exists(directory))
                {
                    return;
                }

                var cutoff = DateTime.Now.AddDays(-LogRetentionDays);
                foreach (var file in Directory.GetFiles(directory, "figoint-*.log"))
                {
                    if (File.GetLastWriteTime(file) < cutoff)
                    {
                        File.Delete(file);
                    }
                }
            });
        }

        private static void Write(AppLogLevel level, string message, Exception exception, string operation, IReadOnlyDictionary<string, string> properties)
        {
            var entry = new AppLogEntry
            {
                Timestamp = DateTimeOffset.Now,
                Level = level,
                Operation = operation,
                Message = message,
                Exception = exception,
                Properties = properties ?? new Dictionary<string, string>()
            };

            List<IAppLogSink> sinks;
            lock (SyncRoot)
            {
                sinks = Sinks.ToList();
            }

            foreach (var sink in sinks)
            {
                SafeLog(() => sink.Write(entry));
            }
        }

        private static void SafeLog(Action action)
        {
            try
            {
                action();
            }
            catch
            {
                // Logging must never interrupt the user's PowerPoint workflow.
            }
        }

        private static string GetLogDirectory()
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "Figoint", "Logs");
        }
    }

    internal sealed class FileAppLogSink : IAppLogSink
    {
        private readonly Func<string> logDirectoryProvider;
        private readonly object syncRoot = new object();

        public FileAppLogSink(Func<string> logDirectoryProvider)
        {
            this.logDirectoryProvider = logDirectoryProvider ?? throw new ArgumentNullException(nameof(logDirectoryProvider));
        }

        public void Write(AppLogEntry entry)
        {
            var directory = logDirectoryProvider();
            Directory.CreateDirectory(directory);

            var fileName = "figoint-" + DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log";
            var path = Path.Combine(directory, fileName);
            var line = Format(entry);

            lock (syncRoot)
            {
                File.AppendAllText(path, line, Encoding.UTF8);
            }
        }

        private static string Format(AppLogEntry entry)
        {
            var builder = new StringBuilder();
            builder.Append(entry.Timestamp.ToString("o", CultureInfo.InvariantCulture));
            builder.Append(" [");
            builder.Append(entry.Level);
            builder.Append("]");

            if (!string.IsNullOrWhiteSpace(entry.Operation))
            {
                builder.Append(" ");
                builder.Append(entry.Operation);
            }

            builder.Append(" - ");
            builder.AppendLine(entry.Message ?? "");

            if (entry.Properties != null)
            {
                foreach (var property in entry.Properties.OrderBy(p => p.Key))
                {
                    builder.Append("  ");
                    builder.Append(property.Key);
                    builder.Append("=");
                    builder.AppendLine(property.Value);
                }
            }

            if (entry.Exception != null)
            {
                builder.AppendLine(entry.Exception.ToString());
            }

            return builder.ToString();
        }
    }

    internal sealed class DebugOutputAppLogSink : IAppLogSink
    {
        public void Write(AppLogEntry entry)
        {
            System.Diagnostics.Debug.Write(Format(entry));
        }

        private static string Format(AppLogEntry entry)
        {
            var builder = new StringBuilder();
            builder.Append("[Figoint] ");
            builder.Append(entry.Timestamp.ToString("o", CultureInfo.InvariantCulture));
            builder.Append(" [");
            builder.Append(entry.Level);
            builder.Append("]");

            if (!string.IsNullOrWhiteSpace(entry.Operation))
            {
                builder.Append(" ");
                builder.Append(entry.Operation);
            }

            builder.Append(" - ");
            builder.AppendLine(entry.Message ?? "");

            if (entry.Exception != null)
            {
                builder.AppendLine(entry.Exception.ToString());
            }

            return builder.ToString();
        }
    }

    public static class AppCommand
    {
        public static void Run(string operation, Action action, bool notifyUserOnError = true)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            try
            {
                action();
            }
            catch (Exception ex)
            {
                AppLog.Error("User action failed.", ex, operation, GetOfficeContext());

                if (notifyUserOnError)
                {
                    MessageBox.Show(
                        "The operation could not be completed. If this keeps happening, please check the log file.\n\nLog directory: " + AppLog.LogDirectory,
                        "Figoint",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }

        private static IReadOnlyDictionary<string, string> GetOfficeContext()
        {
            var properties = new Dictionary<string, string>();

            try
            {
                var application = Globals.ThisAddIn?.Application;
                if (application == null)
                {
                    return properties;
                }

                properties["selectionType"] = application.ActiveWindow?.Selection?.Type.ToString() ?? "";

                var shapes = Util.ListSelectedShapes();
                properties["selectedShapeCount"] = shapes.Count.ToString(CultureInfo.InvariantCulture);

                var slide = application.ActiveWindow?.View?.Slide as Slide;
                if (slide != null)
                {
                    properties["slideIndex"] = slide.SlideIndex.ToString(CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                properties["contextCaptureFailed"] = "true";
            }

            return properties;
        }
    }
}
