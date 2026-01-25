using System;
using System.Diagnostics;
using System.Text;

namespace EasyOpenXml.Excel
{
    public sealed class Logger
    {
        private readonly string _logPath;
        private readonly object _lock = new object();
        private readonly Stopwatch _sw = new Stopwatch();
        private bool _started;

        // public

        public Logger(string logPath)
        {
            if (string.IsNullOrWhiteSpace(logPath))
            {
                throw new ArgumentException("logPath is required.", nameof(logPath));
            }

            _logPath = Path.GetFullPath(logPath);

            var dir = Path.GetDirectoryName(_logPath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        public void Start(string label = null)
        {
            if (_started)
            {
                throw new InvalidOperationException("Logger has already been started.");
            }

            _started = true;
            _sw.Restart();

            WriteLine($"{Prefix()} 開始 : {FormatLabel(label)}");
        }

        public void Log(string message)
        {
            if (!_started)
            {
                throw new InvalidOperationException("Logger has not been started.");
            }

            WriteLine($"{Prefix()} 経過 {FormatElapsed(_sw.Elapsed)} : {message}");
        }

        public void End(string label = null)
        {
            if (!_started)
            {
                throw new InvalidOperationException("Logger has not been started.");
            }

            _sw.Stop();

            WriteLine($"{Prefix()} 終了 : {FormatLabel(label)}");
            WriteLine($"{Prefix()} 合計 : {FormatElapsed(_sw.Elapsed)}");

            _started = false;
        }

        // private

        private static string Prefix()
        {
            return $"{DateTime.Now.ToString("[yyyy/MM/dd HH:mm:ss]")}";
        }

        private static string FormatLabel(string label)
        {
            return string.IsNullOrEmpty(label) ? "処理" : label;
        }

        private static string FormatElapsed(TimeSpan elapsed)
        {
            return $"{elapsed.Minutes:00}分{elapsed.Seconds:00}秒";
        }

        private void WriteLine(string line)
        {
            try
            {
                lock (_lock)
                {
                    using (var fs = new FileStream(_logPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                    using (var sw = new StreamWriter(fs, Encoding.UTF8))
                    {
                        sw.WriteLine(line);
                        sw.Flush();
                    }
                }
            }
            catch
            {
                // 調査用のため「ログ失敗で本処理を落とさない」方針
            }
        }

    }
}

