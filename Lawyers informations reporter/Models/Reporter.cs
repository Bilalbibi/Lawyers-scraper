using System;

namespace Lawyers_informations_reporter.Models
{
    public static class Reporter
    {
        public static EventHandler<string> OnLog;
        public static EventHandler<string> OnError;
        public static EventHandler<(int nbr, int total, string message)> OnProgress;
        public static void Log(string s)
        {
            OnLog?.Invoke(null, s);
        }
        public static void Error(string s)
        {
            OnError?.Invoke(null, s);
        }
        public static void Progress(int nbr, int total, string message)
        {
            OnProgress?.Invoke(null, (nbr, total, message));
        }
    }
}
