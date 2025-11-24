using System;

namespace KoreanIMEFixer.Input
{
    public static class InputState
    {
        private static readonly object _lock = new object();
        private static int _lastVk = 0;
        private static DateTime _lastTime = DateTime.MinValue;
        private static readonly System.Collections.Generic.List<(int vk, DateTime time)> _recent = new System.Collections.Generic.List<(int, DateTime)>();

        public static void ReportPhysicalKey(int vk)
        {
            lock (_lock)
            {
                _lastVk = vk;
                _lastTime = DateTime.UtcNow;
                try
                {
                    _recent.Add((vk, _lastTime));
                    DateTime cutoff = DateTime.UtcNow - TimeSpan.FromMilliseconds(2000);
                    int removeCount = 0;
                    for (int i = 0; i < _recent.Count; i++) { if (_recent[i].time < cutoff) removeCount++; else break; }
                    if (removeCount > 0) _recent.RemoveRange(0, removeCount);
                    const int MaxRecent = 64; if (_recent.Count > MaxRecent) _recent.RemoveRange(0, _recent.Count - MaxRecent);
                }
                catch { }
            }
        }

        public static (int vk, DateTime time) GetLast() { lock (_lock) { return (_lastVk, _lastTime); } }

        public static int CountRecentAscii(int withinMs)
        {
            lock (_lock)
            {
                if (_recent.Count == 0) return 0;
                DateTime now = DateTime.UtcNow;
                int cnt = 0;
                for (int i = _recent.Count - 1; i >= 0; i--)
                {
                    var item = _recent[i];
                    if ((now - item.time).TotalMilliseconds <= withinMs) { if (item.vk >= 65 && item.vk <= 90) cnt++; } else break;
                }
                return cnt;
            }
        }
    }
}
