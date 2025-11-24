using System;
using System.Runtime.InteropServices;
using KoreanIMEFixer.Logging;

namespace KoreanIMEFixer.Input
{
    public static class KeySuppressor
    {
        private static readonly object _lock = new object();
        private static IntPtr _hwnd = IntPtr.Zero;
        private static DateTime _expiry = DateTime.MinValue;
        private static readonly System.Collections.Generic.List<int> _bufferedKeys = new System.Collections.Generic.List<int>();

        public static void StartSuppressionForWindow(IntPtr hwnd, int durationMs)
        {
            if (hwnd == IntPtr.Zero) return;
            lock (_lock)
            {
                _hwnd = hwnd;
                _expiry = DateTime.UtcNow.AddMilliseconds(durationMs);
                _bufferedKeys.Clear();
            }
            System.Threading.Tasks.Task.Run(() =>
            {
                try { System.Threading.Thread.Sleep(durationMs + 30); ReinjectBufferedKeysIfNeeded(hwnd); } catch { }
            });
        }

        public static void RecordSuppressedKey(int vk)
        {
            try { lock (_lock) { if (_hwnd == IntPtr.Zero) return; _bufferedKeys.Add(vk); if (_bufferedKeys.Count > 32) _bufferedKeys.RemoveRange(0, _bufferedKeys.Count - 32); } } catch { }
        }

        private static void ReinjectBufferedKeysIfNeeded(IntPtr hwnd)
        {
            try
            {
                lock (_lock)
                {
                    if (_hwnd != hwnd) { LogService.Write($"KeySuppressor: suppression moved to another window (expected=0x{hwnd.ToInt64():X}, current=0x{_hwnd.ToInt64():X}). Skipping reinjection."); return; }
                    if (_bufferedKeys.Count == 0) { _hwnd = IntPtr.Zero; return; }

                    IntPtr fg = GetForegroundWindow();
                    if (fg != hwnd) { LogService.Write($"KeySuppressor: foreground changed before reinjection (target=0x{hwnd.ToInt64():X}, fg=0x{fg.ToInt64():X}). Dropping {_bufferedKeys.Count} buffered key(s)."); _bufferedKeys.Clear(); _hwnd = IntPtr.Zero; return; }

                    var inputs = new System.Collections.Generic.List<INPUT>();
                    foreach (var vk in _bufferedKeys)
                    {
                        uint scan = MapVirtualKey((uint)vk, 0);
                        inputs.Add(new INPUT { type = 1, ki = new KEYBDINPUT { wVk = 0, wScan = (ushort)scan, dwFlags = 0x0008 } });
                        inputs.Add(new INPUT { type = 1, ki = new KEYBDINPUT { wVk = 0, wScan = (ushort)scan, dwFlags = 0x0008 | 0x0002 } });
                    }
                    if (inputs.Count > 0)
                    {
                        try { uint sent = SendInput((uint)inputs.Count, inputs.ToArray(), INPUT.Size); LogService.Write($"KeySuppressor: SendInput sent={sent} of {inputs.Count} events for reinjection."); } catch (Exception ex) { LogService.Write($"KeySuppressor: SendInput exception during reinjection: {ex.Message}"); }
                    }
                    _bufferedKeys.Clear(); _hwnd = IntPtr.Zero;
                }
            }
            catch { }
        }

        public static bool ShouldSuppressKey(int vk)
        {
            try { lock (_lock) { if (_hwnd == IntPtr.Zero) return false; if (DateTime.UtcNow > _expiry) { _hwnd = IntPtr.Zero; return false; } if (!(vk >= 65 && vk <= 90)) return false; IntPtr fg = GetForegroundWindow(); return fg == _hwnd; } } catch { return false; }
        }

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)] private static extern uint MapVirtualKey(uint uCode, uint uMapType);
        [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint nInputs, [In] INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT { public uint type; public KEYBDINPUT ki; public static int Size => Marshal.SizeOf(typeof(INPUT)); }
        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public UIntPtr dwExtraInfo; }
    }
}
