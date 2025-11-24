using System;
using System.Runtime.InteropServices;

namespace KoreanIMEFixer
{
    internal static class NativeMethods
    {
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] internal static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct GUITHREADINFO
    {
        public int cbSize;
        public uint flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
