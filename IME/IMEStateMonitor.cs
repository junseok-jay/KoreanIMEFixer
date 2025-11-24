using System;
using System.Runtime.InteropServices;

namespace KoreanIMEFixer.IME
{
    public class IMEStateMonitor
    {
        public bool IsKoreanMode()
        {
            IntPtr hWnd = GetForegroundWindow();
            IntPtr hIMC = ImmGetContext(hWnd);
            if (hIMC == IntPtr.Zero)
            {
                IntPtr layout = GetKeyboardLayout(0);
                return (layout.ToInt64() & 0xFFFF) == 0x412;
            }
            int conv = 0, sent = 0;
            ImmGetConversionStatus(hIMC, out conv, out sent);
            ImmReleaseContext(hWnd, hIMC);
            return (conv & 0x01) != 0;
        }

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetContext(IntPtr hWnd);
        [DllImport("imm32.dll")] private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
        [DllImport("imm32.dll")] private static extern int ImmGetConversionStatus(IntPtr hIMC, out int lpfdwConversion, out int lpfdwSentence);
        [DllImport("user32.dll")] private static extern IntPtr GetKeyboardLayout(uint idThread);
    }
}
