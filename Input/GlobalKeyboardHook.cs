using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using KoreanIMEFixer.Logging;
using System;
using System.Runtime.CompilerServices;
using KoreanIMEFixer.Input;

namespace KoreanIMEFixer.Input
{
    public class GlobalKeyboardHook
    {
        public event EventHandler<KeyPressedEventArgs>? KeyPressed;

        private IntPtr _hookId = IntPtr.Zero;
        private LowLevelKeyboardProc _proc;

        public GlobalKeyboardHook()
        {
            _proc = HookCallback;
            _hookId = SetHook(_proc);
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            Process curProcess = Process.GetCurrentProcess();
            ProcessModule? curModule = curProcess.MainModule;
            IntPtr moduleHandle = curModule != null ? GetModuleHandle(curModule.ModuleName) : IntPtr.Zero;
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc, moduleHandle, 0);
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            const int WM_KEYDOWN = 0x0100;
            const int WM_SYSKEYDOWN = 0x0104;

            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    try
                    {
                        // If suppressor asks to block this ASCII key for the focused Notion window, swallow it and record it for reinjection.
                        if (KeySuppressor.ShouldSuppressKey(vkCode))
                        {
                            try { KeySuppressor.RecordSuppressedKey(vkCode); } catch { }
                            LogService.Write($"GlobalKeyboardHook: suppressed ASCII VK={vkCode} due to Notion suppression window.");
                            return (IntPtr)1;
                        }
                    }
                    catch { }
                    int flags = Marshal.ReadInt32(lParam, 8);
                    const int LLKHF_INJECTED = 0x10;
                    if ((flags & LLKHF_INJECTED) != 0)
                    {
                        return CallNextHookEx(_hookId, nCode, wParam, lParam);
                    }

                    var args = new KeyPressedEventArgs((Keys)vkCode);
                    KeyPressed?.Invoke(this, args);
                    if (args.Suppress) return (IntPtr)1;
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        ~GlobalKeyboardHook() { UnhookWindowsHookEx(_hookId); }

        #region WinAPI
        private const int WH_KEYBOARD_LL = 13;
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(string lpModuleName);
        #endregion
    }
}
