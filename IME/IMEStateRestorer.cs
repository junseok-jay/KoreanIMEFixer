using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text;
using System.Threading;
using KoreanIMEFixer.Logging;
using Keys = System.Windows.Forms.Keys;

namespace KoreanIMEFixer.IME
{
    public class IMEStateRestorer
    {
        private const byte VK_HANGUL = 0x15;
        // state for composing jamo sequences when fallback paste is used
        private char? _pendingJamo = null;
        private DateTime _pendingTime = DateTime.MinValue;
        private int _pendingPastedCount = 0;
        private readonly TimeSpan _composeWindow = TimeSpan.FromMilliseconds(1200);
        // debounce for Ctrl+V attempts per HWND to avoid duplicate Ctrl+V calls
        private static readonly object _ctrlVLock = new object();
        private static readonly System.Collections.Generic.Dictionary<long, (DateTime time, bool success)> _lastCtrlVAttempts = new System.Collections.Generic.Dictionary<long, (DateTime, bool)>();

        [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint nInputs, [In] INPUT[] pInputs, int cbSize);
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] private static extern uint MapVirtualKey(uint uCode, uint uMapType);
        // Clipboard / Global memory APIs for paste fallback
        [DllImport("user32.dll", SetLastError = true)] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool CloseClipboard();
        [DllImport("user32.dll", SetLastError = true)] private static extern bool EmptyClipboard();
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        [DllImport("user32.dll")] private static extern bool IsClipboardFormatAvailable(uint format);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GlobalUnlock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr GlobalFree(IntPtr hMem);
        private const uint GMEM_MOVEABLE = 0x0002;
        private const uint CF_UNICODETEXT = 13;

        // Helpers for composing jamo -> syllable
        private static bool IsCompatibilityJamo(char ch)
        {
            // Range cover used compatibility jamo codepoints in this project
            return (ch >= '\u3131' && ch <= '\u3163');
        }

        private static string? TryComposePair(char a, char b)
        {
            // a = leading consonant (compatibility jamo), b = vowel (compatibility jamo)
            int? L = CompatJamoToLIndex(a);
            int? V = CompatJamoToVIndex(b);
            if (L.HasValue && V.HasValue)
            {
                const int SBase = 0xAC00;
                const int VCount = 21;
                const int TCount = 28;
                int SIndex = (L.Value * VCount + V.Value) * TCount;
                char syll = (char)(SBase + SIndex);
                return new string(syll, 1);
            }
            return null;
        }

        private static int? CompatJamoToLIndex(char ch)
        {
            // map compatibility jamo to Choseong index
            return ch switch
            {
                '\u3131' => 0, // ㄱ
                '\u3132' => 0, // ㄱ? (rare)
                '\u3133' => 0,
                '\u3134' => 2, // ㄴ
                '\u3137' => 3, // ㄷ
                '\u3139' => 5, // ㄹ
                '\u3141' => 6, // ㅁ
                '\u3142' => 7, // ㅂ
                '\u3145' => 9, // ㅅ
                '\u3147' => 11, // ㅇ
                '\u3148' => 12, // ㅈ
                '\u314A' => 14, // ㅊ
                '\u314B' => 15, // ㅋ
                '\u314C' => 16, // ㅌ
                '\u314D' => 17, // ㅍ
                '\u314E' => 18, // ㅎ
                // additional mappings for common consonants
                _ => ch switch
                {
                    '\u3140' => 0,
                    '\u3143' => 2,
                    _ => (int?)null
                }
            };
        }

        private static int? CompatJamoToVIndex(char ch)
        {
            return ch switch
            {
                '\u314F' => 0, // ㅏ
                '\u3150' => 1, // ㅐ
                '\u3151' => 2, // ㅑ
                '\u3152' => 3, // ㅒ
                '\u3153' => 4, // ㅓ
                '\u3154' => 5, // ㅔ
                '\u3155' => 6, // ㅕ
                '\u3156' => 7, // ㅖ
                '\u3157' => 8, // ㅗ
                '\u3158' => 9, // ㅘ
                '\u3159' => 10, // ㅙ
                '\u315A' => 11, // ㅚ
                '\u315B' => 12, // ㅛ
                '\u315C' => 13, // ㅜ
                '\u315D' => 14, // ㅝ
                '\u315E' => 15, // ㅞ
                '\u315F' => 16, // ㅟ
                '\u3160' => 17, // ㅠ
                '\u3161' => 18, // ㅡ
                '\u3162' => 19, // ㅢ
                '\u3163' => 20, // ㅣ
                _ => (int?)null
            };
        }

        private static char? TryComposeVowelPair(char a, char b)
        {
            // compose compatibility vowel pairs to composite vowel compatibility jamo
            return (a, b) switch
            {
                ('\u3157', '\u314F') => '\u3158', // ㅗ + ㅏ -> ㅘ
                ('\u3157', '\u3150') => '\u3159', // ㅗ + ㅐ -> ㅙ
                ('\u3157', '\u3163') => '\u315A', // ㅗ + ㅣ -> ㅚ
                ('\u315C', '\u3153') => '\u315D', // ㅜ + ㅓ -> ㅝ
                ('\u315C', '\u3154') => '\u315E', // ㅜ + ㅔ -> ㅞ
                ('\u315C', '\u3163') => '\u315F', // ㅜ + ㅣ -> ㅟ
                ('\u3161', '\u3163') => '\u3162', // ㅡ + ㅣ -> ㅢ
                _ => (char?)null
            };
        }

        private void TrySendBackspaces(IntPtr hwnd, int count)
        {
            if (count <= 0) return;
            try
            {
                const uint WM_KEYDOWN = 0x0100;
                const uint WM_KEYUP = 0x0101;
                const int VK_BACK = 0x08;
                uint pid;
                uint targetThread = GetWindowThreadProcessId(hwnd, out pid);
                uint currentThread = GetCurrentThreadId();
                bool attached = false;
                if (targetThread != 0 && currentThread != 0)
                {
                    attached = AttachThreadInput(currentThread, targetThread, true);
                }
                for (int i = 0; i < count; i++)
                {
                    PostMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_BACK, IntPtr.Zero);
                    PostMessage(hwnd, WM_KEYUP, (IntPtr)VK_BACK, IntPtr.Zero);
                }
                if (attached)
                {
                    try { AttachThreadInput(currentThread, targetThread, false); } catch { }
                }
                LogService.Write($"Sent {count} backspace(es) to HWND=0x{hwnd.ToInt64():X}");
            }
            catch (Exception ex) { LogService.Write("TrySendBackspaces failed: " + ex.Message); }
        }

        // Public helper: send backspace(s) to the focused control of the foreground window.
        public void RemovePreviousChars(int count)
        {
            try
            {
                IntPtr hwndForeground = GetForegroundWindow();
                IntPtr hwndTarget = hwndForeground;
                try
                {
                    uint pid;
                    uint threadId = GetWindowThreadProcessId(hwndForeground, out pid);
                    GUITHREADINFO gti = new GUITHREADINFO();
                    gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                    if (threadId != 0 && GetGUIThreadInfo(threadId, ref gti))
                    {
                        if (gti.hwndFocus != IntPtr.Zero)
                        {
                            hwndTarget = gti.hwndFocus;
                        }
                    }
                }
                catch { }

                if (hwndTarget != IntPtr.Zero && count > 0)
                {
                    TrySendBackspaces(hwndTarget, count);
                    LogService.Write($"RemovePreviousChars: sent {count} backspace(s) to HWND=0x{hwndTarget.ToInt64():X}");
                }
            }
            catch (Exception ex)
            {
                LogService.Write("RemovePreviousChars failed: " + ex.Message);
            }
        }

        private bool TryClipboardPasteString(IntPtr hwnd, string s)
        {
            string? original = null;
            try
            {
                // Read existing clipboard text (if any)
                try
                {
                    if (OpenClipboard(IntPtr.Zero))
                    {
                        if (IsClipboardFormatAvailable(CF_UNICODETEXT))
                        {
                            IntPtr hData = GetClipboardData(CF_UNICODETEXT);
                            if (hData != IntPtr.Zero)
                            {
                                IntPtr ptr = GlobalLock(hData);
                                if (ptr != IntPtr.Zero)
                                {
                                    original = Marshal.PtrToStringUni(ptr);
                                    GlobalUnlock(hData);
                                }
                            }
                        }
                        CloseClipboard();
                    }
                }
                catch { }

                bool setOk = false;
                try
                {
                    if (OpenClipboard(IntPtr.Zero))
                    {
                        EmptyClipboard();
                        byte[] bytes = Encoding.Unicode.GetBytes(s + '\0');
                        IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytes.Length);
                        if (hGlobal != IntPtr.Zero)
                        {
                            IntPtr target = GlobalLock(hGlobal);
                            if (target != IntPtr.Zero)
                            {
                                Marshal.Copy(bytes, 0, target, bytes.Length);
                                GlobalUnlock(hGlobal);
                                if (SetClipboardData(CF_UNICODETEXT, hGlobal) != IntPtr.Zero) setOk = true;
                            }
                            else GlobalFree(hGlobal);
                        }
                        CloseClipboard();
                    }
                }
                catch { }

                if (!setOk) return false;

                const uint WM_PASTE = 0x0302;
                // Post a WM_PASTE to the target window; many Electron apps don't process SendMessage WM_PASTE reliably.
                bool posted = false;
                try
                {
                    posted = PostMessage(hwnd, WM_PASTE, IntPtr.Zero, IntPtr.Zero);
                    LogService.Write($"PostMessage WM_PASTE (composed) posted={posted} to HWND=0x{hwnd.ToInt64():X}");
                }
                catch (Exception ex)
                {
                    LogService.Write("PostMessage WM_PASTE failed: " + ex.Message);
                }

                // If PostMessage posted, schedule a conservative delayed Ctrl+V fallback
                // (many Electron apps may not process PostMessage immediately). This
                // schedules a best-effort secondary attempt after a short sleep so we
                // don't block the calling thread or risk deleting/restoring clipboard prematurely.
                if (posted)
                {
                    try
                    {
                        // Start a short suppression window for ASCII keys in the target Notion window
                        try { Input.KeySuppressor.StartSuppressionForWindow(hwnd, 250); } catch { }

                        ThreadPool.QueueUserWorkItem(_ =>
                        {
                            try
                            {
                                Thread.Sleep(200);
                                bool vk = TrySendCtrlVToWindow(hwnd);
                                LogService.Write($"Delayed Ctrl+V fallback after PostMessage result={vk} to HWND=0x{hwnd.ToInt64():X}");
                                // Heuristic cleanup for Notion: if recent physical ASCII keys were pressed
                                // very close to the paste time, they may have leaked through as Latin
                                // characters before our pasted Hangul. Send as many Backspace
                                // events as recent ASCII keys within the window (limited to Notion).
                                try
                                {
                                    uint pid;
                                    GetWindowThreadProcessId(hwnd, out pid);
                                    if (pid != 0)
                                    {
                                        string proc = string.Empty;
                                        try { proc = System.Diagnostics.Process.GetProcessById((int)pid).ProcessName; } catch { }
                                        if (!string.IsNullOrEmpty(proc) && proc.IndexOf("notion", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            int asciiCount = Input.InputState.CountRecentAscii(700);
                                            const int MaxBackspaces = 6;
                                            int toRemove = Math.Min(asciiCount, MaxBackspaces);
                                            if (toRemove > 0)
                                            {
                                                TrySendBackspaces(hwnd, toRemove);
                                                LogService.Write($"Notion heuristic: sent {toRemove} backspace(s) (counted={asciiCount}) to HWND=0x{hwnd.ToInt64():X}");
                                                // short wait and attempt to remove any additional quick keystrokes
                                                try { Thread.Sleep(40); } catch { }
                                                int asciiRecentAfter = Input.InputState.CountRecentAscii(300);
                                                int extra = Math.Min(MaxBackspaces - toRemove, asciiRecentAfter);
                                                if (extra > 0)
                                                {
                                                    TrySendBackspaces(hwnd, extra);
                                                    LogService.Write($"Notion heuristic: sent additional {extra} backspace(s) after second pass to HWND=0x{hwnd.ToInt64():X}");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception exInner2)
                                {
                                    LogService.Write("Heuristic cleanup exception: " + exInner2.Message);
                                }
                            }
                            catch (Exception exInner)
                            {
                                LogService.Write("Delayed Ctrl+V fallback exception: " + exInner.Message);
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        LogService.Write("Scheduling delayed Ctrl+V fallback failed: " + ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        LogService.Write("PostMessage WM_PASTE not posted; attempting Ctrl+V fallback.");
                        bool vok = TrySendCtrlVToWindow(hwnd);
                        LogService.Write($"Ctrl+V fallback result={vok}");
                    }
                    catch (Exception ex)
                    {
                        LogService.Write("Ctrl+V fallback exception: " + ex.Message);
                    }
                }

                // Restore original clipboard (best-effort)
                try
                {
                    if (OpenClipboard(IntPtr.Zero))
                    {
                        EmptyClipboard();
                        if (!string.IsNullOrEmpty(original))
                        {
                            byte[] origBytes = Encoding.Unicode.GetBytes(original + '\0');
                            IntPtr hGlobal2 = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)origBytes.Length);
                            if (hGlobal2 != IntPtr.Zero)
                            {
                                IntPtr target2 = GlobalLock(hGlobal2);
                                if (target2 != IntPtr.Zero)
                                {
                                    Marshal.Copy(origBytes, 0, target2, origBytes.Length);
                                    GlobalUnlock(hGlobal2);
                                    SetClipboardData(CF_UNICODETEXT, hGlobal2);
                                }
                                else GlobalFree(hGlobal2);
                            }
                        }
                        CloseClipboard();
                    }
                }
                catch { }

                return true;
            }
            catch (Exception ex)
            {
                LogService.Write("TryClipboardPasteString failed: " + ex.Message);
                return false;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
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
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [DllImport("user32.dll")] private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetKeyboardLayout(uint idThread);

        // Helper: try to simulate Ctrl+V targeted at hwnd using AttachThreadInput+PostMessage, with SendInput scancode fallback.
        private bool TrySendCtrlVToWindow(IntPtr hwndTarget)
        {
            try
            {
                long key = hwndTarget.ToInt64();
                lock (_ctrlVLock)
                {
                    if (_lastCtrlVAttempts.TryGetValue(key, out var rec))
                    {
                        if ((DateTime.UtcNow - rec.time) <= TimeSpan.FromMilliseconds(300))
                        {
                            LogService.Write($"TrySendCtrlVToWindow: skipping duplicate attempt for HWND=0x{key:X} (recent success={rec.success})");
                            return rec.success;
                        }
                    }
                }
            }
            catch { }

            try
            {
                const uint WM_KEYDOWN = 0x0100;
                const uint WM_KEYUP = 0x0101;
                const int VK_CONTROL = 0x11;
                const int VK_V = 0x56;

                uint pid;
                uint targetThread = GetWindowThreadProcessId(hwndTarget, out pid);
                uint currentThread = GetCurrentThreadId();
                bool attached = false;
                try
                {
                    if (targetThread != 0 && currentThread != 0)
                    {
                        attached = AttachThreadInput(currentThread, targetThread, true);
                        LogService.Write($"TrySendCtrlVToWindow: AttachThreadInput={attached} (cur={currentThread}, tgt={targetThread})");
                    }

                    // Post Ctrl down, V down, V up, Ctrl up
                    uint scanCtrl = MapVirtualKey((uint)VK_CONTROL, 0);
                    uint scanV = MapVirtualKey((uint)VK_V, 0);
                    uint lParamDownCtrl = 1u | (scanCtrl << 16);
                    uint lParamUpCtrl = 1u | (scanCtrl << 16) | (1u << 30) | (1u << 31);
                    uint lParamDownV = 1u | (scanV << 16);
                    uint lParamUpV = 1u | (scanV << 16) | (1u << 30) | (1u << 31);

                    bool downCtrl = PostMessage(hwndTarget, WM_KEYDOWN, (IntPtr)VK_CONTROL, (IntPtr)lParamDownCtrl);
                    bool downV = PostMessage(hwndTarget, WM_KEYDOWN, (IntPtr)VK_V, (IntPtr)lParamDownV);
                    bool upV = PostMessage(hwndTarget, WM_KEYUP, (IntPtr)VK_V, (IntPtr)lParamUpV);
                    bool upCtrl = PostMessage(hwndTarget, WM_KEYUP, (IntPtr)VK_CONTROL, (IntPtr)lParamUpCtrl);
                    LogService.Write($"Posted Ctrl+V messages: downCtrl={downCtrl}, downV={downV}, upV={upV}, upCtrl={upCtrl}");

                    if (downCtrl && downV && upV && upCtrl)
                    {
                        lock (_ctrlVLock) { _lastCtrlVAttempts[hwndTarget.ToInt64()] = (DateTime.UtcNow, true); }
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    LogService.Write("TrySendCtrlVToWindow PostMessage path failed: " + ex.Message);
                }
                finally
                {
                    if (attached)
                    {
                        try { AttachThreadInput(currentThread, targetThread, false); } catch { }
                    }
                }

                // As a last resort, try SendInput scancode sequence for Ctrl+V globally
                try
                {
                    const uint KEYEVENTF_SCANCODE = 0x0008;
                    const uint KEYEVENTF_KEYUP = 0x0002;
                    uint scanCtrl = MapVirtualKey((uint)VK_CONTROL, 0);
                    uint scanV = MapVirtualKey((uint)VK_V, 0);

                    INPUT[] inputs = new INPUT[4];
                    // Ctrl down
                    inputs[0].type = 1; inputs[0].ki.ki.wVk = 0; inputs[0].ki.ki.wScan = (ushort)scanCtrl; inputs[0].ki.ki.dwFlags = KEYEVENTF_SCANCODE;
                    // V down
                    inputs[1].type = 1; inputs[1].ki.ki.wVk = 0; inputs[1].ki.ki.wScan = (ushort)scanV; inputs[1].ki.ki.dwFlags = KEYEVENTF_SCANCODE;
                    // V up
                    inputs[2].type = 1; inputs[2].ki.ki.wVk = 0; inputs[2].ki.ki.wScan = (ushort)scanV; inputs[2].ki.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
                    // Ctrl up
                    inputs[3].type = 1; inputs[3].ki.ki.wVk = 0; inputs[3].ki.ki.wScan = (ushort)scanCtrl; inputs[3].ki.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;

                    uint sent = SendInput(4, inputs, INPUT.Size);
                    LogService.Write($"SendInput Ctrl+V scancode sent={sent}");
                    bool ok = sent == 4;
                    lock (_ctrlVLock) { _lastCtrlVAttempts[hwndTarget.ToInt64()] = (DateTime.UtcNow, ok); }
                    return ok;
                }
                catch (Exception ex)
                {
                    LogService.Write("TrySendCtrlVToWindow SendInput fallback failed: " + ex.Message);
                }
            }
            catch (Exception ex) { LogService.Write("TrySendCtrlVToWindow failed: " + ex.Message); }
            return false;
        }

        public bool ForceToKorean()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                LogService.Write($"ForceToKorean invoked. Foreground HWND: 0x{hwnd.ToString("X")}");

                INPUT[] inputs = new INPUT[2];
                inputs[0].type = 1;
                inputs[0].ki.ki.wVk = VK_HANGUL;
                inputs[0].ki.ki.dwFlags = 0;
                inputs[1].type = 1;
                inputs[1].ki.ki.wVk = VK_HANGUL;
                inputs[1].ki.ki.dwFlags = 2;

                uint res = SendInput(2, inputs, INPUT.Size);
                if (res == 2) LogService.Write("SendInput for VK_HANGUL reported success.");

                IntPtr hkl = LoadKeyboardLayout("00000412", 0);
                if (hkl != IntPtr.Zero) LogService.Write($"Activated Korean layout via LoadKeyboardLayout: 0x{hkl.ToInt64():X}");

                const uint WM_INPUTLANGCHANGEREQUEST = 0x0050;
                PostMessage(hwnd, WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, hkl);

                IntPtr layout = GetKeyboardLayout(0);
                bool isKorean = (layout.ToInt64() & 0xFFFF) == 0x412;
                LogService.Write($"Post-send check attempt, isKorean={isKorean}");
                return isKorean;
            }
            catch (Exception ex)
            {
                LogService.Write("ForceToKorean failed: " + ex.Message);
                return false;
            }
        }

        public bool SendVirtualKey(ushort vk)
        {
            try
            {
                INPUT[] inputs = new INPUT[2];
                inputs[0].type = 1;
                inputs[0].ki.ki.wVk = vk;
                inputs[0].ki.ki.dwFlags = 0;
                inputs[1].type = 1;
                inputs[1].ki.ki.wVk = vk;
                inputs[1].ki.ki.dwFlags = 2;
                uint res = SendInput(2, inputs, INPUT.Size);
                return res == 2;
            }
            catch { return false; }
        }

        // Send Enter using scancode SendInput as a more hardware-like fallback
        public bool SendEnterScancode()
        {
            try
            {
                const uint KEYEVENTF_SCANCODE = 0x0008;
                const uint KEYEVENTF_KEYUP = 0x0002;
                const int VK_RETURN = 0x0D;
                uint scan = MapVirtualKey((uint)VK_RETURN, 0);

                INPUT[] inputs = new INPUT[2];
                inputs[0].type = 1;
                inputs[0].ki.ki.wVk = 0;
                inputs[0].ki.ki.wScan = (ushort)scan;
                inputs[0].ki.ki.dwFlags = KEYEVENTF_SCANCODE;

                inputs[1].type = 1;
                inputs[1].ki.ki.wVk = 0;
                inputs[1].ki.ki.wScan = (ushort)scan;
                inputs[1].ki.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;

                uint sent = SendInput(2, inputs, INPUT.Size);
                LogService.Write($"SendEnterScancode: SendInput(scancode) scan={scan} sent={sent}");
                return sent == 2;
            }
            catch (Exception ex)
            {
                LogService.Write("SendEnterScancode failed: " + ex.Message);
                return false;
            }
        }

        // Last-resort: post WM_KEYDOWN/WM_KEYUP for Enter to a specific window using AttachThreadInput to target thread.
        public bool TryPostEnterToWindow(IntPtr hwndTarget)
        {
            try
            {
                const uint WM_KEYDOWN = 0x0100;
                const uint WM_KEYUP = 0x0101;
                const int VK_RETURN = 0x0D;

                uint pid;
                uint targetThread = GetWindowThreadProcessId(hwndTarget, out pid);
                uint currentThread = GetCurrentThreadId();
                bool attached = false;
                try
                {
                    if (targetThread != 0 && currentThread != 0)
                    {
                        attached = AttachThreadInput(currentThread, targetThread, true);
                        LogService.Write($"TryPostEnterToWindow: AttachThreadInput={attached} (cur={currentThread}, tgt={targetThread})");
                    }

                    bool down = PostMessage(hwndTarget, WM_KEYDOWN, (IntPtr)VK_RETURN, IntPtr.Zero);
                    bool up = PostMessage(hwndTarget, WM_KEYUP, (IntPtr)VK_RETURN, IntPtr.Zero);
                    LogService.Write($"TryPostEnterToWindow: posted down={down}, up={up} to HWND=0x{hwndTarget.ToInt64():X}");

                    return down && up;
                }
                catch (Exception ex)
                {
                    LogService.Write("TryPostEnterToWindow failed: " + ex.Message);
                    return false;
                }
                finally
                {
                    if (attached)
                    {
                        try { AttachThreadInput(currentThread, targetThread, false); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.Write("TryPostEnterToWindow outer failed: " + ex.Message);
                return false;
            }
        }

        public bool SendUnicodeChar(char ch)
        {
            try
            {
                const uint KEYEVENTF_UNICODE = 0x0004;
                const uint KEYEVENTF_KEYUP = 0x0002;

                INPUT[] inputs = new INPUT[2];
                inputs[0].type = 1;
                inputs[0].ki.ki.wVk = 0;
                inputs[0].ki.ki.wScan = (ushort)ch;
                inputs[0].ki.ki.dwFlags = KEYEVENTF_UNICODE;

                inputs[1].type = 1;
                inputs[1].ki.ki.wVk = 0;
                inputs[1].ki.ki.wScan = (ushort)ch;
                inputs[1].ki.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

                uint res = SendInput(2, inputs, INPUT.Size);
                LogService.Write($"SendUnicodeChar('{ch}') sent={res}");
                if (res == 2) return true;

                int err = Marshal.GetLastWin32Error();
                LogService.Write($"SendUnicodeChar: SendInput failed (res={res}), GetLastError={err}");

                // Diagnostics: log CapsLock / Shift states
                const int VK_SHIFT = 0x10;
                const int VK_LSHIFT = 0xA0;
                const int VK_RSHIFT = 0xA1;
                const int VK_CAPITAL = 0x14;
                bool caps = (GetKeyState(VK_CAPITAL) & 0x0001) != 0;
                bool shift = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0 || (GetAsyncKeyState(VK_LSHIFT) & 0x8000) != 0 || (GetAsyncKeyState(VK_RSHIFT) & 0x8000) != 0;
                LogService.Write($"SendUnicodeChar diagnostics: CapsLock={caps}, ShiftDown={shift}");

                

                // Fallback: try posting WM_CHAR to the foreground window
                const uint WM_CHAR = 0x0102;
                IntPtr hwndForeground = GetForegroundWindow();

                // Try to get focused control (hwndFocus) for the foreground thread
                IntPtr hwndTarget = hwndForeground;
                try
                {
                    uint pid;
                    uint threadId = GetWindowThreadProcessId(hwndForeground, out pid);
                    GUITHREADINFO gti = new GUITHREADINFO();
                    gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                    if (threadId != 0 && GetGUIThreadInfo(threadId, ref gti))
                    {
                        if (gti.hwndFocus != IntPtr.Zero)
                        {
                            hwndTarget = gti.hwndFocus;
                            LogService.Write($"GetGUIThreadInfo: hwndFocus=0x{gti.hwndFocus.ToInt64():X}, hwndCaret=0x{gti.hwndCaret.ToInt64():X}");
                        }
                    }
                }
                catch { }

                // Try SendMessage WM_UNICHAR to the focused control first (A)
                const uint WM_UNICHAR = 0x0109;
                string procName = string.Empty;
                try { uint pid; GetWindowThreadProcessId(hwndTarget, out pid); if (pid != 0) procName = Process.GetProcessById((int)pid).ProcessName; } catch { }
                LogService.Write($"SendMessage WM_UNICHAR('{ch}') to target HWND=0x{hwndTarget.ToInt64():X} (process='{procName}')");
                
                // If target process looks like Notion, prefer to buffer the jamo and avoid immediate single-char paste.
                try
                {
                    if (!string.IsNullOrEmpty(procName) && procName.IndexOf("notion", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        LogService.Write("Target is Notion: using buffered clipboard compose flow (delay single-char paste).");
                        // If we have a recent pending jamo and the new char is a compatibility jamo, try composing them (use hwndTarget now)
                        if (_pendingJamo.HasValue && IsCompatibilityJamo(_pendingJamo.Value) && IsCompatibilityJamo(ch) && (DateTime.UtcNow - _pendingTime) <= _composeWindow)
                        {
                            string? composed = TryComposePair(_pendingJamo.Value, ch);
                            if (!string.IsNullOrEmpty(composed))
                            {
                                LogService.Write($"Notion: composing pending jamo '{_pendingJamo}' + '{ch}' -> '{composed}' and replacing previous pasted chars");
                                TrySendBackspaces(hwndTarget, _pendingPastedCount);
                                if (TryClipboardPasteString(hwndTarget, composed))
                                {
                                    _pendingJamo = null; _pendingPastedCount = 0; return true;
                                }
                            }
                        }

                        // Do NOT paste single char immediately for Notion; instead buffer it so the UI timer can decide when to flush.
                        if (IsCompatibilityJamo(ch))
                        {
                            _pendingJamo = ch;
                            _pendingTime = DateTime.UtcNow;
                            _pendingPastedCount = 0; // not actually pasted yet
                            LogService.Write($"Notion: buffered single jamo '{ch}' for possible composition (no immediate paste).");
                            return true;
                        }
                        // if not a compatibility jamo, fall through to normal flow
                    }
                }
                catch (Exception ex)
                {
                    LogService.Write("Notion buffered-flow attempt failed: " + ex.Message);
                }

                // If we have a recent pending jamo and the new char is a compatibility jamo, try composing them (use hwndTarget now)
                if (_pendingJamo.HasValue && IsCompatibilityJamo(_pendingJamo.Value) && IsCompatibilityJamo(ch) && (DateTime.UtcNow - _pendingTime) <= _composeWindow)
                {
                    string? composed = TryComposePair(_pendingJamo.Value, ch);
                    if (!string.IsNullOrEmpty(composed))
                    {
                        LogService.Write($"Composing pending jamo '{_pendingJamo}' + '{ch}' -> '{composed}' and replacing previous pasted chars");
                        // send backspaces to delete previously pasted chars
                        TrySendBackspaces(hwndTarget, _pendingPastedCount);
                        // paste composed string
                        if (TryClipboardPasteString(hwndTarget, composed))
                        {
                            _pendingJamo = null; _pendingPastedCount = 0; return true;
                        }
                        // if paste failed, continue to other fallbacks
                    }
                }
                IntPtr smRes = SendMessage(hwndTarget, WM_UNICHAR, (IntPtr)ch, IntPtr.Zero);
                LogService.Write($"SendMessage WM_UNICHAR result=0x{smRes.ToInt64():X}");

                bool handled = smRes != IntPtr.Zero;

                // If WM_UNICHAR wasn't handled, try AttachThreadInput + SendMessage WM_CHAR (B)
                if (!handled)
                {
                    const uint WM_CHAR_LOCAL = 0x0102;
                    uint pid;
                    uint targetThread = GetWindowThreadProcessId(hwndTarget, out pid);
                    uint currentThread = GetCurrentThreadId();
                    bool attached = false;
                    try
                    {
                        if (targetThread != 0 && currentThread != 0)
                        {
                            attached = AttachThreadInput(currentThread, targetThread, true);
                            LogService.Write($"AttachThreadInput result={attached} (current={currentThread}, target={targetThread})");
                        }

                        LogService.Write($"SendMessage WM_CHAR('{ch}') to target HWND=0x{hwndTarget.ToInt64():X} (process='{procName}') with AttachThreadInput={(attached ? "yes" : "no")}");
                        IntPtr smc = SendMessage(hwndTarget, WM_CHAR_LOCAL, (IntPtr)ch, IntPtr.Zero);
                        LogService.Write($"SendMessage WM_CHAR result=0x{smc.ToInt64():X}");

                        // If SendMessage indicates handled (non-zero), consider success
                        handled = smc != IntPtr.Zero;
                    }
                    catch (Exception ex)
                    {
                        LogService.Write("Attach/SendMessage WM_CHAR failed: " + ex.Message);
                    }
                    finally
                    {
                        if (attached)
                        {
                            try { AttachThreadInput(currentThread, targetThread, false); } catch { }
                        }
                    }

                    // If still not handled, try posting WM_KEYDOWN/WM_KEYUP for the corresponding virtual key so the IME receives real key events
                    if (!handled)
                    {
                        // Map jamo char to virtual-key using two-beolsik mapping
                        ushort vk = 0;
                        switch (ch)
                        {
                            // Consonants
                            case '\u3131': vk = (ushort)0x52; break; // R -> Keys.R
                            case '\u3134': vk = (ushort)0x53; break; // S
                            case '\u3137': vk = (ushort)0x45; break; // E
                            case '\u3139': vk = (ushort)0x46; break; // F
                            case '\u3141': vk = (ushort)0x41; break; // A
                            case '\u3142': vk = (ushort)0x51; break; // Q
                            case '\u3145': vk = (ushort)0x54; break; // T
                            case '\u3147': vk = (ushort)0x44; break; // D
                            case '\u3148': vk = (ushort)0x57; break; // W
                            case '\u314A': vk = (ushort)0x43; break; // C
                            case '\u314B': vk = (ushort)0x5A; break; // Z
                            case '\u314C': vk = (ushort)0x58; break; // X
                            case '\u314D': vk = (ushort)0x56; break; // V
                            case '\u314E': vk = (ushort)0x47; break; // G
                            // Vowels
                            case '\u314F': vk = (ushort)0x4B; break; // K
                            case '\u3150': vk = (ushort)0x4F; break; // O
                            case '\u3151': vk = (ushort)0x49; break; // I
                            case '\u3153': vk = (ushort)0x4A; break; // J
                            case '\u3154': vk = (ushort)0x50; break; // P
                            case '\u3155': vk = (ushort)0x55; break; // U
                            case '\u3157': vk = (ushort)0x48; break; // H
                            case '\u315B': vk = (ushort)0x59; break; // Y
                            case '\u315C': vk = (ushort)0x4E; break; // N
                            case '\u3160': vk = (ushort)0x42; break; // B
                            case '\u3161': vk = (ushort)0x4D; break; // M
                            case '\u3163': vk = (ushort)0x4C; break; // L
                            default: vk = 0; break;
                        }

                        if (vk != 0)
                        {
                            try
                            {
                                uint scan = MapVirtualKey(vk, 0);
                                uint pid2;
                                uint targetThread2 = GetWindowThreadProcessId(hwndTarget, out pid2);
                                uint currentThread2 = GetCurrentThreadId();
                                bool attached2 = false;
                                try
                                {
                                    if (targetThread2 != 0 && currentThread2 != 0)
                                    {
                                        attached2 = AttachThreadInput(currentThread2, targetThread2, true);
                                        LogService.Write($"AttachThreadInput(for keys) result={attached2} (current={currentThread2}, target={targetThread2})");
                                    }

                                    const uint WM_KEYDOWN = 0x0100;
                                    const uint WM_KEYUP = 0x0101;
                                    uint lParamDown = (1u) | (scan << 16);
                                    uint lParamUp = (1u) | (scan << 16) | (1u << 30) | (1u << 31);

                                    bool downPosted = PostMessage(hwndTarget, WM_KEYDOWN, (IntPtr)vk, (IntPtr)lParamDown);
                                    bool upPosted = PostMessage(hwndTarget, WM_KEYUP, (IntPtr)vk, (IntPtr)lParamUp);
                                    LogService.Write($"Posted WM_KEYDOWN vk=0x{vk:X} scan={scan} downPosted={downPosted}, upPosted={upPosted}");
                                    handled = downPosted && upPosted;
                                }
                                catch (Exception ex)
                                {
                                    LogService.Write("Send key messages failed: " + ex.Message);
                                }
                                finally
                                {
                                    if (attached2)
                                    {
                                        try { AttachThreadInput(currentThread2, targetThread2, false); } catch { }
                                    }
                                }
                                // If posting key messages didn't work, or if the target process is Notion (which often
                                // doesn't perform IME composition for posted messages), try SendInput using scancodes
                                // to better emulate hardware.
                                if (!handled || (procName != null && procName.IndexOf("notion", StringComparison.OrdinalIgnoreCase) >= 0))
                                {
                                    try
                                    {
                                        const uint KEYEVENTF_SCANCODE = 0x0008;
                                        uint scanCode = MapVirtualKey(vk, 0);
                                        INPUT[] sin = new INPUT[2];
                                        sin[0].type = 1;
                                        sin[0].ki.ki.wVk = 0;
                                        sin[0].ki.ki.wScan = (ushort)scanCode;
                                        sin[0].ki.ki.dwFlags = KEYEVENTF_SCANCODE;
                                        sin[1].type = 1;
                                        sin[1].ki.ki.wVk = 0;
                                        sin[1].ki.ki.wScan = (ushort)scanCode;
                                        sin[1].ki.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
                                        uint sent = SendInput(2, sin, INPUT.Size);
                                        LogService.Write($"SendInput(scancode) vk=0x{vk:X} scan={scanCode} sent={sent}");
                                        if (sent == 2)
                                        {
                                            handled = true;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogService.Write("SendInput(scancode) failed: " + ex.Message);
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }

                // Last resort: try clipboard + WM_PASTE (C) before falling back to PostMessage
                try
                {
                    string? original = null;
                    // Read existing clipboard text (if any)
                    try
                    {
                        if (OpenClipboard(IntPtr.Zero))
                        {
                            if (IsClipboardFormatAvailable(CF_UNICODETEXT))
                            {
                                IntPtr hData = GetClipboardData(CF_UNICODETEXT);
                                if (hData != IntPtr.Zero)
                                {
                                    IntPtr ptr = GlobalLock(hData);
                                    if (ptr != IntPtr.Zero)
                                    {
                                        original = Marshal.PtrToStringUni(ptr);
                                        GlobalUnlock(hData);
                                    }
                                }
                            }
                            CloseClipboard();
                        }
                    }
                    catch (Exception ex) { LogService.Write("Clipboard read failed: " + ex.Message); }

                    // Set clipboard to our character
                    bool setOk = false;
                    try
                    {
                        if (OpenClipboard(IntPtr.Zero))
                        {
                            EmptyClipboard();
                            byte[] bytes = Encoding.Unicode.GetBytes(ch.ToString() + '\0');
                            IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)bytes.Length);
                            if (hGlobal != IntPtr.Zero)
                            {
                                IntPtr target = GlobalLock(hGlobal);
                                if (target != IntPtr.Zero)
                                {
                                    Marshal.Copy(bytes, 0, target, bytes.Length);
                                    GlobalUnlock(hGlobal);
                                    if (SetClipboardData(CF_UNICODETEXT, hGlobal) != IntPtr.Zero)
                                    {
                                        setOk = true;
                                    }
                                }
                                else
                                {
                                    GlobalFree(hGlobal);
                                }
                            }
                            CloseClipboard();
                        }
                    }
                    catch (Exception ex) { LogService.Write("Clipboard set failed: " + ex.Message); }

                    if (setOk)
                    {
                        LogService.Write($"Clipboard set to '{ch}', sending WM_PASTE to target HWND=0x{hwndTarget.ToInt64():X}");
                        const uint WM_PASTE = 0x0302;
                            bool pastePosted = PostMessage(hwndTarget, WM_PASTE, IntPtr.Zero, IntPtr.Zero);
                            LogService.Write($"PostMessage WM_PASTE result posted={pastePosted} to HWND=0x{hwndTarget.ToInt64():X}");

                            // record pending pasted jamo for possible composition with next key
                            if (IsCompatibilityJamo(ch))
                            {
                                _pendingJamo = ch;
                                _pendingTime = DateTime.UtcNow;
                                _pendingPastedCount = 1;
                            }

                            // Restore original clipboard (best-effort)
                        try
                        {
                            if (OpenClipboard(IntPtr.Zero))
                            {
                                EmptyClipboard();
                                if (!string.IsNullOrEmpty(original))
                                {
                                    byte[] origBytes = Encoding.Unicode.GetBytes(original + '\0');
                                    IntPtr hGlobal2 = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)origBytes.Length);
                                    if (hGlobal2 != IntPtr.Zero)
                                    {
                                        IntPtr target2 = GlobalLock(hGlobal2);
                                        if (target2 != IntPtr.Zero)
                                        {
                                            Marshal.Copy(origBytes, 0, target2, origBytes.Length);
                                            GlobalUnlock(hGlobal2);
                                            SetClipboardData(CF_UNICODETEXT, hGlobal2);
                                        }
                                        else
                                        {
                                            GlobalFree(hGlobal2);
                                        }
                                    }
                                }
                                CloseClipboard();
                            }
                        }
                        catch (Exception ex) { LogService.Write("Clipboard restore failed: " + ex.Message); }

                        LogService.Write("WM_PASTE attempt finished");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    LogService.Write("Clipboard fallback failed: " + ex.Message);
                }

                // If clipboard paste didn't run/wasn't set, fallback to PostMessage
                LogService.Write($"PostMessage WM_CHAR('{ch}') to target HWND=0x{hwndTarget.ToInt64():X} (process='{procName}') as final fallback");
                bool posted = PostMessage(hwndTarget, WM_CHAR, (IntPtr)ch, IntPtr.Zero);
                LogService.Write($"PostMessage WM_CHAR to target posted={posted}");
                // if we posted a char (non-clipboard), still record pending jamo so we can attempt composition
                if (posted && IsCompatibilityJamo(ch))
                {
                    _pendingJamo = ch;
                    _pendingTime = DateTime.UtcNow;
                    _pendingPastedCount = 1;
                }
                if (!posted && hwndTarget != hwndForeground)
                {
                    try { uint pid; GetWindowThreadProcessId(hwndForeground, out pid); if (pid != 0) procName = Process.GetProcessById((int)pid).ProcessName; } catch { }
                    LogService.Write($"PostMessage WM_CHAR('{ch}') to foreground HWND=0x{hwndForeground.ToInt64():X} (process='{procName}') as final fallback");
                    posted = PostMessage(hwndForeground, WM_CHAR, (IntPtr)ch, IntPtr.Zero);
                    LogService.Write($"PostMessage WM_CHAR to foreground posted={posted}");
                }
                return posted;
            }
            catch (Exception ex)
            {
                LogService.Write("SendUnicodeChar failed: " + ex.Message);
                return false;
            }
        }

        // Public helper: compose two compatibility jamo into a precomposed syllable and paste it to foreground/focused control.
        public bool ComposeAndPastePair(char a, char b)
        {
            try
            {
                if (!IsCompatibilityJamo(a) || !IsCompatibilityJamo(b)) return false;
                // If both are vowels, try to compose vowel pair first
                int? vA = CompatJamoToVIndex(a);
                int? vB = CompatJamoToVIndex(b);
                string? composed = null;
                if (vA.HasValue && vB.HasValue)
                {
                    char? compV = TryComposeVowelPair(a, b);
                    if (compV.HasValue)
                    {
                        composed = compV.Value.ToString();
                    }
                }

                // Otherwise try consonant+vowel composition
                if (composed == null)
                {
                    composed = TryComposePair(a, b);
                    if (string.IsNullOrEmpty(composed)) return false;
                }

                // determine target hwnd (focused control of foreground)
                IntPtr hwndForeground = GetForegroundWindow();
                IntPtr hwndTarget = hwndForeground;
                try
                {
                    uint pid;
                    uint threadId = GetWindowThreadProcessId(hwndForeground, out pid);
                    GUITHREADINFO gti = new GUITHREADINFO();
                    gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                    if (threadId != 0 && GetGUIThreadInfo(threadId, ref gti))
                    {
                        if (gti.hwndFocus != IntPtr.Zero)
                        {
                            hwndTarget = gti.hwndFocus;
                        }
                    }
                }
                catch { }

                LogService.Write($"ComposeAndPastePair: composing '{a}' + '{b}' -> '{composed}' and pasting to HWND=0x{hwndTarget.ToInt64():X}");

                // try to paste composed string using existing helper
                bool ok = TryClipboardPasteString(hwndTarget, composed);
                if (ok)
                {
                    // reset any pending state since we directly pasted composed syllable
                    _pendingJamo = null;
                    _pendingPastedCount = 0;
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogService.Write("ComposeAndPastePair failed: " + ex.Message);
            }
            return false;
        }

        // Public helper: paste a string to the focused control of the foreground window via clipboard (with Ctrl+V fallback).
        public bool PasteStringToForeground(string s)
        {
            try
            {
                IntPtr hwndForeground = GetForegroundWindow();
                IntPtr hwndTarget = hwndForeground;
                try
                {
                    uint pid;
                    uint threadId = GetWindowThreadProcessId(hwndForeground, out pid);
                    GUITHREADINFO gti = new GUITHREADINFO();
                    gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
                    if (threadId != 0 && GetGUIThreadInfo(threadId, ref gti))
                    {
                        if (gti.hwndFocus != IntPtr.Zero)
                        {
                            hwndTarget = gti.hwndFocus;
                        }
                    }
                }
                catch { }

                LogService.Write($"PasteStringToForeground: pasting '{s}' to HWND=0x{hwndTarget.ToInt64():X}");
                bool ok = TryClipboardPasteString(hwndTarget, s);
                LogService.Write($"PasteStringToForeground: TryClipboardPasteString result={ok}");
                return ok;
            }
            catch (Exception ex)
            {
                LogService.Write("PasteStringToForeground failed: " + ex.Message);
                return false;
            }
        }

        [StructLayout(LayoutKind.Sequential)] 
        public struct INPUT { 
            public uint type; 
            public InputUnion ki; 
            public static int Size => Marshal.SizeOf(typeof(INPUT)); 
        }
        [StructLayout(LayoutKind.Explicit)] 
        public struct InputUnion { 
            [FieldOffset(0)] public KEYBDINPUT ki; 
        }
        [StructLayout(LayoutKind.Sequential)] 
        public struct KEYBDINPUT { 
            public ushort wVk; 
            public ushort wScan; 
            public uint dwFlags; 
            public uint time; 
            public UIntPtr dwExtraInfo; 
        }
    }
}
