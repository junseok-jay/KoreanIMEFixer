using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using KoreanIMEFixer.Logging;

namespace KoreanIMEFixer.Input
{
    public class InputAnalyzer
    {
        private bool _isExpectingKorean;
        private readonly IME.IMEStateMonitor _imeMonitor;
        public event EventHandler? FirstLetterBugDetected;

        public InputAnalyzer() { _imeMonitor = new IME.IMEStateMonitor(); }

        public void OnKeyPressed(Keys key)
        {
            if (TryGetCharFromVirtualKey(key, out char ch))
            {
                bool isLetterKey = ((int)key >= 65 && (int)key <= 90) || char.IsLetter(ch);
                if (!isLetterKey) return;
                bool isKorean = _imeMonitor.IsKoreanMode();
                LogService.Write($"KeyPressed detected: VK={(int)key} ('{ch}'), expecting={_isExpectingKorean}, isKorean={isKorean}");
                if (_isExpectingKorean && !isKorean) FirstLetterBugDetected?.Invoke(this, EventArgs.Empty);
            }
        }

        public bool IsFirstLetterBugCandidate(Keys key)
        {
            if (!_isExpectingKorean) return false;
            if (!TryGetCharFromVirtualKey(key, out char ch)) return false;
            bool isLetterKey = ((int)key >= 65 && (int)key <= 90) || char.IsLetter(ch);
            if (!isLetterKey) return false;
            bool isKorean = _imeMonitor.IsKoreanMode();
            return !isKorean;
        }

        public void BeginTypingExpectingKorean() { _isExpectingKorean = true; }
        public void EndTypingExpectingKorean() { _isExpectingKorean = false; }

        public static bool IsVirtualKeyLetter(Keys key)
        {
            return TryGetCharFromVirtualKey(key, out char c) && char.IsLetter(c);
        }

        /// <summary>
        /// Try map physical key (two-beolsik) to Hangul compatibility jamo.
        /// Returns a compatibility jamo Unicode char (e.g. 'ㅁ') if known.
        /// </summary>
        public static bool TryGetDubeolsikJamo(Keys key, out char jamo)
        {
            jamo = '\0';
            bool shift = (Control.ModifierKeys & Keys.Shift) != 0;
            switch (key)
            {
                case Keys.R: jamo = shift ? '\u3132' : '\u3131'; return true;
                case Keys.S: jamo = '\u3134'; return true;
                case Keys.E: jamo = shift ? '\u3138' : '\u3137'; return true;
                case Keys.F: jamo = '\u3139'; return true;
                case Keys.A: jamo = '\u3141'; return true;
                case Keys.Q: jamo = shift ? '\u3143' : '\u3142'; return true;
                case Keys.T: jamo = shift ? '\u3146' : '\u3145'; return true;
                case Keys.D: jamo = '\u3147'; return true;
                case Keys.W: jamo = shift ? '\u3149' : '\u3148'; return true;
                case Keys.C: jamo = '\u314A'; return true;
                case Keys.Z: jamo = '\u314B'; return true;
                case Keys.X: jamo = '\u314C'; return true;
                case Keys.V: jamo = '\u314D'; return true;
                case Keys.G: jamo = '\u314E'; return true;
                case Keys.K: jamo = '\u314F'; return true;
                case Keys.O: jamo = shift ? '\u3152' : '\u3150'; return true;
                case Keys.I: jamo = '\u3151'; return true;
                case Keys.J: jamo = '\u3153'; return true;
                case Keys.P: jamo = shift ? '\u3156' : '\u3154'; return true;
                case Keys.U: jamo = '\u3155'; return true;
                case Keys.H: jamo = '\u3157'; return true;
                case Keys.Y: jamo = '\u315B'; return true;
                case Keys.N: jamo = '\u315C'; return true;
                case Keys.B: jamo = '\u3160'; return true;
                case Keys.M: jamo = '\u3161'; return true;
                case Keys.L: jamo = '\u3163'; return true;
                default: return false;
            }
        }

        public static bool TryGetCharFromVirtualKey(Keys key, out char result)
        {
            result = '\0';
            try
            {
                byte[] keyboardState = new byte[256];
                if (!GetKeyboardState(keyboardState)) return false;
                StringBuilder sb = new StringBuilder(8);
                uint scanCode = MapVirtualKey((uint)key, 0);
                IntPtr layout = GetKeyboardLayout(0);
                int resultCount = ToUnicodeEx((uint)key, scanCode, keyboardState, sb, sb.Capacity, 0, layout);
                if (resultCount > 0 && sb.Length > 0) { result = char.ToLowerInvariant(sb[0]); return true; }
                return false;
            }
            catch { return false; }
        }

        [DllImport("user32.dll")] private static extern bool GetKeyboardState([Out] byte[] lpKeyState);
        [DllImport("user32.dll")] private static extern int ToUnicodeEx(uint wVirtKey, uint wScanCode, byte[] lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags, IntPtr dwhkl);
        [DllImport("user32.dll")] private static extern uint MapVirtualKey(uint uCode, uint uMapType);
        [DllImport("user32.dll")] private static extern IntPtr GetKeyboardLayout(uint idThread);
    }
}
