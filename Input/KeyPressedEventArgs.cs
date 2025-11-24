using System;
using System.Windows.Forms;

namespace KoreanIMEFixer.Input
{
    public class KeyPressedEventArgs : EventArgs
    {
        public Keys Key { get; }
        public bool Suppress { get; set; }
        public KeyPressedEventArgs(Keys key) { Key = key; Suppress = false; }
    }
}
