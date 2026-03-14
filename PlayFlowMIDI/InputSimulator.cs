using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PlayFlowMIDI
{
    public static class InputSimulator
    {
        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public uint type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        const uint INPUT_KEYBOARD = 1;
        const uint KEYEVENTF_KEYDOWN = 0x0000;
        const uint KEYEVENTF_KEYUP = 0x0002;
        const uint KEYEVENTF_SCANCODE = 0x0008;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        static extern IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll")]
        static extern short MapVirtualKey(uint uCode, uint uMapType);

        public static void SendKeyDown(Keys key)
        {
            SendKey(key, KEYEVENTF_KEYDOWN);
        }

        public static void SendKeyUp(Keys key)
        {
            SendKey(key, KEYEVENTF_KEYUP);
        }

        private static void SendKey(Keys key, uint flags)
        {
            ushort scanCode = (ushort)MapVirtualKey((uint)key, 0);
            
            INPUT[] inputs = new INPUT[]
            {
                new INPUT
                {
                    type = INPUT_KEYBOARD,
                    u = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = (ushort)key,
                            wScan = scanCode,
                            dwFlags = flags | KEYEVENTF_SCANCODE,
                            time = 0,
                            dwExtraInfo = GetMessageExtraInfo()
                        }
                    }
                }
            };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void SimulateKeyDown(string keyCombo, IntPtr hWnd = default)
        {
            SimulateKeyAction(keyCombo, true, hWnd);
        }

        public static void SimulateKeyUp(string keyCombo, IntPtr hWnd = default)
        {
            SimulateKeyAction(keyCombo, false, hWnd);
        }

        private static void SimulateKeyAction(string keyCombo, bool isDown, IntPtr hWnd = default)
        {
            if (string.IsNullOrWhiteSpace(keyCombo)) return;

            var parts = keyCombo.Split('+');
            List<Keys> modifiers = new List<Keys>();
            Keys mainKey = Keys.None;

            foreach (var part in parts)
            {
                string p = part.Trim().ToLower();
                if (p == "ctrl") modifiers.Add(Keys.LControlKey);
                else if (p == "shift") modifiers.Add(Keys.LShiftKey);
                else if (p == "alt") modifiers.Add(Keys.LMenu);
                else
                {
                    string trimmedPart = part.Trim();
                    if (trimmedPart.Length == 1 && char.IsDigit(trimmedPart[0]))
                    {
                        mainKey = (Keys)trimmedPart[0];
                    }
                    else if (Enum.TryParse<Keys>(trimmedPart, true, out var k))
                    {
                        mainKey = k;
                    }
                    else if (trimmedPart.Length == 1)
                    {
                        char c = trimmedPart.ToUpper()[0];
                        mainKey = c switch
                        {
                            ',' => Keys.Oemcomma,
                            '.' => Keys.OemPeriod,
                            '/' => Keys.OemQuestion,
                            ';' => Keys.Oem1,
                            '\'' => Keys.Oem7,
                            '[' => Keys.OemOpenBrackets,
                            ']' => Keys.Oem6,
                            '\\' => Keys.Oem5,
                            '-' => Keys.OemMinus,
                            '=' => Keys.Oemplus,
                            '`' => Keys.Oemtilde,
                            _ => (Keys)c
                        };
                    }
                }
            }

            if (mainKey == Keys.None && modifiers.Count == 0) return;

            if (hWnd != IntPtr.Zero)
            {
                BackgroundKeyAction(hWnd, mainKey, modifiers, isDown);
            }
            else
            {
                if (isDown)
                {
                    foreach (var mod in modifiers) SendKeyDown(mod);
                    if (mainKey != Keys.None) SendKeyDown(mainKey);
                }
                else
                {
                    if (mainKey != Keys.None) SendKeyUp(mainKey);
                    modifiers.Reverse();
                    foreach (var mod in modifiers) SendKeyUp(mod);
                }
            }
        }

        public static void SimulateKeyPress(string keyCombo, IntPtr hWnd = default)
        {
            if (string.IsNullOrWhiteSpace(keyCombo)) return;

            var parts = keyCombo.Split('+');
            List<Keys> modifiers = new List<Keys>();
            Keys mainKey = Keys.None;

            foreach (var part in parts)
            {
                string p = part.Trim().ToLower();
                if (p == "ctrl") modifiers.Add(Keys.LControlKey);
                else if (p == "shift") modifiers.Add(Keys.LShiftKey);
                else if (p == "alt") modifiers.Add(Keys.LMenu);
                else
                {
                    string trimmedPart = part.Trim();
                    if (trimmedPart.Length == 1 && char.IsDigit(trimmedPart[0]))
                    {
                        mainKey = (Keys)trimmedPart[0];
                    }
                    else if (Enum.TryParse<Keys>(trimmedPart, true, out var k))
                    {
                        mainKey = k;
                    }
                    else if (trimmedPart.Length == 1)
                    {
                        char c = trimmedPart.ToUpper()[0];
                        mainKey = c switch
                        {
                            ',' => Keys.Oemcomma,
                            '.' => Keys.OemPeriod,
                            '/' => Keys.OemQuestion,
                            ';' => Keys.Oem1,
                            '\'' => Keys.Oem7,
                            '[' => Keys.OemOpenBrackets,
                            ']' => Keys.Oem6,
                            '\\' => Keys.Oem5,
                            '-' => Keys.OemMinus,
                            '=' => Keys.Oemplus,
                            '`' => Keys.Oemtilde,
                            _ => (Keys)c
                        };
                    }
                }
            }

            if (mainKey == Keys.None && modifiers.Count == 0) return;

            if (hWnd != IntPtr.Zero)
            {
                BackgroundKeyAction(hWnd, mainKey, modifiers, true);
                BackgroundKeyAction(hWnd, mainKey, modifiers, false);
            }
            else
            {
                foreach (var mod in modifiers) SendKeyDown(mod);
                if (mainKey != Keys.None)
                {
                    SendKeyDown(mainKey);
                    System.Threading.Thread.Sleep(1);
                    SendKeyUp(mainKey);
                }
                modifiers.Reverse();
                foreach (var mod in modifiers) SendKeyUp(mod);
            }
        }
        
        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public const uint WM_KEYDOWN = 0x0100;
        public const uint WM_KEYUP = 0x0101;
        public const uint WM_SYSKEYDOWN = 0x0104;
        public const uint WM_SYSKEYUP = 0x0105;

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public static void BackgroundKeyAction(IntPtr hWnd, Keys key, List<Keys> modifiers, bool isDown)
        {
            bool hasAlt = modifiers.Contains(Keys.LMenu) || modifiers.Contains(Keys.RMenu) || modifiers.Contains(Keys.Menu);
            uint msgDown = hasAlt ? WM_SYSKEYDOWN : WM_KEYDOWN;
            uint msgUp = hasAlt ? WM_SYSKEYUP : WM_KEYUP;

            if (isDown)
            {
                foreach (var mod in modifiers)
                {
                    uint modScanCode = (uint)MapVirtualKey((uint)mod, 0);
                    IntPtr modLParam = (IntPtr)(1 | (modScanCode << 16));
                    if (mod == Keys.LMenu || mod == Keys.RMenu || mod == Keys.Menu)
                        modLParam = (IntPtr)((int)modLParam | (1 << 29));
                    PostMessage(hWnd, msgDown, (IntPtr)mod, modLParam);
                }
                if (key != Keys.None)
                {
                    uint scanCode = (uint)MapVirtualKey((uint)key, 0);
                    IntPtr lParamDown = (IntPtr)(1 | (scanCode << 16));
                    if (hasAlt) lParamDown = (IntPtr)((int)lParamDown | (1 << 29));
                    PostMessage(hWnd, msgDown, (IntPtr)key, lParamDown);
                }
            }
            else
            {
                if (key != Keys.None)
                {
                    uint scanCode = (uint)MapVirtualKey((uint)key, 0);
                    IntPtr lParamUp = (IntPtr)(1 | (scanCode << 16) | (1u << 30) | (1u << 31));
                    if (hasAlt) lParamUp = (IntPtr)((int)lParamUp | (1 << 29));
                    PostMessage(hWnd, msgUp, (IntPtr)key, lParamUp);
                }
                modifiers.Reverse();
                foreach (var mod in modifiers)
                {
                    uint modScanCode = (uint)MapVirtualKey((uint)mod, 0);
                    IntPtr modLParam = (IntPtr)(1 | (modScanCode << 16) | (1u << 30) | (1u << 31));
                    if (mod == Keys.LMenu || mod == Keys.RMenu || mod == Keys.Menu)
                        modLParam = (IntPtr)((int)modLParam | (1 << 29));
                    PostMessage(hWnd, msgUp, (IntPtr)mod, modLParam);
                }
            }
        }

        public static void BackgroundKeyPress(IntPtr hWnd, Keys key, List<Keys> modifiers)
        {
            BackgroundKeyAction(hWnd, key, modifiers, true);
            BackgroundKeyAction(hWnd, key, modifiers, false);
        }

        public static void ReleaseAllKeys(IntPtr hWnd = default)
        {
            Keys[] modifiers = { 
                Keys.LControlKey, Keys.RControlKey, 
                Keys.LShiftKey, Keys.RShiftKey, 
                Keys.LMenu, Keys.RMenu,
                Keys.LWin, Keys.RWin
            };

            foreach (var key in modifiers)
            {
                if (hWnd != IntPtr.Zero)
                    BackgroundKeyAction(hWnd, key, new List<Keys>(), false);
                else
                    SendKeyUp(key);
            }
        }
    }
}
