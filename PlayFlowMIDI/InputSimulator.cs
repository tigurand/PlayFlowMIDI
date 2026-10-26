using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PlayFlowMIDI
{
    public enum InputDeliveryMode
    {
        /// <summary>Original behavior: SendInput while focused, plain window messages while unfocused.</summary>
        Legacy,
        /// <summary>Focus-Free behavior: messages are delivered to the game window together with activation notifications.</summary>
        FocusFree,
        /// <summary>Opt-in background injection: synthetic raw keyboard events are fed by an injected hook.</summary>
        Injected
    }

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
        const uint MAPVK_VK_TO_VSC_EX = 4;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        static extern IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll")]
        static extern short MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
                                                uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        public const uint WM_KEYDOWN = 0x0100;
        public const uint WM_KEYUP = 0x0101;
        public const uint WM_SYSKEYDOWN = 0x0104;
        public const uint WM_SYSKEYUP = 0x0105;
        public const uint WM_ACTIVATE = 0x0006;
        const uint WA_ACTIVE = 1;
        const uint WA_INACTIVE = 0;
        const uint SMTO_BLOCK = 0x0001;
        const uint SMTO_ABORTIFHUNG = 0x0002;

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

        public static void SimulateKeyDown(string keyCombo, IntPtr hWnd = default, InputDeliveryMode mode = InputDeliveryMode.Legacy)
        {
            SimulateKeyAction(keyCombo, true, hWnd, mode);
        }

        public static void SimulateKeyUp(string keyCombo, IntPtr hWnd = default, InputDeliveryMode mode = InputDeliveryMode.Legacy)
        {
            SimulateKeyAction(keyCombo, false, hWnd, mode);
        }

        private static void SimulateKeyAction(string keyCombo, bool isDown, IntPtr hWnd, InputDeliveryMode mode)
        {
            if (!TryParseKeyCombo(keyCombo, out Keys mainKey, out List<Keys> modifiers)) return;

            if (hWnd != IntPtr.Zero)
            {
                BackgroundKeyAction(hWnd, mainKey, modifiers, isDown, mode == InputDeliveryMode.FocusFree);
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

        public static void SimulateKeyPress(string keyCombo, IntPtr hWnd = default, InputDeliveryMode mode = InputDeliveryMode.Legacy)
        {
            if (!TryParseKeyCombo(keyCombo, out Keys mainKey, out List<Keys> modifiers)) return;

            if (hWnd != IntPtr.Zero)
            {
                BackgroundKeyPress(hWnd, mainKey, modifiers, mode == InputDeliveryMode.FocusFree);
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

        public static bool TryParseKeyCombo(string keyCombo, out Keys mainKey, out List<Keys> modifiers)
        {
            mainKey = Keys.None;
            modifiers = new List<Keys>();
            if (string.IsNullOrWhiteSpace(keyCombo)) return false;

            foreach (var part in keyCombo.Split('+'))
            {
                string trimmedPart = part.Trim();
                string p = trimmedPart.ToLowerInvariant();
                if (p == "ctrl") modifiers.Add(Keys.LControlKey);
                else if (p == "shift") modifiers.Add(Keys.LShiftKey);
                else if (p == "alt") modifiers.Add(Keys.LMenu);
                else if (trimmedPart.Length == 1 && char.IsDigit(trimmedPart[0]))
                {
                    mainKey = (Keys)trimmedPart[0];
                }
                else if (Enum.TryParse<Keys>(trimmedPart, true, out var k))
                {
                    mainKey = k;
                }
                else if (trimmedPart.Length == 1)
                {
                    char c = char.ToUpperInvariant(trimmedPart[0]);
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

            return mainKey != Keys.None || modifiers.Count > 0;
        }

        /// <summary>
        /// Notifies a window that it is active without taking the foreground window.
        /// Matches the sequence used by working NTE background-input tools.
        /// </summary>
        public static void ActivateBackgroundWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            SendMessageTimeout(hWnd, WM_ACTIVATE, (IntPtr)WA_ACTIVE, IntPtr.Zero,
                               SMTO_BLOCK | SMTO_ABORTIFHUNG, 50, out _);
        }

        /// <summary>
        /// Lightweight activation notification sent once per chord, before its key-downs.
        /// The synchronous activation above is sent at playback start, on focus changes,
        /// and periodically while playing.
        /// </summary>
        public static void NotifyBackgroundWindowActive(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            PostMessage(hWnd, WM_ACTIVATE, (IntPtr)WA_ACTIVE, IntPtr.Zero);
        }

        /// <summary>
        /// Restores the "inactive" notification state after playback stops in window-message mode.
        /// Only call this while the window is not actually in the foreground.
        /// </summary>
        public static void DeactivateBackgroundWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;
            SendMessageTimeout(hWnd, WM_ACTIVATE, (IntPtr)WA_INACTIVE, IntPtr.Zero,
                               SMTO_BLOCK | SMTO_ABORTIFHUNG, 50, out _);
        }

        public static void BackgroundKeyAction(IntPtr hWnd, Keys key, List<Keys> modifiers, bool isDown, bool enhanced = false)
        {
            if (enhanced) EnhancedBackgroundKeyAction(hWnd, key, modifiers, isDown);
            else LegacyBackgroundKeyAction(hWnd, key, modifiers, isDown);
        }

        private static void LegacyBackgroundKeyAction(IntPtr hWnd, Keys key, List<Keys> modifiers, bool isDown)
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

        private static void EnhancedBackgroundKeyAction(IntPtr hWnd, Keys key, List<Keys> modifiers, bool isDown)
        {
            bool hasAlt = modifiers.Contains(Keys.LMenu) || modifiers.Contains(Keys.RMenu) || modifiers.Contains(Keys.Menu);
            uint msgDown = hasAlt ? WM_SYSKEYDOWN : WM_KEYDOWN;
            uint msgUp = hasAlt ? WM_SYSKEYUP : WM_KEYUP;

            if (isDown)
            {
                // Activation is sent once per chord by the caller, not before every key.
                // Posting WM_ACTIVATE between the keys of one chord can make the game
                // re-evaluate its held-key state and turn the chord into separate notes.
                foreach (var mod in modifiers)
                {
                    PostMessage(hWnd, msgDown, (IntPtr)mod, BuildEnhancedKeyLParam(mod, false, hasAlt));
                }
                if (key != Keys.None)
                {
                    PostMessage(hWnd, msgDown, (IntPtr)key, BuildEnhancedKeyLParam(key, false, hasAlt));
                }
            }
            else
            {
                if (key != Keys.None)
                {
                    PostMessage(hWnd, msgUp, (IntPtr)key, BuildEnhancedKeyLParam(key, true, hasAlt));
                }
                modifiers.Reverse();
                foreach (var mod in modifiers)
                {
                    PostMessage(hWnd, msgUp, (IntPtr)mod, BuildEnhancedKeyLParam(mod, true, hasAlt));
                }
            }
        }

        private static IntPtr BuildEnhancedKeyLParam(Keys key, bool isKeyUp, bool altContext)
        {
            uint mapped = (uint)MapVirtualKey((uint)key, MAPVK_VK_TO_VSC_EX);
            uint scanCode = mapped & 0xFF;
            if (scanCode == 0) scanCode = (uint)MapVirtualKey((uint)key, 0);
            bool extended = ((mapped >> 8) & 0xFF) == 0xE0 || IsExtendedKey(key);

            uint lParam = 1; // repeat count
            lParam |= (scanCode & 0xFF) << 16;
            if (extended) lParam |= 1u << 24;
            if (altContext) lParam |= 1u << 29;
            if (isKeyUp) lParam |= (1u << 30) | (1u << 31);
            return unchecked((IntPtr)(int)lParam);
        }

        private static bool IsExtendedKey(Keys key)
        {
            return key == Keys.Insert || key == Keys.Delete || key == Keys.Home || key == Keys.End ||
                   key == Keys.PageUp || key == Keys.PageDown || key == Keys.Up || key == Keys.Down ||
                   key == Keys.Left || key == Keys.Right || key == Keys.NumLock || key == Keys.Divide ||
                   key == Keys.PrintScreen || key == Keys.Apps || key == Keys.RControlKey ||
                   key == Keys.RMenu || key == Keys.LWin || key == Keys.RWin;
        }

        public static void BackgroundKeyPress(IntPtr hWnd, Keys key, List<Keys> modifiers, bool enhanced = false)
        {
            BackgroundKeyAction(hWnd, key, modifiers, true, enhanced);
            if (enhanced)
            {
                // Give the window a moment to observe the key-down before the key-up arrives.
                System.Threading.Thread.Sleep(2);
            }
            BackgroundKeyAction(hWnd, key, modifiers, false, enhanced);
        }

        /// <summary>
        /// Sends one key transition without any implicit modifiers. Used by chord batching
        /// so each note can be sent with exactly its own modifier state.
        /// </summary>
        public static void SendSingleKeyState(IntPtr hWnd, Keys key, bool isDown,
                                              InputDeliveryMode mode, bool altContext = false)
        {
            if (key == Keys.None) return;

            if (hWnd == IntPtr.Zero)
            {
                if (isDown) SendKeyDown(key);
                else SendKeyUp(key);
                return;
            }

            uint message = isDown
                ? (altContext ? WM_SYSKEYDOWN : WM_KEYDOWN)
                : (altContext ? WM_SYSKEYUP : WM_KEYUP);

            if (mode == InputDeliveryMode.FocusFree)
            {
                PostMessage(hWnd, message, (IntPtr)key, BuildEnhancedKeyLParam(key, !isDown, altContext));
                return;
            }

            uint scanCode = (uint)MapVirtualKey((uint)key, 0);
            uint lParam = 1u | ((scanCode & 0xFF) << 16);
            if (altContext) lParam |= 1u << 29;
            if (!isDown) lParam |= (1u << 30) | (1u << 31);
            PostMessage(hWnd, message, (IntPtr)key, unchecked((IntPtr)(int)lParam));
        }

        public static void ReleaseAllKeys(IntPtr hWnd = default, InputDeliveryMode mode = InputDeliveryMode.Legacy)
        {
            Keys[] modifiers = {
                Keys.LControlKey, Keys.RControlKey,
                Keys.LShiftKey, Keys.RShiftKey,
                Keys.LMenu, Keys.RMenu,
                Keys.LWin, Keys.RWin
            };

            bool enhanced = mode == InputDeliveryMode.FocusFree;
            foreach (var key in modifiers)
            {
                if (hWnd != IntPtr.Zero)
                    BackgroundKeyAction(hWnd, key, new List<Keys>(), false, enhanced);
                else
                    SendKeyUp(key);
            }
        }
    }
}
