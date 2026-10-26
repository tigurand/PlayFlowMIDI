using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Runtime.InteropServices;
using System.Threading;
using Forms = System.Windows.Forms;

namespace PlayFlowMIDI
{
    /// <summary>
    /// Opt-in Background Injection controller. Injects PlayFlowMIDI.BackgroundHook into the
    /// game process, then feeds synthetic raw keyboard events through shared memory while the
    /// game stays unfocused. This loads code into the game process and can be detected by
    /// anti-cheat; it is never enabled unless the user selects the experimental mode.
    /// </summary>
    public sealed class RawInputInjector : IDisposable
    {
        private const int HeaderSize = 64;
        private const int SlotCount = 1024;
        private const int SlotSize = 8;
        private const int TotalSize = HeaderSize + SlotCount * SlotSize;
        private const int OffEnabled = 0;
        private const int OffStop = 12;
        private const int OffMagic = 16;
        private const int OffReady = 20;
        private const int OffHeartbeat = 24;
        private const int OffDebug = 28;
        private const int OffHwnd = 32;
        private const int OffSequence = 40;
        private const int OffPayloadVersion = 48;
        private const int ProtocolVersion = 2;
        private const uint Magic = 0x50464D32;
        private const ulong SentinelBase = 0x50464D0000000000;
        private const uint WM_INPUT = 0x00FF;
        private const uint WM_ACTIVATEAPP = 0x001C;
        private const uint WM_ACTIVATE = 0x0006;
        private const uint WM_SETFOCUS = 0x0007;

        private readonly object _gate = new object();
        private static readonly bool DebugEnabled =
            string.Equals(Environment.GetEnvironmentVariable("PLAYFLOWMIDI_DEBUG"), "1", StringComparison.Ordinal);
        private MemoryMappedFile? _mmf;
        private MemoryMappedViewAccessor? _view;
        private EventWaitHandle? _readyEvent;
        private Process? _helperProcess;
        private System.Threading.Timer? _heartbeatTimer;
        private System.Threading.Timer? _disableTimer;
        private int _disableGeneration;
        private int _pid;
        private IntPtr _hwnd;
        private bool _active;
        private long _sequence;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        public bool IsActive
        {
            get { lock (_gate) return _active; }
        }

        public bool IsActiveFor(int pid, IntPtr hwnd)
        {
            lock (_gate) return _active && _pid == pid && _hwnd == hwnd;
        }

        /// <summary>
        /// Injects the hook into the game process. Blocking; call it from a worker thread.
        /// The hook is kept installed per game process and reused by later app sessions,
        /// because re-injecting after EasyHook unloads a payload is unreliable.
        /// </summary>
        public bool Start(int pid, IntPtr hwnd)
        {
            lock (_gate)
            {
                if (_active && _pid == pid && _hwnd == hwnd)
                {
                    EnableLocked(true);
                    return true;
                }

                StopInternalLocked(signalStop: false);

                string mmfName = $"Local\\PlayFlowMIDI_RawHook2_{pid}";
                string readyName = $"Local\\PlayFlowMIDI_RawHook2Ready_{pid}";
                string baseDir = AppContext.BaseDirectory;
                string injectorPath = Path.Combine(baseDir, "PlayFlowMIDI.Injector.exe");
                string payloadPath = Path.Combine(baseDir, "PlayFlowMIDI.BackgroundHook.dll");

                // Reuse a hook that is already installed in this game process.
                try
                {
                    _mmf = MemoryMappedFile.OpenExisting(mmfName, MemoryMappedFileRights.ReadWrite);
                    _view = _mmf.CreateViewAccessor(0, TotalSize, MemoryMappedFileAccess.ReadWrite);
                    if (_view.ReadInt32(OffMagic) == Magic &&
                        _view.ReadInt32(OffReady) != 0 &&
                        _view.ReadInt32(OffPayloadVersion) == ProtocolVersion)
                    {
                        _readyEvent = new EventWaitHandle(false, EventResetMode.ManualReset, readyName);
                        _pid = pid;
                        _hwnd = hwnd;
                        _active = true;
                        _sequence = _view.ReadInt64(OffSequence);
                        _view.Write(OffHwnd, hwnd.ToInt64());
                        _view.Write(OffDebug, DebugEnabled ? 1 : 0);
                        EnableLocked(true);
                        _heartbeatTimer = new System.Threading.Timer(_ => WriteHeartbeat(), null, 0, 1000);
                        Log($"reused existing injection pid={pid} hwnd=0x{hwnd.ToInt64():X}");
                        return true;
                    }

                    _view.Write(OffStop, 1);
                    _view.Dispose();
                    _view = null;
                    _mmf.Dispose();
                    _mmf = null;
                }
                catch (FileNotFoundException)
                {
                }
                catch (Exception ex)
                {
                    Log("reuse check failed: " + ex.Message);
                    _view?.Dispose();
                    _view = null;
                    _mmf?.Dispose();
                    _mmf = null;
                }

                if (!File.Exists(injectorPath) || !File.Exists(payloadPath))
                {
                    Log($"missing injection files: injector={injectorPath} payload={payloadPath}");
                    return false;
                }

                try
                {
                    _mmf = MemoryMappedFile.CreateOrOpen(mmfName, TotalSize, MemoryMappedFileAccess.ReadWrite);
                    _view = _mmf.CreateViewAccessor(0, TotalSize, MemoryMappedFileAccess.ReadWrite);
                    _view.Write(OffEnabled, 0);
                    _view.Write(OffStop, 0);
                    _view.Write(OffMagic, Magic);
                    _view.Write(OffReady, 0);
                    _view.Write(OffHeartbeat, Environment.TickCount);
                    _view.Write(OffHwnd, hwnd.ToInt64());
                    _view.Write(OffDebug, DebugEnabled ? 1 : 0);
                    _view.Write(OffSequence, 0L);
                    _view.Write(OffPayloadVersion, 0);
                    _sequence = 0;

                    _readyEvent = new EventWaitHandle(false, EventResetMode.ManualReset, readyName);
                    _readyEvent.Reset();

                    var psi = new ProcessStartInfo
                    {
                        FileName = injectorPath,
                        Arguments = $"{pid} 0x{hwnd.ToInt64():X} \"{mmfName}\" \"{payloadPath}\" \"{readyName}\" {Environment.ProcessId}",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WorkingDirectory = baseDir,
                    };
                    _helperProcess = Process.Start(psi);

                    bool ready = _readyEvent.WaitOne(8000);
                    if (!ready || _helperProcess == null || _helperProcess.HasExited)
                    {
                        Log("injection did not become ready");
                        StopInternalLocked(signalStop: true);
                        return false;
                    }

                    int payloadVersion = _view.ReadInt32(OffPayloadVersion);
                    if (payloadVersion != ProtocolVersion)
                    {
                        Log($"payload protocol mismatch: got {payloadVersion}, expected {ProtocolVersion}. " +
                            "Use a matching build and restart the game.");
                        StopInternalLocked(signalStop: true);
                        return false;
                    }

                    _pid = pid;
                    _hwnd = hwnd;
                    _active = true;
                    EnableLocked(true);
                    _heartbeatTimer = new System.Threading.Timer(_ => WriteHeartbeat(), null, 0, 1000);
                    Log($"injection ready pid={pid} hwnd=0x{hwnd.ToInt64():X}");
                    return true;
                }
                catch (Exception ex)
                {
                    Log("injection start failed: " + ex);
                    StopInternalLocked(signalStop: true);
                    return false;
                }
            }
        }

        public void SetEnabled(bool enabled)
        {
            lock (_gate) EnableLocked(enabled);
        }

        /// <summary>Lets the final key-up events drain before removing the focus spoof.</summary>
        public void DisableAfterDelay(int milliseconds = 600)
        {
            lock (_gate)
            {
                int generation = ++_disableGeneration;
                _disableTimer?.Dispose();
                _disableTimer = new System.Threading.Timer(_ =>
                {
                    lock (_gate)
                    {
                        if (generation == _disableGeneration) EnableLocked(false);
                    }
                }, null, Math.Max(0, milliseconds), Timeout.Infinite);
            }
        }

        /// <summary>Posts the activation messages the game expects before accepting input.</summary>
        public void ActivateWindow()
        {
            lock (_gate)
            {
                if (!_active || _hwnd == IntPtr.Zero) return;
                // Never send fake activation messages to a game that is genuinely focused.
                if (GetForegroundWindow() == _hwnd) return;
                PostMessage(_hwnd, WM_ACTIVATEAPP, (IntPtr)1, IntPtr.Zero);
                PostMessage(_hwnd, WM_ACTIVATE, (IntPtr)1, IntPtr.Zero);
                PostMessage(_hwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
            }
        }

        public void SendKeyCombo(string keyCombo, bool? isDown)
        {
            lock (_gate)
            {
                if (!_active || _view == null || _hwnd == IntPtr.Zero) return;
                if (!InputSimulator.TryParseKeyCombo(keyCombo, out Forms.Keys mainKey, out List<Forms.Keys> modifiers))
                    return;

                EnableLocked(true);

                if (isDown == true)
                {
                    foreach (var mod in modifiers) EnqueueKeyLocked(mod, false);
                    if (mainKey != Forms.Keys.None) EnqueueKeyLocked(mainKey, false);
                }
                else if (isDown == false)
                {
                    if (mainKey != Forms.Keys.None) EnqueueKeyLocked(mainKey, true);
                    modifiers.Reverse();
                    foreach (var mod in modifiers) EnqueueKeyLocked(mod, true);
                }
                else
                {
                    foreach (var mod in modifiers) EnqueueKeyLocked(mod, false);
                    if (mainKey != Forms.Keys.None) EnqueueKeyLocked(mainKey, false);
                    Thread.Sleep(2);
                    if (mainKey != Forms.Keys.None) EnqueueKeyLocked(mainKey, true);
                    modifiers.Reverse();
                    foreach (var mod in modifiers) EnqueueKeyLocked(mod, true);
                }
            }
        }

        /// <summary>
        /// Enqueues one raw key transition without implicit modifiers, for chord batching.
        /// </summary>
        public void SendSingleKey(Forms.Keys key, bool up)
        {
            lock (_gate)
            {
                if (!_active || _view == null || _hwnd == IntPtr.Zero) return;
                EnableLocked(true);
                EnqueueKeyLocked(key, up);
            }
        }

        public void ReleaseModifiers()
        {
            lock (_gate)
            {
                if (!_active || _view == null) return;
                EnqueueRawLocked(0x10, false, false); // VK_SHIFT
                EnqueueRawLocked(0x10, true, true);   // right shift
                EnqueueRawLocked(0x11, false, false); // VK_CONTROL
                EnqueueRawLocked(0x11, true, true);   // right ctrl
                EnqueueRawLocked(0x12, false, false); // VK_MENU
                EnqueueRawLocked(0x12, true, true);   // right alt
                EnqueueRawLocked(0x5B, false, false); // left win
                EnqueueRawLocked(0x5C, false, true);  // right win
            }
        }

        public void Stop()
        {
            bool wasActive;
            lock (_gate)
            {
                wasActive = _active;
                if (!wasActive)
                {
                    StopInternalLocked(signalStop: false);
                    return;
                }

                // Keep the hook enabled briefly so final key-up events are consumed.
                _disableGeneration++;
                _disableTimer?.Dispose();
                _disableTimer = null;
            }

            Thread.Sleep(700);

            lock (_gate)
            {
                EnableLocked(false);
                StopInternalLocked(signalStop: false);
            }
        }

        public void Dispose() => Stop();

        private void EnableLocked(bool enabled)
        {
            if (_view == null) return;
            try
            {
                if (enabled)
                {
                    _disableGeneration++;
                    _disableTimer?.Dispose();
                    _disableTimer = null;
                    _view.Write(OffHeartbeat, Environment.TickCount);
                }
                _view.Write(OffEnabled, enabled ? 1 : 0);
            }
            catch (Exception ex)
            {
                Log("enabled write failed: " + ex.Message);
            }
        }

        private void WriteHeartbeat()
        {
            lock (_gate)
            {
                if (!_active || _view == null) return;
                try { _view.Write(OffHeartbeat, Environment.TickCount); }
                catch { }
            }
        }

        private void EnqueueKeyLocked(Forms.Keys key, bool up)
        {
            GetRawKey(key, out byte vk, out bool extended);
            EnqueueRawLocked(vk, up, extended);
        }

        private void EnqueueRawLocked(byte vk, bool up, bool extended)
        {
            if (_view == null || _hwnd == IntPtr.Zero) return;
            try
            {
                long sequence = ++_sequence;
                long slot = HeaderSize + (sequence % SlotCount) * SlotSize;
                byte flags = (byte)((up ? 1 : 0) | (extended ? 2 : 0));

                _view.Write(slot, vk);
                _view.Write(slot + 1, flags);
                _view.Write(slot + 2, (ushort)0);
                _view.Write(slot + 4, (uint)sequence);
                Thread.MemoryBarrier();
                _view.Write(OffSequence, sequence);

                // The sentinel identifies this exact event. The injected hook serves it
                // once and ignores the extra GetRawInputData calls some games make.
                IntPtr sentinel = unchecked((IntPtr)(long)(SentinelBase | (uint)sequence));
                PostMessage(_hwnd, WM_INPUT, IntPtr.Zero, sentinel);
            }
            catch (Exception ex)
            {
                Log("enqueue failed: " + ex.Message);
            }
        }

        private static void GetRawKey(Forms.Keys key, out byte vk, out bool extended)
        {
            extended = false;
            switch (key)
            {
                case Forms.Keys.LShiftKey: vk = 0x10; return;
                case Forms.Keys.RShiftKey: vk = 0x10; extended = true; return;
                case Forms.Keys.LControlKey: vk = 0x11; return;
                case Forms.Keys.RControlKey: vk = 0x11; extended = true; return;
                case Forms.Keys.LMenu: vk = 0x12; return;
                case Forms.Keys.RMenu: vk = 0x12; extended = true; return;
                case Forms.Keys.LWin: vk = 0x5B; return;
                case Forms.Keys.RWin: vk = 0x5C; extended = true; return;
                default:
                    vk = (byte)key;
                    extended = IsExtendedKey(key);
                    return;
            }
        }

        private static bool IsExtendedKey(Forms.Keys key)
        {
            return key == Forms.Keys.Insert || key == Forms.Keys.Delete || key == Forms.Keys.Home ||
                   key == Forms.Keys.End || key == Forms.Keys.PageUp || key == Forms.Keys.PageDown ||
                   key == Forms.Keys.Up || key == Forms.Keys.Down || key == Forms.Keys.Left ||
                   key == Forms.Keys.Right || key == Forms.Keys.NumLock || key == Forms.Keys.Divide ||
                   key == Forms.Keys.PrintScreen || key == Forms.Keys.Apps;
        }

        private void StopInternalLocked(bool signalStop)
        {
            try
            {
                if (_view != null)
                {
                    _view.Write(OffEnabled, 0);
                    if (signalStop) _view.Write(OffStop, 1);
                }
            }
            catch { }

            _heartbeatTimer?.Dispose();
            _heartbeatTimer = null;
            _disableTimer?.Dispose();
            _disableTimer = null;

            try { _readyEvent?.Dispose(); } catch { }
            _readyEvent = null;

            if (_helperProcess != null)
            {
                try
                {
                    if (!_helperProcess.HasExited)
                    {
                        _helperProcess.WaitForExit(800);
                        if (!_helperProcess.HasExited) _helperProcess.Kill();
                    }
                }
                catch { }
                _helperProcess.Dispose();
                _helperProcess = null;
            }

            _view?.Dispose();
            _view = null;
            _mmf?.Dispose();
            _mmf = null;
            _active = false;
            _pid = 0;
            _hwnd = IntPtr.Zero;
        }

        private static void Log(string message)
        {
            try
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PlayFlowMIDI");
                Directory.CreateDirectory(folder);
                File.AppendAllText(Path.Combine(folder, "background-injector.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
