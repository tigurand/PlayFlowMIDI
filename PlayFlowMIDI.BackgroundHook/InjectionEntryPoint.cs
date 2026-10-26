using System;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Runtime.InteropServices;
using System.Threading;
using EasyHook;

namespace PlayFlowMIDI.BackgroundHook
{
    /// <summary>
    /// Opt-in in-process hook used only by the experimental Background Injection mode.
    /// It makes the game process believe it is focused and returns synthetic raw keyboard
    /// events for WM_INPUT messages posted by PlayFlowMIDI. No anti-cheat bypass or hiding
    /// is performed; the injection is visible and may be detected.
    /// </summary>
    public class InjectionEntryPoint : IEntryPoint
    {
        private const int HeaderSize = 64;
        private const int SlotCount = 1024;
        private const int SlotSize = 8;
        private const int OffEnabled = 0;
        private const int OffStop = 12;
        private const int OffReady = 20;
        private const int OffHeartbeat = 24;
        private const int OffDebug = 28;
        private const int OffHwnd = 32;
        private const int OffSequence = 40;
        private const int OffPayloadVersion = 48;
        private const int ProtocolVersion = 2;
        private const ulong SentinelBase = 0x50464D0000000000;
        private const ulong SentinelMask = 0xFFFFFFFF00000000;
        private const uint RidInput = 0x10000003;
        private const int RawInputSize = 40;
        private const int HeartbeatTimeoutMs = 5000;

        private static readonly object Gate = new object();

        private static IntPtr _gameHwnd;
        private static string _readyEventName;
        private static MemoryMappedViewAccessor _view;
        private static LocalHook _hookForeground;
        private static LocalHook _hookActive;
        private static LocalHook _hookFocus;
        private static LocalHook _hookRaw;
        private static GetForegroundWindowDelegate _origForeground;
        private static GetActiveWindowDelegate _origActive;
        private static GetFocusDelegate _origFocus;
        private static GetRawInputDataDelegate _origRaw;
        private static int _rawCalls;
        private static int _fgCalls;
        private static int _activeCalls;
        private static int _focusCalls;
        private static int _syntheticCalls;
        private static long _lastServedSequence;

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate IntPtr GetForegroundWindowDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate IntPtr GetActiveWindowDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate IntPtr GetFocusDelegate();

        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        private delegate uint GetRawInputDataDelegate(IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        public InjectionEntryPoint(RemoteHooking.IContext context, string gameHwndHex, string mmfName, string readyEventName)
        {
            _gameHwnd = (IntPtr)Convert.ToInt64(gameHwndHex, 16);
            _readyEventName = readyEventName;
            _view = MemoryMappedFile.OpenExisting(mmfName, MemoryMappedFileRights.ReadWrite).CreateViewAccessor();
            Log("constructed");
        }

        public void Run(RemoteHooking.IContext context, string gameHwndHex, string mmfName, string readyEventName)
        {
            try
            {
                _hookForeground = LocalHook.Create(
                    LocalHook.GetProcAddress("user32.dll", "GetForegroundWindow"),
                    new GetForegroundWindowDelegate(GetForegroundWindowHook), null);
                _origForeground = (GetForegroundWindowDelegate)Marshal.GetDelegateForFunctionPointer(
                    _hookForeground.HookBypassAddress, typeof(GetForegroundWindowDelegate));
                _hookForeground.ThreadACL.SetExclusiveACL(new[] { 0 });

                _hookActive = LocalHook.Create(
                    LocalHook.GetProcAddress("user32.dll", "GetActiveWindow"),
                    new GetActiveWindowDelegate(GetActiveWindowHook), null);
                _origActive = (GetActiveWindowDelegate)Marshal.GetDelegateForFunctionPointer(
                    _hookActive.HookBypassAddress, typeof(GetActiveWindowDelegate));
                _hookActive.ThreadACL.SetExclusiveACL(new[] { 0 });

                _hookFocus = LocalHook.Create(
                    LocalHook.GetProcAddress("user32.dll", "GetFocus"),
                    new GetFocusDelegate(GetFocusHook), null);
                _origFocus = (GetFocusDelegate)Marshal.GetDelegateForFunctionPointer(
                    _hookFocus.HookBypassAddress, typeof(GetFocusDelegate));
                _hookFocus.ThreadACL.SetExclusiveACL(new[] { 0 });

                _hookRaw = LocalHook.Create(
                    LocalHook.GetProcAddress("user32.dll", "GetRawInputData"),
                    new GetRawInputDataDelegate(GetRawInputDataHook), null);
                _origRaw = (GetRawInputDataDelegate)Marshal.GetDelegateForFunctionPointer(
                    _hookRaw.HookBypassAddress, typeof(GetRawInputDataDelegate));
                _hookRaw.ThreadACL.SetExclusiveACL(new[] { 0 });

                WriteHeader(OffPayloadVersion, ProtocolVersion);
                WriteHeader(OffReady, 1);
                SignalReady();
                Log($"hooks installed (protocol {ProtocolVersion})");

                MonitorLoop();
            }
            catch (Exception ex)
            {
                Log("ERROR " + ex);
            }
            finally
            {
                WriteHeader(OffPayloadVersion, 0);
                WriteHeader(OffReady, 0);
                DisposeHooks();
                Log("hooks disposed");
            }
        }

        private static void MonitorLoop()
        {
            int lastHeartbeat = ReadHeader(OffHeartbeat);
            int lastChange = Environment.TickCount;
            while (true)
            {
                Thread.Sleep(200);
                if (ReadHeader(OffStop) != 0) return;

                int heartbeat = ReadHeader(OffHeartbeat);
                int now = Environment.TickCount;
                if (heartbeat != lastHeartbeat)
                {
                    lastHeartbeat = heartbeat;
                    lastChange = now;
                }
                else if (IsEnabled() && unchecked(now - lastChange) > HeartbeatTimeoutMs)
                {
                    // The controller stopped sending heartbeats (crashed or was killed).
                    // Stop spoofing focus so the game returns to normal behavior.
                    WriteHeader(OffEnabled, 0);
                    Log("heartbeat lost; disabled");
                }
            }
        }

        private static bool IsEnabled()
        {
            lock (Gate)
            {
                try { return _view.ReadInt32(OffEnabled) != 0; }
                catch { return false; }
            }
        }

        private static int ReadHeader(int offset)
        {
            lock (Gate)
            {
                try { return _view.ReadInt32(offset); }
                catch { return 0; }
            }
        }

        private static void WriteHeader(int offset, int value)
        {
            lock (Gate)
            {
                try { _view.Write(offset, value); }
                catch { }
            }
        }

        private static bool IsVerbose()
        {
            lock (Gate)
            {
                try { return _view.ReadInt32(OffDebug) != 0; }
                catch { return false; }
            }
        }

        private static IntPtr GameHwnd
        {
            get
            {
                lock (Gate)
                {
                    try
                    {
                        long value = _view.ReadInt64(OffHwnd);
                        return value != 0 ? (IntPtr)value : _gameHwnd;
                    }
                    catch
                    {
                        return _gameHwnd;
                    }
                }
            }
        }

        private static IntPtr GetForegroundWindowHook()
        {
            int n = Interlocked.Increment(ref _fgCalls);
            bool enabled = IsEnabled();
            IntPtr real = _origForeground();
            bool background = real != GameHwnd;
            if (IsVerbose() && (n <= 3 || n % 1000 == 0))
                Log($"GetForegroundWindow #{n} enabled={enabled} background={background}");
            return enabled && background ? GameHwnd : real;
        }

        private static IntPtr GetActiveWindowHook()
        {
            int n = Interlocked.Increment(ref _activeCalls);
            bool enabled = IsEnabled();
            bool background = _origForeground() != GameHwnd;
            if (IsVerbose() && (n <= 3 || n % 1000 == 0))
                Log($"GetActiveWindow #{n} enabled={enabled} background={background}");
            if (enabled && background) return GameHwnd;
            return _origActive();
        }

        private static IntPtr GetFocusHook()
        {
            int n = Interlocked.Increment(ref _focusCalls);
            bool enabled = IsEnabled();
            bool background = _origForeground() != GameHwnd;
            if (IsVerbose() && (n <= 3 || n % 1000 == 0))
                Log($"GetFocus #{n} enabled={enabled} background={background}");
            if (enabled && background) return GameHwnd;
            return _origFocus();
        }

        private static uint GetRawInputDataHook(IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader)
        {
            int n = Interlocked.Increment(ref _rawCalls);
            bool enabled = IsEnabled();
            bool verbose = IsVerbose();

            ulong rawHandle = unchecked((ulong)hRawInput.ToInt64());
            bool sentinel = (rawHandle & SentinelMask) == SentinelBase;
            uint sequence = sentinel ? (uint)(rawHandle & 0xFFFFFFFF) : 0;

            if (verbose)
                Log($"GetRawInputData #{n} handle=0x{hRawInput.ToInt64():X} sentinel={sentinel} seq={sequence} cmd=0x{uiCommand:X} pData={(pData == IntPtr.Zero ? "null" : "ptr")} enabled={enabled}");

            if (sentinel && uiCommand == RidInput && enabled)
            {
                if (pData == IntPtr.Zero)
                {
                    // Size query for the tagged message. Do not consume the event yet.
                    if (IsSlotValid(sequence))
                    {
                        pcbSize = RawInputSize;
                        return 0;
                    }
                    return 0;
                }

                if (TryServe(sequence, out byte vk, out byte flags))
                {
                    Interlocked.Increment(ref _syntheticCalls);
                    if (verbose) Log($"synthetic seq={sequence} vk=0x{vk:X2} flags={flags}");
                    pcbSize = RawInputSize;
                    WriteRawInput(pData, vk, flags);
                    return RawInputSize;
                }

                // Duplicate GetRawInputData call for the same message, or an overwritten/stale event. Never consume a different queued event here.
                return 0;
            }

            return _origRaw(hRawInput, uiCommand, pData, ref pcbSize, cbSizeHeader);
        }

        private static bool IsSlotValid(uint sequence)
        {
            lock (Gate)
            {
                try
                {
                    if (sequence <= _lastServedSequence) return false;
                    long slot = HeaderSize + (long)(sequence % SlotCount) * SlotSize;
                    return _view.ReadUInt32(slot + 4) == sequence;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static bool TryServe(uint sequence, out byte vk, out byte flags)
        {
            vk = 0;
            flags = 0;
            lock (Gate)
            {
                try
                {
                    if (sequence <= _lastServedSequence) return false;

                    long slot = HeaderSize + (long)(sequence % SlotCount) * SlotSize;
                    if (_view.ReadUInt32(slot + 4) != sequence) return false;

                    vk = _view.ReadByte(slot);
                    flags = _view.ReadByte(slot + 1);
                    _lastServedSequence = sequence;
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static void WriteRawInput(IntPtr pData, byte vk, byte flags)
        {
            bool up = (flags & 1) != 0;
            bool extended = (flags & 2) != 0;
            ushort scan = (ushort)(MapVirtualKey(vk, 0) & 0xFFFF);

            Marshal.WriteInt32(pData, 0, 1);              // RIM_TYPEKEYBOARD
            Marshal.WriteInt32(pData, 4, RawInputSize);   // dwSize
            Marshal.WriteIntPtr(pData, 8, IntPtr.Zero);   // hDevice
            Marshal.WriteIntPtr(pData, 16, IntPtr.Zero);  // wParam
            Marshal.WriteInt16(pData, 24, (short)scan);
            Marshal.WriteInt16(pData, 26, (short)((up ? 1 : 0) | (extended ? 2 : 0)));
            Marshal.WriteInt16(pData, 28, 0);
            Marshal.WriteInt16(pData, 30, (short)vk);
            Marshal.WriteInt32(pData, 32, up ? 0x0101 : 0x0100);
            Marshal.WriteInt32(pData, 36, 0);
        }

        private static void SignalReady()
        {
            try
            {
                using (var ready = EventWaitHandle.OpenExisting(_readyEventName))
                {
                    ready.Set();
                }
            }
            catch (Exception ex)
            {
                Log("ready event error: " + ex.Message);
            }
        }

        private static void DisposeHooks()
        {
            try { _hookRaw?.Dispose(); } catch { }
            try { _hookFocus?.Dispose(); } catch { }
            try { _hookActive?.Dispose(); } catch { }
            try { _hookForeground?.Dispose(); } catch { }
        }

        private static void Log(string message)
        {
            try
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PlayFlowMIDI");
                Directory.CreateDirectory(folder);
                File.AppendAllText(Path.Combine(folder, "background-hook.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
