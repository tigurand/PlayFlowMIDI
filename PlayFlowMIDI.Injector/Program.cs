using System;
using System.Diagnostics;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Threading;
using EasyHook;

namespace PlayFlowMIDI.Injector
{
    /// <summary>
    /// Small helper process that injects PlayFlowMIDI.BackgroundHook into the game process.
    /// It only performs the injection requested by the opt-in Background Injection mode and
    /// stays alive until the controller signals stop.
    /// </summary>
    internal static class Program
    {
        private const int OffStop = 12;

        private static void Main(string[] args)
        {
            try
            {
                if (args.Length < 5)
                {
                    Log("usage: pid hwndHex mmfName payloadPath readyEventName");
                    Environment.ExitCode = 2;
                    return;
                }

                int pid = int.Parse(args[0]);
                string hwndHex = args[1];
                string mmfName = args[2];
                string payloadPath = Path.GetFullPath(args[3]);
                string readyEventName = args[4];
                int parentPid = args.Length > 5 ? int.Parse(args[5]) : 0;

                Log($"injecting pid={pid} hwnd={hwndHex} mmf={mmfName} payload={payloadPath}");
                RemoteHooking.Inject(pid, InjectionOptions.DoNotRequireStrongName,
                    payloadPath, payloadPath, hwndHex, mmfName, readyEventName);
                Log("injection call returned");

                using (var mmf = MemoryMappedFile.OpenExisting(mmfName, MemoryMappedFileRights.ReadWrite))
                using (var view = mmf.CreateViewAccessor())
                {
                    while (true)
                    {
                        Thread.Sleep(500);
                        if (view.ReadInt32(OffStop) != 0) break;

                        if (parentPid != 0)
                        {
                            try { using (Process.GetProcessById(parentPid)) { } }
                            catch { break; }
                        }
                    }
                }
                Log("stop signaled; exiting");
            }
            catch (Exception ex)
            {
                Log("FATAL " + ex);
                Environment.ExitCode = 1;
            }
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
