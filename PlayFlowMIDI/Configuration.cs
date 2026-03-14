using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace PlayFlowMIDI
{
    public class AppConfig
    {
        public MainConfig Main { get; set; } = new();
        public SettingsConfig Settings { get; set; } = new();
        public List<ProfileConfig> Profiles { get; set; } = new();
        public string LastSelectedProfile { get; set; } = "Heartopia";
        public WindowConfig Window { get; set; } = new();

        public static string GetConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string folder = Path.Combine(appData, "PlayFlowMIDI");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            return Path.Combine(folder, "config.json");
        }

        public void Save()
        {
            try
            {
                string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(GetConfigPath(), json);
            }
            catch { }
        }

        public static AppConfig Load()
        {
            try
            {
                string path = GetConfigPath();
                if (File.Exists(path))
                {
                    string json = File.ReadAllText(path);
                    return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                }
            }
            catch { }
            return new AppConfig();
        }
    }

    public class MainConfig
    {
        public string FolderPath { get; set; } = string.Empty;
        public int SortByIndex { get; set; } = 0;
        public bool SortAscending { get; set; } = true;
        public string PlaybackMode { get; set; } = "PlayOnce"; // PlayOnce, RepeatSong, RepeatPlaylist, RandomPlaylist
        public List<MidiSongConfig> Playlist { get; set; } = new();
    }

    public class MidiSongConfig
    {
        public string Title { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string BPM { get; set; } = string.Empty;
        public string TimeAdded { get; set; } = string.Empty;
        public string Duration { get; set; } = string.Empty;
        public int CustomOrder { get; set; } = 0;
        public List<int> DisabledTrackIndices { get; set; } = new();
        public bool TrackStatesManuallySet { get; set; } = false;
    }

    public class SettingsConfig
    {
        public string ShortcutPlayPause { get; set; } = "Ctrl+F1";
        public string ShortcutStop { get; set; } = "Ctrl+F2";
        public string ShortcutPrev { get; set; } = "Ctrl+F3";
        public string ShortcutNext { get; set; } = "Ctrl+F4";
        public string ShortcutPreview { get; set; } = "Ctrl+F5";
        public string ShortcutToggleMode { get; set; } = "Ctrl+F6";
        public string OctaveHandling { get; set; } = "Default"; // Default, Wrap, Clamp
        public bool TrimLeadingSilence { get; set; } = false;
        public bool HoldNotes { get; set; } = true;
        public int ReleaseDelay { get; set; } = 20;
        public bool DisablePercussion { get; set; } = true;
        public bool AutoAdapt15Keys { get; set; } = false;
        public bool AutoUpdateEnabled { get; set; } = false;
        public bool AlwaysOnTop { get; set; } = false;
        public string MidiInputDevice { get; set; } = string.Empty;
        public double Speed { get; set; } = 1.0;
        public int Pitch { get; set; } = 0;
    }

    public class ProfileConfig
    {
        public string Name { get; set; } = string.Empty;
        public string ExeName { get; set; } = string.Empty;
        public bool SwitchToWindow { get; set; }
        public bool AutoPause { get; set; }
        public string SelectedMode { get; set; } = string.Empty; // 37, 36, 15
        public bool IsDefault { get; set; } = false;
        public Dictionary<string, List<string>> KeyMappings { get; set; } = new();
    }

    public class WindowConfig
    {
        public double Width { get; set; } = 1250;
        public double Height { get; set; } = 760;
        public double Left { get; set; } = -1;
        public double Top { get; set; } = -1;
        public WindowState State { get; set; } = WindowState.Normal;
    }
}
