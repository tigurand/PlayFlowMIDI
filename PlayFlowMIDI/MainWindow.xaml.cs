using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Text;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Win32;
using Melanchall.DryWetMidi.Core;
using Melanchall.DryWetMidi.Interaction;
using Melanchall.DryWetMidi.Multimedia;
using Melanchall.DryWetMidi.Common;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Windows.Interop;
using Forms = System.Windows.Forms;

namespace PlayFlowMIDI
{
    public class MidiSong
    {
        public string Title { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string BPM { get; set; } = "N/A";
        public string TimeAdded { get; set; } = string.Empty;
        public string Duration { get; set; } = "0:00";
        public int CustomOrder { get; set; } = 0;
        public List<int> DisabledTrackIndices { get; set; } = new();
        public bool TrackStatesManuallySet { get; set; } = false;
    }

    public partial class MainWindow : Window
    {
        private AppConfig _config = null!;
        private bool _isLoading = true;
        private System.Windows.Point _dragStartPoint;
        private List<MidiSong>? _draggedSongs;
        private List<MidiSong> _masterPlaylist = new();
        private Playback? _playback;
        private MidiFile? _currentMidiFile;
        private IOutputDevice? _outputDevice;
        private IInputDevice? _midiInputDevice;
        private bool _isPreviewMode = false;
        private double _currentSpeed = 1.0;
        private int _currentTranspose = 0;
        private bool _isPlaying = false;
        private bool _isUserDraggingSlider = false;
        private bool _isAutoPaused = false;
        private DispatcherTimer? _uiTimer;
        private MidiSong? _currentPlayingSong;
        private List<int> _shuffleIndices = new();
        private List<int> _previousShuffle = new();
        private int _shufflePosition = -1;

        private const string RepoOwner = "tigurand";
        private const string RepoName = "PlayFlowMIDI";
        private const string AssetPrefix = "PlayFlowMIDI-v";

        public MainWindow()
        {
            InitializeComponent();

            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Title = $"PlayFlowMIDI - {version}";

            _config = AppConfig.Load();
            
            StartInputProcessor();

            if (!_config.Profiles.Any(p => p.Name == "Heartopia"))
            {
                _config.Profiles.Add(new ProfileConfig { Name = "Heartopia", ExeName = "xdt.exe", SwitchToWindow = true, AutoPause = true, SelectedMode = "37", IsDefault = true });
            }
            else { _config.Profiles.First(p => p.Name == "Heartopia").IsDefault = true; }

            if (!_config.Profiles.Any(p => p.Name == "Where Winds Meet"))
            {
                _config.Profiles.Add(new ProfileConfig { Name = "Where Winds Meet", ExeName = "wwm.exe", SwitchToWindow = false, AutoPause = false, SelectedMode = "36", IsDefault = true });
            }
            else { _config.Profiles.First(p => p.Name == "Where Winds Meet").IsDefault = true; }

            ApplyConfig();
            
            _isLoading = false;
            SaveConfig();
            _isLoading = true;

            _uiTimer = new DispatcherTimer();
            _uiTimer.Interval = TimeSpan.FromMilliseconds(100);
            _uiTimer.Tick += UiTimer_Tick;
            _uiTimer.Start();

            _isLoading = false;
            SortPlaylist();

            string lastPath = _config.Main.LastPlayedSongPath;

            ReloadSongs();

            if (!string.IsNullOrEmpty(lastPath))
            {
                var lastSong = _masterPlaylist.FirstOrDefault(s => s.FilePath.Equals(lastPath, StringComparison.OrdinalIgnoreCase));
                if (lastSong != null)
                {
                    PlaylistListView.SelectedItem = lastSong;
                    PlaylistListView.ScrollIntoView(lastSong);
                    _ = StartPlayback(lastSong, false);
                }
            }

            _ = MaybeAutoCheckUpdates();

            try
            {
                if (OutputDevice.GetAll().Any())
                {
                    _outputDevice = OutputDevice.GetByIndex(0);
                }
            }
            catch { }

            RefreshMidiInputDevices();

            this.Closed += (s, e) => 
            {
                SaveConfig();
                _outputDevice?.Dispose();
                _midiInputDevice?.Dispose();
            };
        }

        private void UiTimer_Tick(object? sender, EventArgs e)
        {
            if (_playback != null && !_isUserDraggingSlider)
            {
                // Auto Pause Logic
                string profileName = _config.LastSelectedProfile;
                var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
                if (!_isPreviewMode && profile != null && profile.AutoPause && _targetHwnd != IntPtr.Zero)
                {
                    IntPtr foregroundHwnd = GetForegroundWindow();
                    bool isForeground = (foregroundHwnd == _targetHwnd);

                    if (!isForeground && _playback.IsRunning && !_isAutoPaused)
                    {
                        _playback.Stop();
                        ReleaseActiveKeys();
                        _isAutoPaused = true;
                        _isPlaying = false;
                        PlayPauseButton.Content = "▶";
                    }
                    else if (isForeground && !_playback.IsRunning && _isAutoPaused)
                    {
                        _playback.Start();
                        _isAutoPaused = false;
                        _isPlaying = true;
                        PlayPauseButton.Content = "⏸";
                    }
                }

                if (_playback.IsRunning)
                {
                    var currentTime = _playback.GetCurrentTime<MetricTimeSpan>();
                    var duration = _playback.GetDuration<MetricTimeSpan>();

                    PlaybackSlider.Maximum = duration.TotalMicroseconds;
                    PlaybackSlider.Value = currentTime.TotalMicroseconds;

                    PlaybackTimeText.Text = $"{(int)currentTime.Minutes}:{currentTime.Seconds:D2} / {(int)duration.Minutes}:{duration.Seconds:D2}";
                }
            }
        }

        private void PlayPauseButton_Click(object sender, RoutedEventArgs e)
        {
            TogglePlayPause(false);
        }

        private void PreviewButton_Click(object sender, RoutedEventArgs e)
        {
            TogglePlayPause(true);
        }

        private async void TogglePlayPause(bool isPreview)
        {
            if (_playback == null)
            {
                var selectedSong = PlaylistListView.SelectedItem as MidiSong;
                if (selectedSong != null)
                {
                    _isPreviewMode = isPreview;
                    await StartPlayback(selectedSong);
                }
            }
            else
            {
                if (_isPreviewMode != isPreview)
                {
                    var currentSong = _currentPlayingSong;
                    StopPlayback(false);
                    if (currentSong != null)
                    {
                        _isPreviewMode = isPreview;
                        await StartPlayback(currentSong);
                    }
                    return;
                }

                if (_playback.IsRunning)
                {
                    _playback.Stop();
                    ReleaseActiveKeys();
                    _isPlaying = false;
                    _isAutoPaused = false;
                    PlayPauseButton.Content = "▶";
                    var previewBtn = (System.Windows.Controls.Button)FindName("PreviewButton");
                    if (previewBtn != null) previewBtn.Content = "🎧";
                }
                else
                {
                    if (!_isPreviewMode) await SwitchToGameWindowAsync();
                    _playback.Start();
                    _isPlaying = true;
                    if (_isPreviewMode)
                    {
                        var previewBtn = (System.Windows.Controls.Button)FindName("PreviewButton");
                        if (previewBtn != null) previewBtn.Content = "⏸";
                    }
                    else PlayPauseButton.Content = "⏸";
                }
            }
        }

        private Dictionary<int, int> _pitchActiveCount = new();
        private Dictionary<int, string> _pitchToKey = new();
        private Dictionary<int, System.Threading.CancellationTokenSource> _noteReleaseTokens = new();
        private Dictionary<int, long> _lastNoteOnTime = new();

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

        private delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private IntPtr _targetHwnd = IntPtr.Zero;
        private System.Collections.Concurrent.BlockingCollection<(string Key, IntPtr Target, bool? IsDown)> _inputQueue = new();

        private void StartInputProcessor()
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                foreach (var item in _inputQueue.GetConsumingEnumerable())
                {
                    if (item.IsDown == true) InputSimulator.SimulateKeyDown(item.Key, item.Target);
                    else if (item.IsDown == false) InputSimulator.SimulateKeyUp(item.Key, item.Target);
                    else InputSimulator.SimulateKeyPress(item.Key, item.Target);
                    
                    // Small delay between inputs to ensure game registers them correctly
                    System.Threading.Thread.Sleep(1);
                }
            });
        }

        private void SwitchToGameWindow(bool forceForeground = false)
        {
            _targetHwnd = IntPtr.Zero;
            string profileName = _config.LastSelectedProfile;
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile != null && !string.IsNullOrEmpty(profile.ExeName))
            {
                var processName = Path.GetFileNameWithoutExtension(profile.ExeName);
                var processes = Process.GetProcessesByName(processName);
                
                foreach (var process in processes)
                {
                    _targetHwnd = process.MainWindowHandle;
                    
                    if (_targetHwnd == IntPtr.Zero)
                    {
                        foreach (ProcessThread thread in process.Threads)
                        {
                            EnumThreadWindows(thread.Id, (hwnd, lParam) =>
                            {
                                if (IsWindowVisible(hwnd))
                                {
                                    _targetHwnd = hwnd;
                                    return false;
                                }
                                return true;
                            }, IntPtr.Zero);
                            if (_targetHwnd != IntPtr.Zero) break;
                        }
                    }

                    if (_targetHwnd != IntPtr.Zero)
                    {
                        if (profile.SwitchToWindow && forceForeground)
                        {
                            InputSimulator.SetForegroundWindow(_targetHwnd);
                        }
                        break;
                    }
                }
            }
        }

        private async System.Threading.Tasks.Task SwitchToGameWindowAsync()
        {
            SwitchToGameWindow(true);

            string profileName = _config.LastSelectedProfile;
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile != null && profile.SwitchToWindow && _targetHwnd != IntPtr.Zero)
            {
                // Wait up to 1 second for the window to become active
                var stopwatch = Stopwatch.StartNew();
                while (stopwatch.ElapsedMilliseconds < 1000)
                {
                    if (GetForegroundWindow() == _targetHwnd)
                        break;
                    await System.Threading.Tasks.Task.Delay(100);
                }
            }
        }

        private async System.Threading.Tasks.Task StartPlayback(MidiSong song, bool start = true)
        {
            StopPlayback(false);

            if (!File.Exists(song.FilePath)) return;

            try
            {
                _currentPlayingSong = song;
                
                // Read MIDI and resolve overlapping notes of same pitch
                var rawMidiFile = MidiFile.Read(song.FilePath);

                if (_config.Settings.TrimLeadingSilence)
                {
                    var firstNoteTime = rawMidiFile.GetNotes().FirstOrDefault()?.Time ?? 0;
                    if (firstNoteTime > 0)
                    {
                        foreach (var trackChunk in rawMidiFile.GetTrackChunks())
                        {
                            using (var manager = trackChunk.ManageTimedEvents())
                            {
                                foreach (var timedEvent in manager.Objects)
                                {
                                    timedEvent.Time = Math.Max(0, timedEvent.Time - firstNoteTime);
                                }
                            }
                        }
                    }
                }

                var trackChunks = GetEffectiveTrackChunks(rawMidiFile);

                // Process each track individually to resolve overlapping notes of same pitch and channel
                // This preserves layering across different tracks while fixing overlapping notes within a track
                foreach (var trackChunk in trackChunks)
                {
                    using (var notesManager = trackChunk.ManageNotes())
                    {
                        var notes = notesManager.Objects;
                        var processedNotes = new List<Note>();

                        foreach (var group in notes.GroupBy(n => new { n.NoteNumber, n.Channel }))
                        {
                            var sortedGroup = group.OrderBy(n => n.Time).ToList();
                            for (int i = 0; i < sortedGroup.Count; i++)
                            {
                                var current = sortedGroup[i];
                                if (i < sortedGroup.Count - 1)
                                {
                                    var next = sortedGroup[i + 1];
                                    if (current.Time + current.Length > next.Time)
                                    {
                                        // Truncate current note so it ends when the next one starts
                                        long newLength = next.Time - current.Time;
                                        current.Length = Math.Max(0, newLength);
                                    }
                                }
                                if (current.Length > 0) processedNotes.Add(current);
                            }
                        }

                        notes.Clear();
                        notes.Add(processedNotes);
                    }
                }

                // Automatically adapt to 15 keys if enabled
                string profileName = _config.LastSelectedProfile;
                var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
                if (_config.Settings.AutoAdapt15Keys && profile?.SelectedMode == "15")
                {
                    var allNotes = trackChunks.SelectMany(t => t.GetNotes()).ToList();
                    int root = DetectMajorKey(allNotes);
                    foreach (var trackChunk in trackChunks)
                    {
                        using (var notesManager = trackChunk.ManageNotes())
                        {
                            foreach (var note in notesManager.Objects)
                            {
                                if (note.Channel != 9) // Skip percussion channel
                                {
                                    note.NoteNumber = (Melanchall.DryWetMidi.Common.SevenBitNumber)FoldTo15KeyScale(note.NoteNumber, root);
                                }
                            }
                        }
                    }
                }

                // Create a new MIDI file with the resolved notes preserved in their tracks
                _currentMidiFile = new MidiFile();
                _currentMidiFile.TimeDivision = rawMidiFile.TimeDivision;
                
                // Keep non-track chunks
                foreach (var chunk in rawMidiFile.Chunks)
                {
                    if (!(chunk is TrackChunk))
                    {
                        _currentMidiFile.Chunks.Add(chunk.Clone());
                    }
                }

                // Add track chunks with their meta events and filtered notes
                for (int i = 0; i < trackChunks.Count; i++)
                {
                    var trackChunk = trackChunks[i];
                    
                    // Always include non-note events to preserve tempo map and duration
                    var nonNoteEvents = trackChunk.GetTimedEvents()
                        .Where(te => !(te.Event is NoteOnEvent || te.Event is NoteOffEvent))
                        .ToList();
                    
                    var newTrackChunk = nonNoteEvents.ToTrackChunk();
                    
                    // Only add notes if the track is not disabled
                    if (!IsTrackDisabled(song, trackChunk, i))
                    {
                        using (var notesManager = newTrackChunk.ManageNotes())
                        {
                            notesManager.Objects.Add(trackChunk.GetNotes());
                        }
                    }
                    
                    _currentMidiFile.Chunks.Add(newTrackChunk);
                }
                
                _playback = _currentMidiFile.GetPlayback();
                _playback.Speed = _currentSpeed;
                _playback.InterruptNotesOnStop = true;
                
                _playback.EventPlayed += OnEventPlayed;
                _playback.Finished += OnPlaybackFinished;

                var duration = _playback.GetDuration<MetricTimeSpan>();
                PlaybackSlider.Maximum = duration.TotalMicroseconds;
                PlaybackSlider.Value = 0;
                PlaybackTimeText.Text = $"0:00 / {(int)duration.Minutes}:{duration.Seconds:D2}";
                PlayingSongTitle.Text = song.Title;

                if (start)
                {
                    if (!_isPreviewMode) await SwitchToGameWindowAsync();
                    _playback.Start();
                    _isPlaying = true;
                    _isAutoPaused = false;
                    
                    if (_isPreviewMode)
                    {
                        var previewBtn = (System.Windows.Controls.Button)FindName("PreviewButton");
                        if (previewBtn != null) previewBtn.Content = "⏸";
                    }
                    else
                    {
                        PlayPauseButton.Content = "⏸";
                    }
                }
                else
                {
                    _isPlaying = false;
                    PlayPauseButton.Content = "▶";
                    var previewBtn = (System.Windows.Controls.Button)FindName("PreviewButton");
                    if (previewBtn != null) previewBtn.Content = "🎧";
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error playing MIDI: {ex.Message}");
            }
        }

        private void StopPlayback(bool resetSongInfo = true)
        {
            if (_playback != null)
            {
                _playback.EventPlayed -= OnEventPlayed;
                _playback.Finished -= OnPlaybackFinished;
                _playback.Dispose();
                _playback = null;
            }
            
            ReleaseActiveKeys();

            _isPlaying = false;
            _isAutoPaused = false;
            PlayPauseButton.Content = "▶";
            var previewBtn = (System.Windows.Controls.Button)FindName("PreviewButton");
            if (previewBtn != null) previewBtn.Content = "🎧";

            if (resetSongInfo)
            {
                PlayingSongTitle.Text = "No song playing";
                PlaybackSlider.Value = 0;
                PlaybackTimeText.Text = "0:00 / 0:00";
                _currentPlayingSong = null;
            }
            else
            {
                PlaybackSlider.Value = 0;
                if (_playback == null && _currentPlayingSong != null)
                {
                    PlaybackTimeText.Text = $"0:00 / {_currentPlayingSong.Duration}";
                }
            }
        }

        private void ReleaseActiveKeys()
        {
            while (_inputQueue.Count > 0) _inputQueue.TryTake(out _);

            IntPtr foregroundHwnd = GetForegroundWindow();
            IntPtr target = (_targetHwnd != IntPtr.Zero && foregroundHwnd != _targetHwnd) ? _targetHwnd : IntPtr.Zero;

            foreach (var kvp in _pitchToKey)
            {
                InputSimulator.SimulateKeyUp(kvp.Value, target);
            }
            _pitchToKey.Clear();
            _pitchActiveCount.Clear();
            foreach (var token in _noteReleaseTokens.Values) token.Cancel();
            _noteReleaseTokens.Clear();

            InputSimulator.ReleaseAllKeys(target);
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            StopPlayback(false);
        }

        private async void PrevButton_Click(object sender, RoutedEventArgs e)
        {
            if (_config.Main.PlaybackMode == "RandomPlaylist")
            {
                await PlayPreviousRandom();
            }
            else
            {
                if (PlaylistListView.SelectedIndex > 0)
                {
                    PlaylistListView.SelectedIndex--;
                    var song = PlaylistListView.SelectedItem as MidiSong;
                    if (song != null) await StartPlayback(song, _isPlaying);
                }
            }
        }

        private async void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_config.Main.PlaybackMode == "RandomPlaylist")
            {
                await PlayNextRandom();
            }
            else
            {
                if (PlaylistListView.SelectedIndex < PlaylistListView.Items.Count - 1)
                {
                    PlaylistListView.SelectedIndex++;
                }
                else if (_config.Main.PlaybackMode == "RepeatPlaylist")
                {
                    PlaylistListView.SelectedIndex = 0;
                }
                else
                {
                    return;
                }

                var song = PlaylistListView.SelectedItem as MidiSong;
                if (song != null) await StartPlayback(song, _isPlaying);
            }
        }

        private async System.Threading.Tasks.Task PlayNextRandom()
        {
            if (PlaylistListView.Items.Count == 0) return;

            if (_shuffleIndices.Count != PlaylistListView.Items.Count)
            {
                _shuffleIndices = CreateShuffle(PlaylistListView.Items.Count);
                _shufflePosition = 0;
            }
            else
            {
                _shufflePosition++;
                if (_shufflePosition >= _shuffleIndices.Count)
                {
                    int lastSongIdx = _shuffleIndices.Last();
                    _previousShuffle = _shuffleIndices;
                    _shuffleIndices = CreateShuffle(PlaylistListView.Items.Count, lastSongIdx);
                    _shufflePosition = 0;
                }
            }

            int index = _shuffleIndices[_shufflePosition];
            if (index >= 0 && index < PlaylistListView.Items.Count)
            {
                PlaylistListView.SelectedIndex = index;
                PlaylistListView.ScrollIntoView(PlaylistListView.SelectedItem);
                var song = PlaylistListView.SelectedItem as MidiSong;
                if (song != null) await StartPlayback(song, _isPlaying);
            }
        }

        private async System.Threading.Tasks.Task PlayPreviousRandom()
        {
            if (PlaylistListView.Items.Count == 0) return;

            if (_shuffleIndices.Count != PlaylistListView.Items.Count)
            {
                _shuffleIndices = CreateShuffle(PlaylistListView.Items.Count);
                _shufflePosition = 0;
            }
            else
            {
                _shufflePosition--;
                if (_shufflePosition < 0)
                {
                    if (_previousShuffle.Count != PlaylistListView.Items.Count)
                    {
                        _previousShuffle = CreateShuffle(PlaylistListView.Items.Count, _shuffleIndices.Count > 0 ? _shuffleIndices.First() : null);
                    }

                    _shuffleIndices = _previousShuffle;
                    _previousShuffle = new List<int>();
                    _shufflePosition = _shuffleIndices.Count - 1;
                }
            }

            int index = _shuffleIndices[_shufflePosition];
            if (index >= 0 && index < PlaylistListView.Items.Count)
            {
                PlaylistListView.SelectedIndex = index;
                PlaylistListView.ScrollIntoView(PlaylistListView.SelectedItem);
                var song = PlaylistListView.SelectedItem as MidiSong;
                if (song != null) await StartPlayback(song, _isPlaying);
            }
        }

        private List<int> CreateShuffle(int count, int? excludeFirst = null)
        {
            var list = Enumerable.Range(0, count).ToList();
            Random rng = new Random();
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                int value = list[k];
                list[k] = list[n];
                list[n] = value;
            }

            if (excludeFirst.HasValue && list.Count > 1 && list[0] == excludeFirst.Value)
            {
                int k = rng.Next(1, list.Count);
                int temp = list[0];
                list[0] = list[k];
                list[k] = temp;
            }

            return list;
        }

        private void PlaybackSlider_DragStarted(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e)
        {
            _isUserDraggingSlider = true;
        }

        private void PlaybackSlider_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            if (_playback != null)
            {
                _playback.MoveToTime(new MetricTimeSpan((long)PlaybackSlider.Value));
            }
            _isUserDraggingSlider = false;
        }

        private void PlaybackSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isUserDraggingSlider && _playback != null)
            {
                var time = new MetricTimeSpan((long)PlaybackSlider.Value);
                var duration = _playback.GetDuration<MetricTimeSpan>();
                PlaybackTimeText.Text = $"{(int)time.Minutes}:{time.Seconds:D2} / {(int)duration.Minutes}:{duration.Seconds:D2}";
            }
        }

        private void OnPlaybackFinished(object? sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(async () => {
                string mode = _config.Main.PlaybackMode;
                if (mode == "RepeatSong")
                {
                    await System.Threading.Tasks.Task.Delay(1000);
                    if (_currentPlayingSong != null) await StartPlayback(_currentPlayingSong);
                }
                else if (mode == "RepeatPlaylist")
                {
                    await System.Threading.Tasks.Task.Delay(1000);
                    NextButton_Click(this, new RoutedEventArgs());
                }
                else if (mode == "RandomPlaylist")
                {
                    await System.Threading.Tasks.Task.Delay(1000);
                    await PlayNextRandom();
                }
                else // PlayOnce
                {
                    StopPlayback(false);
                }
            }));
        }

        private void OnEventPlayed(object? sender, MidiEventPlayedEventArgs e)
        {
            if (e.Event is NoteOnEvent noteOn && noteOn.Velocity > 0)
            {
                int noteNumber = noteOn.NoteNumber + _currentTranspose;
                noteNumber = ApplyOctaveHandling(noteNumber);
                HandleNoteOn(noteNumber);
            }
            else if (e.Event is NoteOffEvent noteOff)
            {
                int noteNumber = noteOff.NoteNumber + _currentTranspose;
                noteNumber = ApplyOctaveHandling(noteNumber);
                HandleNoteOff(noteNumber);
            }
            else if (e.Event is NoteOnEvent noteOnZero && noteOnZero.Velocity == 0)
            {
                int noteNumber = noteOnZero.NoteNumber + _currentTranspose;
                noteNumber = ApplyOctaveHandling(noteNumber);
                HandleNoteOff(noteNumber);
            }
        }

        private int ApplyOctaveHandling(int noteNumber)
        {
            string octaveHandling = _config.Settings.OctaveHandling;

            int minNote = 48;
            int maxNote = 83;

            string profileName = _config.LastSelectedProfile;
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile != null)
            {
                if (profile.SelectedMode == "37") { minNote = 48; maxNote = 84; }
                else if (profile.SelectedMode == "36") { minNote = 48; maxNote = 83; }
                else if (profile.SelectedMode == "15") { minNote = 60; maxNote = 84; }
            }

            if (octaveHandling == "Wrap")
            {
                if (noteNumber < minNote)
                {
                    while (noteNumber < minNote) noteNumber += 12;
                    while (noteNumber + 12 <= maxNote) noteNumber += 12;
                }
                else if (noteNumber > maxNote)
                {
                    while (noteNumber > maxNote) noteNumber -= 12;
                    while (noteNumber - 12 >= minNote) noteNumber -= 12;
                }
            }
            else if (octaveHandling == "Clamp")
            {
                if (noteNumber < minNote)
                {
                    while (noteNumber < minNote) noteNumber += 12;
                }
                else if (noteNumber > maxNote)
                {
                    while (noteNumber > maxNote) noteNumber -= 12;
                }
            }
            return noteNumber;
        }

        private void HandleNoteOn(int noteNumber)
        {
            if (_isPreviewMode && _outputDevice != null)
            {
                _outputDevice.SendEvent(new NoteOnEvent((SevenBitNumber)noteNumber, (SevenBitNumber)100));
                return;
            }

            int minNote = 48;
            int maxNote = 83;

            string profileName = _config.LastSelectedProfile;
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile != null)
            {
                if (profile.SelectedMode == "37") { minNote = 48; maxNote = 84; }
                else if (profile.SelectedMode == "36") { minNote = 48; maxNote = 83; }
                else if (profile.SelectedMode == "15") { minNote = 60; maxNote = 84; }
            }

            if (noteNumber < minNote || noteNumber > maxNote) return;

            string? keyToPress = GetKeyForNote(noteNumber);

            if (!string.IsNullOrEmpty(keyToPress))
            {
                IntPtr foregroundHwnd = GetForegroundWindow();
                IntPtr target = (_targetHwnd != IntPtr.Zero && foregroundHwnd != _targetHwnd) ? _targetHwnd : IntPtr.Zero;

                if (profileName == "Where Winds Meet")
                {
                    _inputQueue.Add((keyToPress, target, null));
                    return;
                }

                bool holdNotes = _config.Settings.HoldNotes;

                long now = Stopwatch.GetTimestamp();
                if (_lastNoteOnTime.TryGetValue(noteNumber, out long lastTime))
                {
                    double elapsedMs = (double)(now - lastTime) * 1000 / Stopwatch.Frequency;
                    if (elapsedMs < 10) return;
                }
                _lastNoteOnTime[noteNumber] = now;

                if (holdNotes)
                {
                    if (_pitchActiveCount.TryGetValue(noteNumber, out int count) && count > 0)
                    {
                        _inputQueue.Add((keyToPress, target, false));
                    }
                    else
                    {
                        count = 0;
                    }

                    _pitchActiveCount[noteNumber] = count + 1;
                    _pitchToKey[noteNumber] = keyToPress;
                    _inputQueue.Add((keyToPress, target, true));
                }
                else
                {
                    if (_noteReleaseTokens.ContainsKey(noteNumber))
                    {
                        _inputQueue.Add((keyToPress, target, false));
                        if (_noteReleaseTokens.TryGetValue(noteNumber, out var oldToken)) oldToken.Cancel();
                    }

                    _inputQueue.Add((keyToPress, target, true));
                    _pitchToKey[noteNumber] = keyToPress;

                    var cts = new System.Threading.CancellationTokenSource();
                    _noteReleaseTokens[noteNumber] = cts;
                    int delay = _config.Settings.ReleaseDelay;

                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        try
                        {
                            await System.Threading.Tasks.Task.Delay(delay, cts.Token);
                            if (!cts.IsCancellationRequested)
                            {
                                Dispatcher.BeginInvoke(new Action(() => {
                                    if (_noteReleaseTokens.TryGetValue(noteNumber, out var currentToken) && currentToken == cts)
                                    {
                                        HandleNoteOff(noteNumber);
                                    }
                                }));
                            }
                        }
                        catch (System.Threading.Tasks.TaskCanceledException) { }
                    });
                }
            }
        }

        private void HandleNoteOff(int noteNumber)
        {
            if (_isPreviewMode && _outputDevice != null)
            {
                _outputDevice.SendEvent(new NoteOffEvent((SevenBitNumber)noteNumber, (SevenBitNumber)0));
                return;
            }

            string profileName = _config.LastSelectedProfile;
            if (profileName == "Where Winds Meet") return;

            if (!_config.Settings.HoldNotes) 
            {
                if (_pitchToKey.TryGetValue(noteNumber, out string? keyToRelease))
                {
                    IntPtr foregroundHwnd = GetForegroundWindow();
                    IntPtr target = (_targetHwnd != IntPtr.Zero && foregroundHwnd != _targetHwnd) ? _targetHwnd : IntPtr.Zero;
                    _inputQueue.Add((keyToRelease, target, false));
                    _pitchToKey.Remove(noteNumber);
                }
                return;
            }

            if (_pitchActiveCount.TryGetValue(noteNumber, out int count) && count > 0)
            {
                _pitchActiveCount[noteNumber] = --count;
                
                if (count == 0)
                {
                    if (_pitchToKey.TryGetValue(noteNumber, out string? keyToRelease))
                    {
                        IntPtr foregroundHwnd = GetForegroundWindow();
                        IntPtr target = (_targetHwnd != IntPtr.Zero && foregroundHwnd != _targetHwnd) ? _targetHwnd : IntPtr.Zero;
                        _inputQueue.Add((keyToRelease, target, false));
                        _pitchToKey.Remove(noteNumber);
                    }
                    _pitchActiveCount.Remove(noteNumber);
                }
            }
        }

        private string? GetKeyForNote(int noteNumber)
        {
            string profileName = _config.LastSelectedProfile;
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile == null) return null;

            if (profile.SelectedMode == "36")
            {
                int minNote = 48;
                int index = noteNumber - minNote;
                if (index < 0 || index >= 36) return null;

                string gridName;
                int gridIndex;

                if (index < 12) { gridName = "WWMGrid3"; gridIndex = index; }
                else if (index < 24) { gridName = "WWMGrid2"; gridIndex = index - 12; }
                else { gridName = "WWMGrid1"; gridIndex = index - 24; }

                if (profile.KeyMappings.TryGetValue(gridName, out var keys) && gridIndex < keys.Count)
                {
                    return keys[gridIndex];
                }
            }
            else if (profile.SelectedMode == "37")
            {
                int minNote = 48;
                int index = noteNumber - minNote;
                if (index < 0 || index >= 37) return null;

                string gridName;
                int gridIndex;

                if (index < 12) { gridName = "Heartopia37Grid3"; gridIndex = index; }
                else if (index < 24) { gridName = "Heartopia37Grid2"; gridIndex = index - 12; }
                else { gridName = "Heartopia37Grid1"; gridIndex = index - 24; }

                if (profile.KeyMappings.TryGetValue(gridName, out var keys) && gridIndex < keys.Count)
                {
                    return keys[gridIndex];
                }
            }
            else if (profile.SelectedMode == "15")
            {
                int[] majorNotes = { 60, 62, 64, 65, 67, 69, 71, 72, 74, 76, 77, 79, 81, 83, 84 };
                int index = Array.IndexOf(majorNotes, noteNumber);
                if (index == -1) return null;

                string gridName;
                int gridIndex;

                if (index < 7) { gridName = "Heartopia15Grid2"; gridIndex = index; }
                else { gridName = "Heartopia15Grid1"; gridIndex = index - 7; }

                if (profile.KeyMappings.TryGetValue(gridName, out var keys) && gridIndex < keys.Count)
                {
                    return keys[gridIndex];
                }
            }
            
            return null;
        }

        private void UpdateSpeedUI()
        {
            if (SpeedTextBox == null || SpeedSlider == null) return;
            if (SpeedTextBox.Text != SpeedSlider.Value.ToString("F2"))
            {
                SpeedTextBox.Text = SpeedSlider.Value.ToString("F2");
            }
            _currentSpeed = SpeedSlider.Value;
            if (_playback != null)
            {
                _playback.Speed = _currentSpeed;
            }
        }

        private void UpdatePitchUI()
        {
            if (PitchTextBox == null || PitchSlider == null) return;
            if (PitchTextBox.Text != ((int)PitchSlider.Value).ToString())
            {
                PitchTextBox.Text = ((int)PitchSlider.Value).ToString();
            }
            _currentTranspose = (int)PitchSlider.Value;
        }

        private void SpeedSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            UpdateSpeedUI();
            SaveConfig();
        }

        private void SpeedTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoading) return;
            if (double.TryParse(SpeedTextBox.Text, out double val))
            {
                if (val >= SpeedSlider.Minimum && val <= SpeedSlider.Maximum)
                {
                    SpeedSlider.Value = val;
                }
            }
        }

        private void PitchSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            UpdatePitchUI();
            SaveConfig();
        }

        private void PitchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoading) return;
            if (int.TryParse(PitchTextBox.Text, out int val))
            {
                if (val >= PitchSlider.Minimum && val <= PitchSlider.Maximum)
                {
                    PitchSlider.Value = val;
                }
            }
        }

        private void ResetAdjustments_Click(object sender, RoutedEventArgs e)
        {
            SpeedSlider.Value = 1.0;
            PitchSlider.Value = 0;
            SaveConfig();
        }

        private void ApplyConfig()
        {
            if (_config.Window.Left != -1)
            {
                this.Left = _config.Window.Left;
                this.Top = _config.Window.Top;
                this.Width = _config.Window.Width;
                this.Height = _config.Window.Height;
                this.WindowState = _config.Window.State;
            }

            FolderPathTextBox.Text = _config.Main.FolderPath;
            SortByComboBox.SelectedIndex = _config.Main.SortByIndex;
            SortDirectionButton.Content = _config.Main.SortAscending ? "▲" : "▼";
            SortDirectionButton.ToolTip = _config.Main.SortAscending ? "Ascending order" : "Descending order";
            
            switch (_config.Main.PlaybackMode)
            {
                case "PlayOnce": PlayOnceRadio.IsChecked = true; break;
                case "RepeatSong": RepeatSongRadio.IsChecked = true; break;
                case "RepeatPlaylist": RepeatPlaylistRadio.IsChecked = true; break;
                case "RandomPlaylist": RandomPlaylistRadio.IsChecked = true; break;
            }

            ShortcutPlayPause.Text = _config.Settings.ShortcutPlayPause;
            ShortcutStop.Text = _config.Settings.ShortcutStop;
            ShortcutPrev.Text = _config.Settings.ShortcutPrev;
            ShortcutNext.Text = _config.Settings.ShortcutNext;
            var scPreview = (System.Windows.Controls.TextBox)this.FindName("ShortcutPreview");
            if (scPreview != null) scPreview.Text = _config.Settings.ShortcutPreview;
            ShortcutToggleMode.Text = _config.Settings.ShortcutToggleMode;

            switch (_config.Settings.OctaveHandling)
            {
                case "Default": OctaveDefaultRadio.IsChecked = true; break;
                case "Wrap": OctaveWrapRadio.IsChecked = true; break;
                case "Clamp": OctaveClampRadio.IsChecked = true; break;
            }

            HoldNotesCheck.IsChecked = _config.Settings.HoldNotes;
            TrimLeadingSilenceCheck.IsChecked = _config.Settings.TrimLeadingSilence;
            ReleaseDelayTextBox.IsEnabled = !(_config.Settings.HoldNotes);
            ReleaseDelayTextBox.Text = _config.Settings.ReleaseDelay.ToString();
            DisablePercussionCheck.IsChecked = _config.Settings.DisablePercussion;
            var autoAdaptCheck = (System.Windows.Controls.CheckBox)this.FindName("AutoAdapt15KeysCheck");
            if (autoAdaptCheck != null) autoAdaptCheck.IsChecked = _config.Settings.AutoAdapt15Keys;
            
            RefreshMidiInputDevices();
            InitializeMidiInput();

            AlwaysOnTopCheck.IsChecked = _config.Settings.AlwaysOnTop;
            this.Topmost = AlwaysOnTopCheck.IsChecked == true;

            var autoUpdateCheckApply = (System.Windows.Controls.CheckBox)this.FindName("AutoUpdateCheck");
            if (autoUpdateCheckApply != null) autoUpdateCheckApply.IsChecked = _config.Settings.AutoUpdateEnabled;

            SpeedSlider.Value = _config.Settings.Speed;
            PitchSlider.Value = _config.Settings.Pitch;
            UpdateSpeedUI();
            UpdatePitchUI();

            _masterPlaylist.Clear();
            foreach (var songConfig in _config.Main.Playlist)
            {
                var song = new MidiSong
                {
                    Title = songConfig.Title,
                    FilePath = songConfig.FilePath,
                    BPM = songConfig.BPM,
                    TimeAdded = songConfig.TimeAdded,
                    Duration = songConfig.Duration,
                    CustomOrder = songConfig.CustomOrder,
                    DisabledTrackIndices = songConfig.DisabledTrackIndices ?? new(),
                    TrackStatesManuallySet = songConfig.TrackStatesManuallySet
                };
                _masterPlaylist.Add(song);
            }

            ProfileList.Items.Clear();
            foreach (var p in _config.Profiles)
            {
                ProfileList.Items.Add(new ListBoxItem { Content = p.Name });
            }

            foreach (var item in ProfileList.Items)
            {
                if (item is ListBoxItem lbi && lbi.Content.ToString() == _config.LastSelectedProfile)
                {
                    ProfileList.SelectedItem = item;
                    break;
                }
            }
            if (ProfileList.SelectedIndex == -1 && ProfileList.Items.Count > 0) ProfileList.SelectedIndex = 0;

            var profile = _config.Profiles.FirstOrDefault(p => p.Name == _config.LastSelectedProfile);
            if (profile != null)
            {
                ExeNameBox.Text = profile.ExeName;
                SwitchToWindowCheck.IsChecked = profile.SwitchToWindow;
                AutoPauseCheck.IsChecked = profile.AutoPause;

                ActiveProfileTextBlock.Text = profile.Name;
                if (profile.SelectedMode == "37") MainMode37.IsChecked = true;
                else if (profile.SelectedMode == "36") MainMode36.IsChecked = true;
                else if (profile.SelectedMode == "15") MainMode15.IsChecked = true;
                UpdateMainModeVisibility(profile.Name);
            }

            UpdateTooltips();
        }

        private void SaveConfig()
        {
            if (_isLoading) return;

            _config.Window.Left = this.Left;
            _config.Window.Top = this.Top;
            _config.Window.Width = this.Width;
            _config.Window.Height = this.Height;
            _config.Window.State = this.WindowState;

            _config.Main.FolderPath = FolderPathTextBox.Text;
            _config.Main.SortByIndex = SortByComboBox.SelectedIndex;
            _config.Main.SortAscending = SortDirectionButton.Content.ToString() == "▲";
            
            if (_currentPlayingSong != null)
            {
                _config.Main.LastPlayedSongPath = _currentPlayingSong.FilePath;
            }
            else if (_playback == null && PlaylistListView.SelectedItem is MidiSong selected)
            {
                _config.Main.LastPlayedSongPath = selected.FilePath;
            }
            
            if (PlayOnceRadio.IsChecked == true) _config.Main.PlaybackMode = "PlayOnce";
            else if (RepeatSongRadio.IsChecked == true) _config.Main.PlaybackMode = "RepeatSong";
            else if (RepeatPlaylistRadio.IsChecked == true) _config.Main.PlaybackMode = "RepeatPlaylist";
            else if (RandomPlaylistRadio.IsChecked == true) _config.Main.PlaybackMode = "RandomPlaylist";

            _config.Settings.ShortcutPlayPause = ShortcutPlayPause.Text;
            _config.Settings.ShortcutStop = ShortcutStop.Text;
            _config.Settings.ShortcutPrev = ShortcutPrev.Text;
            _config.Settings.ShortcutNext = ShortcutNext.Text;
            var scPreviewSave = (System.Windows.Controls.TextBox)this.FindName("ShortcutPreview");
            if (scPreviewSave != null) _config.Settings.ShortcutPreview = scPreviewSave.Text;
            _config.Settings.ShortcutToggleMode = ShortcutToggleMode.Text;

            if (OctaveDefaultRadio.IsChecked == true) _config.Settings.OctaveHandling = "Default";
            else if (OctaveWrapRadio.IsChecked == true) _config.Settings.OctaveHandling = "Wrap";
            else if (OctaveClampRadio.IsChecked == true) _config.Settings.OctaveHandling = "Clamp";

            _config.Settings.HoldNotes = HoldNotesCheck.IsChecked ?? false;
            _config.Settings.TrimLeadingSilence = TrimLeadingSilenceCheck.IsChecked ?? false;
            ReleaseDelayTextBox.IsEnabled = !(_config.Settings.HoldNotes);
            if (int.TryParse(ReleaseDelayTextBox.Text, out int delay)) _config.Settings.ReleaseDelay = delay;
            _config.Settings.DisablePercussion = DisablePercussionCheck.IsChecked ?? false;
            var autoAdaptCheck = (System.Windows.Controls.CheckBox)this.FindName("AutoAdapt15KeysCheck");
            _config.Settings.AutoAdapt15Keys = autoAdaptCheck?.IsChecked ?? false;
            _config.Settings.AlwaysOnTop = AlwaysOnTopCheck.IsChecked ?? false;
            this.Topmost = _config.Settings.AlwaysOnTop;

            var autoUpdateCheck = (System.Windows.Controls.CheckBox)this.FindName("AutoUpdateCheck");
            if (autoUpdateCheck != null) _config.Settings.AutoUpdateEnabled = autoUpdateCheck.IsChecked ?? false;

            _config.Settings.Speed = SpeedSlider.Value;
            _config.Settings.Pitch = (int)PitchSlider.Value;

            if (ProfileList.SelectedItem is ListBoxItem selectedItem)
            {
                string profileName = selectedItem.Content?.ToString() ?? "";
                _config.LastSelectedProfile = profileName;

                var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
                if (profile == null)
                {
                    profile = new ProfileConfig { Name = profileName };
                    _config.Profiles.Add(profile);
                }

                profile.ExeName = ExeNameBox.Text;
                profile.SwitchToWindow = SwitchToWindowCheck.IsChecked ?? false;
                profile.AutoPause = AutoPauseCheck.IsChecked ?? false;
                
                if (Mode37.IsChecked == true || MainMode37.IsChecked == true) profile.SelectedMode = "37";
                else if (Mode36.IsChecked == true || MainMode36.IsChecked == true) profile.SelectedMode = "36";
                else if (Mode15.IsChecked == true || MainMode15.IsChecked == true) profile.SelectedMode = "15";

                profile.KeyMappings.Clear();
                if (profileName == "Heartopia")
                {
                    SaveGridKeys(profile, Heartopia37Grid1, "Heartopia37Grid1");
                    SaveGridKeys(profile, Heartopia37Grid2, "Heartopia37Grid2");
                    SaveGridKeys(profile, Heartopia37Grid3, "Heartopia37Grid3");
                    SaveGridKeys(profile, Heartopia15Grid1, "Heartopia15Grid1");
                    SaveGridKeys(profile, Heartopia15Grid2, "Heartopia15Grid2");
                }
                else if (profileName == "Where Winds Meet")
                {
                    SaveGridKeys(profile, WWMGrid1, "WWMGrid1");
                    SaveGridKeys(profile, WWMGrid2, "WWMGrid2");
                    SaveGridKeys(profile, WWMGrid3, "WWMGrid3");
                }
                else
                {
                    SaveGridKeys(profile, WWMGrid1, "WWMGrid1");
                    SaveGridKeys(profile, WWMGrid2, "WWMGrid2");
                    SaveGridKeys(profile, WWMGrid3, "WWMGrid3");
                    SaveGridKeys(profile, Heartopia37Grid1, "Heartopia37Grid1");
                    SaveGridKeys(profile, Heartopia37Grid2, "Heartopia37Grid2");
                    SaveGridKeys(profile, Heartopia37Grid3, "Heartopia37Grid3");
                    SaveGridKeys(profile, Heartopia15Grid1, "Heartopia15Grid1");
                    SaveGridKeys(profile, Heartopia15Grid2, "Heartopia15Grid2");
                }
            }

            _config.Main.Playlist.Clear();
            foreach (MidiSong song in _masterPlaylist)
            {
                _config.Main.Playlist.Add(new MidiSongConfig
                {
                    Title = song.Title,
                    FilePath = song.FilePath,
                    BPM = song.BPM,
                    TimeAdded = song.TimeAdded,
                    Duration = song.Duration,
                    CustomOrder = song.CustomOrder,
                    DisabledTrackIndices = new List<int>(song.DisabledTrackIndices),
                    TrackStatesManuallySet = song.TrackStatesManuallySet
                });
            }

            _config.Save();
            RegisterAllHotKeys();
            
            if (!_isLoading) SwitchToGameWindow();
        }

        private void SaveGridKeys(ProfileConfig profile, System.Windows.Controls.Panel panel, string name)
        {
            var keys = new List<string>();
            foreach (var child in panel.Children)
            {
                if (child is System.Windows.Controls.TextBox tb) keys.Add(tb.Text);
            }
            profile.KeyMappings[name] = keys;
        }

        private void ResetToDefault_Click(object sender, RoutedEventArgs e)
        {
            if (ProfileList.SelectedItem is not ListBoxItem selectedItem) return;
            string profileName = selectedItem.Content?.ToString() ?? "";

            bool wasLoading = _isLoading;
            _isLoading = true;

            if (profileName == "Heartopia")
            {
                // Mode 37
                SetGridKeys(Heartopia37Grid1, new[] { "Q", "2", "W", "3", "E", "R", "5", "T", "6", "Y", "7", "U", "I" });
                SetGridKeys(Heartopia37Grid2, new[] { "Z", "S", "X", "D", "C", "V", "G", "B", "H", "N", "J", "M" });
                SetGridKeys(Heartopia37Grid3, new[] { ",", "L", ".", ";", "/", "O", "0", "P", "-", "[", "=", "]" });

                // Mode 15
                SetGridKeys(Heartopia15Grid1, new[] { "Q", "W", "E", "R", "T", "Y", "U", "I" });
                SetGridKeys(Heartopia15Grid2, new[] { "A", "S", "D", "F", "G", "H", "J" });

                ExeNameBox.Text = "xdt.exe";
                SwitchToWindowCheck.IsChecked = true;
                AutoPauseCheck.IsChecked = true;
                Mode37.IsChecked = true;
            }
            else if (profileName == "Where Winds Meet")
            {
                SetGridKeys(WWMGrid1, new[] { "Q", "Shift+Q", "W", "Ctrl+E", "E", "R", "Shift+R", "T", "Shift+T", "Y", "Ctrl+U", "U" });
                SetGridKeys(WWMGrid2, new[] { "A", "Shift+A", "S", "Ctrl+D", "D", "F", "Shift+F", "G", "Shift+G", "H", "Ctrl+J", "J" });
                SetGridKeys(WWMGrid3, new[] { "Z", "Shift+Z", "X", "Ctrl+C", "C", "V", "Shift+V", "B", "Shift+B", "N", "Ctrl+M", "M" });

                ExeNameBox.Text = "wwm.exe";
                SwitchToWindowCheck.IsChecked = false;
                AutoPauseCheck.IsChecked = false;
                Mode36.IsChecked = true;
            }
            else
            {
                SetGridKeys(Heartopia37Grid1, Array.Empty<string>());
                SetGridKeys(Heartopia37Grid2, Array.Empty<string>());
                SetGridKeys(Heartopia37Grid3, Array.Empty<string>());
                SetGridKeys(Heartopia15Grid1, Array.Empty<string>());
                SetGridKeys(Heartopia15Grid2, Array.Empty<string>());
                SetGridKeys(WWMGrid1, Array.Empty<string>());
                SetGridKeys(WWMGrid2, Array.Empty<string>());
                SetGridKeys(WWMGrid3, Array.Empty<string>());
                ExeNameBox.Text = "";
                SwitchToWindowCheck.IsChecked = false;
                AutoPauseCheck.IsChecked = false;
            }

            UpdateModeVisibility();
            _isLoading = wasLoading;
            SaveConfig();
        }

        private void LoadGridKeys(ProfileConfig profile, System.Windows.Controls.Panel panel, string name)
        {
            if (profile.KeyMappings.TryGetValue(name, out var keys))
            {
                int i = 0;
                foreach (var child in panel.Children)
                {
                    if (child is System.Windows.Controls.TextBox tb)
                    {
                        if (i < keys.Count)
                            tb.Text = keys[i++];
                        else
                            tb.Text = "";
                    }
                }
            }
            else
            {
                foreach (var child in panel.Children)
                {
                    if (child is System.Windows.Controls.TextBox tb) tb.Text = "";
                }
            }
        }

        private void SetGridKeys(System.Windows.Controls.Panel panel, string[] keys)
        {
            int i = 0;
            foreach (var child in panel.Children)
            {
                if (child is System.Windows.Controls.TextBox tb)
                {
                    if (i < keys.Length)
                    {
                        tb.Text = keys[i++];
                    }
                }
            }
        }

        private void SaveTrigger(object sender, EventArgs e)
        {
            SaveConfig();
        }

        private void Shortcut_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.OriginalSource is not System.Windows.Controls.TextBox tb) return;

            e.Handled = true;

            if (e.Key == Key.Back || e.Key == Key.Delete)
            {
                tb.Text = "";
                return;
            }

            Key key = (e.Key == Key.System) ? e.SystemKey : e.Key;
            if (key == Key.LeftCtrl || key == Key.RightCtrl ||
                key == Key.LeftAlt || key == Key.RightAlt ||
                key == Key.LeftShift || key == Key.RightShift ||
                key == Key.LWin || key == Key.RWin ||
                key == Key.Capital || key == Key.NumLock || key == Key.Scroll)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) sb.Append("Ctrl+");
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)) sb.Append("Alt+");
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) sb.Append("Shift+");
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Windows)) sb.Append("Win+");

            sb.Append(key.ToString());
            tb.Text = sb.ToString();
        }

        private void SortBy_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SortPlaylist();
            SaveConfig();
        }

        private void SortDirectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (SortDirectionButton.Content.ToString() == "▲")
            {
                SortDirectionButton.Content = "▼";
                SortDirectionButton.ToolTip = "Descending order";
            }
            else
            {
                SortDirectionButton.Content = "▲";
                SortDirectionButton.ToolTip = "Ascending order";
            }

            SortPlaylist();
            SaveConfig();
        }

        private void SortPlaylist()
        {
            if (_isLoading || PlaylistListView == null) return;

            string sortBy = (SortByComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Name";
            bool ascending = SortDirectionButton.Content.ToString() == "▲";

            string searchText = SearchTextBox?.Text ?? "";
            IEnumerable<MidiSong> filteredSongs = _masterPlaylist;
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                filteredSongs = filteredSongs.Where(s => s.Title.Contains(searchText, StringComparison.OrdinalIgnoreCase));
            }

            IEnumerable<MidiSong> sortedSongs;

            if (sortBy == "Custom Order")
            {
                sortedSongs = ascending ? filteredSongs.OrderBy(s => s.CustomOrder) : filteredSongs.OrderByDescending(s => s.CustomOrder);
            }
            else
            {
                switch (sortBy)
                {
                    case "Name":
                        sortedSongs = ascending ? filteredSongs.OrderBy(s => s.Title) : filteredSongs.OrderByDescending(s => s.Title);
                        break;
                    case "Recently Added":
                        sortedSongs = ascending ? filteredSongs.OrderBy(s => s.TimeAdded) : filteredSongs.OrderByDescending(s => s.TimeAdded);
                        break;
                    case "Duration":
                        sortedSongs = ascending ? filteredSongs.OrderBy(s => ParseDuration(s.Duration)) : filteredSongs.OrderByDescending(s => ParseDuration(s.Duration));
                        break;
                    default:
                        return;
                }
            }

            var sortedList = sortedSongs.ToList();

            PlaylistListView.Items.Clear();
            foreach (var song in sortedList)
            {
                PlaylistListView.Items.Add(song);
            }

            SongCountText.Text = $"{PlaylistListView.Items.Count} songs loaded";
            _shuffleIndices.Clear();
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SortPlaylist();
        }

        private void PlaylistListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateTracksList();
        }

        private List<TrackChunk> GetEffectiveTrackChunks(MidiFile midiFile)
        {
            var trackChunks = midiFile.GetTrackChunks().ToList();

            if (midiFile.OriginalFormat == MidiFileFormat.SingleTrack && trackChunks.Count == 1)
            {
                var firstTrack = trackChunks[0];
                var absoluteEvents = firstTrack.GetTimedEvents();

                var channelGroups = absoluteEvents
                    .GroupBy(e => (e.Event is ChannelEvent ce) ? (int)ce.Channel : -1)
                    .OrderBy(g => g.Key)
                    .ToList();

                if (channelGroups.Count > 1)
                {
                    var newChunks = new List<TrackChunk>();
                    var metaEvents = channelGroups.FirstOrDefault(g => g.Key == -1)?.ToList() ?? new List<TimedEvent>();

                    bool isFirstTrack = true;
                    foreach (var group in channelGroups)
                    {
                        if (group.Key == -1) continue;

                        // Only put meta events in the first track to avoid duplication and duration issues
                        var combinedEvents = isFirstTrack 
                            ? metaEvents.Concat(group).OrderBy(e => e.Time)
                            : group.OrderBy(e => e.Time);
                        
                        newChunks.Add(combinedEvents.ToTrackChunk());
                        isFirstTrack = false;
                    }

                    if (newChunks.Count > 0)
                        return newChunks;
                }
            }

            return trackChunks;
        }

        private bool IsTrackDisabled(MidiSong song, TrackChunk trackChunk, int trackIndex)
        {
            if (song.TrackStatesManuallySet)
            {
                return song.DisabledTrackIndices.Contains(trackIndex);
            }
            
            bool isPercussion = trackChunk.Events.OfType<ChannelEvent>().Any(e => e.Channel == 9);
            if (isPercussion && DisablePercussionCheck.IsChecked == true)
            {
                return true;
            }
            
            return song.DisabledTrackIndices.Contains(trackIndex);
        }

        private void UpdateTracksList()
        {
            if (TracksStackPanel == null) return;

            TracksStackPanel.Children.Clear();

            var selectedSong = PlaylistListView.SelectedItem as MidiSong;
            if (selectedSong == null)
            {
                TracksStackPanel.Children.Add(new System.Windows.Controls.TextBlock
                {
                    Text = "No song selected",
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x88, 0x88, 0x88)),
                    FontStyle = FontStyles.Italic
                });
                return;
            }

            if (string.IsNullOrEmpty(selectedSong.FilePath) || !File.Exists(selectedSong.FilePath))
            {
                TracksStackPanel.Children.Add(new System.Windows.Controls.TextBlock
                {
                    Text = "MIDI file not found",
                    Foreground = System.Windows.Media.Brushes.Red,
                    FontStyle = FontStyles.Italic
                });
                return;
            }

            try
            {
                Encoding midiEncoding = Encoding.UTF8;
                try
                {
                    midiEncoding = Encoding.GetEncoding(932);
                }
                catch
                {
                    midiEncoding = Encoding.UTF8;
                }

                var midiFile = MidiFile.Read(selectedSong.FilePath, new ReadingSettings
                {
                    TextEncoding = midiEncoding
                });
                var trackChunks = GetEffectiveTrackChunks(midiFile);

                int displayedTrackCount = 0;
                for (int i = 0; i < trackChunks.Count; i++)
                {
                    var trackChunk = trackChunks[i];
                    bool hasNotes = trackChunk.Events.OfType<NoteOnEvent>().Any();
                    if (!hasNotes) continue;

                    displayedTrackCount++;
                    var trackName = trackChunk.Events
                        .OfType<SequenceTrackNameEvent>()
                        .FirstOrDefault()?.Text;

                    string displayName = !string.IsNullOrWhiteSpace(trackName) 
                        ? trackName 
                        : $"Track {i + 1}";

                    bool isTrackDisabled = IsTrackDisabled(selectedSong, trackChunk, i);

                    var checkBox = new System.Windows.Controls.CheckBox
                    {
                        Content = displayName,
                        Foreground = System.Windows.Media.Brushes.White,
                        Margin = new Thickness(0, 4, 0, 4),
                        IsChecked = !isTrackDisabled,
                        Tag = i
                    };
                    checkBox.Click += TrackCheckBox_Click;

                    TracksStackPanel.Children.Add(checkBox);
                }

                if (displayedTrackCount == 0)
                {
                    TracksStackPanel.Children.Add(new System.Windows.Controls.TextBlock
                    {
                        Text = "No tracks found",
                        Foreground = System.Windows.Media.Brushes.Gray,
                        FontStyle = FontStyles.Italic
                    });
                }
            }
            catch (Exception ex)
            {
                TracksStackPanel.Children.Add(new System.Windows.Controls.TextBlock
                {
                    Text = $"Error reading tracks: {ex.Message}",
                    Foreground = System.Windows.Media.Brushes.Red,
                    FontSize = 10,
                    TextWrapping = TextWrapping.Wrap
                });
            }
        }

        private void TrackCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.CheckBox clickedCheckBox && clickedCheckBox.Tag is int clickedTrackIndex)
            {
                var selectedSong = PlaylistListView.SelectedItem as MidiSong;
                if (selectedSong != null)
                {
                    if (!selectedSong.TrackStatesManuallySet)
                    {
                        selectedSong.DisabledTrackIndices.Clear();
                        foreach (var child in TracksStackPanel.Children)
                        {
                            if (child is System.Windows.Controls.CheckBox cb && cb.Tag is int idx)
                            {
                                if (cb.IsChecked != true)
                                {
                                    selectedSong.DisabledTrackIndices.Add(idx);
                                }
                            }
                        }
                        selectedSong.TrackStatesManuallySet = true;
                    }
                    else
                    {
                        if (clickedCheckBox.IsChecked == true)
                        {
                            selectedSong.DisabledTrackIndices.Remove(clickedTrackIndex);
                        }
                        else
                        {
                            if (!selectedSong.DisabledTrackIndices.Contains(clickedTrackIndex))
                                selectedSong.DisabledTrackIndices.Add(clickedTrackIndex);
                        }
                    }
                    SaveConfig();
                }
            }
        }

        private void UpdateCustomOrderIndices()
        {
            for (int i = 0; i < _masterPlaylist.Count; i++)
            {
                _masterPlaylist[i].CustomOrder = i;
            }
        }

        private int ParseDuration(string duration)
        {
            if (string.IsNullOrEmpty(duration) || !duration.Contains(':')) return 0;
            var parts = duration.Split(':');
            if (parts.Length == 2 && int.TryParse(parts[0], out int minutes) && int.TryParse(parts[1], out int seconds))
            {
                return minutes * 60 + seconds;
            }
            return 0;
        }

        private void OpenSettings(object sender, RoutedEventArgs e)
        {
            UnregisterAllHotKeys();
            SettingsUI.Visibility = Visibility.Visible;
        }

        private void CloseSettings(object sender, RoutedEventArgs e)
        {
            SettingsUI.Visibility = Visibility.Collapsed;
            RegisterAllHotKeys();
        }

        private void RefreshMidiInputDevices_Click(object sender, RoutedEventArgs e)
        {
            RefreshMidiInputDevices();
        }

        private void RefreshMidiInputDevices()
        {
            var comboBox = (System.Windows.Controls.ComboBox)this.FindName("MidiInputComboBox");
            if (comboBox == null) return;

            string previousSelection = _config.Settings.MidiInputDevice;
            comboBox.Items.Clear();
            comboBox.Items.Add("None");

            var devices = Melanchall.DryWetMidi.Multimedia.InputDevice.GetAll();
            foreach (var device in devices)
            {
                comboBox.Items.Add(device.Name);
            }

            if (comboBox.Items.Contains(previousSelection))
            {
                comboBox.SelectedItem = previousSelection;
            }
            else
            {
                comboBox.SelectedIndex = 0;
            }
        }

        private void MidiInputComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var comboBox = (System.Windows.Controls.ComboBox)sender;
            if (_isLoading || comboBox == null) return;

            string? selectedDeviceName = comboBox.SelectedItem as string;
            if (selectedDeviceName == null) return;

            _config.Settings.MidiInputDevice = selectedDeviceName == "None" ? string.Empty : selectedDeviceName;
            SaveConfig();

            InitializeMidiInput();
        }

        private void InitializeMidiInput()
        {
            _midiInputDevice?.StopEventsListening();
            _midiInputDevice?.Dispose();
            _midiInputDevice = null;

            string deviceName = _config.Settings.MidiInputDevice;
            if (string.IsNullOrEmpty(deviceName)) return;

            try
            {
                _midiInputDevice = Melanchall.DryWetMidi.Multimedia.InputDevice.GetAll().FirstOrDefault(d => d.Name == deviceName);
                if (_midiInputDevice != null)
                {
                    _midiInputDevice.EventReceived += OnMidiInputEventReceived;
                    _midiInputDevice.StartEventsListening();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error initializing MIDI input: {ex.Message}");
            }
        }

        private void OnMidiInputEventReceived(object? sender, MidiEventReceivedEventArgs e)
        {
            if (e.Event is NoteOnEvent noteOn && noteOn.Velocity > 0)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    HandleNoteOn(noteOn.NoteNumber);
                }));
            }
            else if (e.Event is NoteOffEvent noteOff)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    HandleNoteOff(noteOff.NoteNumber);
                }));
            }
            else if (e.Event is NoteOnEvent noteOnZero && noteOnZero.Velocity == 0)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    HandleNoteOff(noteOnZero.NoteNumber);
                }));
            }
        }

        private void OpenProfiles(object sender, RoutedEventArgs e)
        {
            UnregisterAllHotKeys();
            ProfilesUI.Visibility = Visibility.Visible;
        }

        private void CloseProfiles(object sender, RoutedEventArgs e)
        {
            ProfilesUI.Visibility = Visibility.Collapsed;
            RegisterAllHotKeys();
        }

        private void LoadFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                FolderPathTextBox.Text = dialog.FolderName;
                ReloadSongs(true);
            }
        }

        private void ReloadSongs(bool reset = false)
        {
            string folderPath = FolderPathTextBox.Text;
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                PlaylistListView.Items.Clear();
                return;
            }

            try
            {
                var midiFiles = Directory.GetFiles(folderPath, "*.mid", SearchOption.AllDirectories)
                    .Concat(Directory.GetFiles(folderPath, "*.midi", SearchOption.AllDirectories))
                    .Distinct().ToList();

                if (reset)
                {
                    _masterPlaylist.Clear();
                    PlaylistListView.Items.Clear();
                    foreach (var file in midiFiles)
                    {
                        var song = CreateMidiSong(file);
                        _masterPlaylist.Add(song);
                        PlaylistListView.Items.Add(song);
                    }
                    UpdateCustomOrderIndices();
                }
                else
                {
                    var existingFiles = _masterPlaylist.Select(s => s.FilePath).ToHashSet();
                    bool added = false;
                    foreach (var file in midiFiles)
                    {
                        if (!existingFiles.Contains(file))
                        {
                            var song = CreateMidiSong(file);
                            _masterPlaylist.Add(song);
                            added = true;
                        }
                    }
                    if (added) UpdateCustomOrderIndices();
                }
                
                SortPlaylist();
                SaveConfig();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error loading MIDI files: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private MidiSong CreateMidiSong(string filePath)
        {
            var song = new MidiSong
            {
                Title = Path.GetFileNameWithoutExtension(filePath),
                FilePath = filePath,
                TimeAdded = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };

            try
            {
                var midiFile = MidiFile.Read(filePath);
                var tempoMap = midiFile.GetTempoMap();
                var bpm = tempoMap.GetTempoAtTime((MidiTimeSpan)0).BeatsPerMinute;
                song.BPM = Math.Round(bpm).ToString();
                var duration = midiFile.GetDuration<MetricTimeSpan>();
                song.Duration = $"{(int)duration.Minutes}:{duration.Seconds:D2}";
            }
            catch
            {
                song.BPM = "N/A";
                song.Duration = "0:00";
            }

            return song;
        }

        private void AddSongs_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "MIDI files (*.mid;*.midi)|*.mid;*.midi|All files (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (string file in dialog.FileNames)
                {
                    var song = CreateMidiSong(file);
                    _masterPlaylist.Add(song);
                }
                UpdateCustomOrderIndices();
                SortPlaylist();
                SaveConfig();
            }
        }

        private void RemoveSongs_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = PlaylistListView.SelectedItems.Cast<MidiSong>().ToList();
            foreach (var item in selectedItems)
            {
                _masterPlaylist.Remove(item);
            }
            UpdateCustomOrderIndices();
            SortPlaylist();
            SaveConfig();
        }

        private void ReloadSongs_Click(object sender, RoutedEventArgs e)
        {
            ReloadSongs();
        }

        private void AddProfile_Click(object sender, RoutedEventArgs e)
        {
            string newName = "New Profile";
            int counter = 1;
            while (_config.Profiles.Any(p => p.Name == newName))
            {
                newName = $"New Profile ({counter++})";
            }

            var newProfile = new ProfileConfig
            {
                Name = newName,
                ExeName = "",
                SwitchToWindow = false,
                AutoPause = false,
                SelectedMode = "37",
                IsDefault = false
            };

            _config.Profiles.Add(newProfile);
            
            var lbi = new ListBoxItem { Content = newName };
            ProfileList.Items.Add(lbi);
            ProfileList.SelectedItem = lbi;

            SaveConfig();
        }

        private void DeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (ProfileList.SelectedItem is not ListBoxItem selectedItem) return;
            string profileName = selectedItem.Content?.ToString() ?? "";
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);

            if (profile != null && !profile.IsDefault)
            {
                var result = System.Windows.MessageBox.Show($"Are you sure you want to delete the profile '{profileName}'?", "Delete Profile", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    _config.Profiles.Remove(profile);
                    ProfileList.Items.Remove(selectedItem);
                    ProfileList.SelectedIndex = 0;
                    SaveConfig();
                }
            }
        }

        private void ProfileNameBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isLoading || ProfileList.SelectedItem is not ListBoxItem selectedItem) return;
            
            string oldName = selectedItem.Content?.ToString() ?? "";
            string newName = ProfileNameBox.Text.Trim();

            if (string.IsNullOrEmpty(newName) || newName == oldName) return;

            if (_config.Profiles.Any(p => p.Name == newName && p.Name != oldName)) return;

            var profile = _config.Profiles.FirstOrDefault(p => p.Name == oldName);
            if (profile != null && !profile.IsDefault)
            {
                profile.Name = newName;
                selectedItem.Content = newName;
                ActiveProfileTextBlock.Text = newName;
                _config.LastSelectedProfile = newName;
                _config.Save(); 
            }
        }

        private void ProfileList_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (ProfileList.SelectedItem is not ListBoxItem selectedItem) return;
            
            bool wasLoading = _isLoading;
            _isLoading = true;

            string profileName = selectedItem.Content?.ToString() ?? "";
            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);

            if (profile != null)
            {
                if (!profile.IsDefault)
                {
                    Mode37.Visibility = Visibility.Visible;
                    Mode15.Visibility = Visibility.Visible;
                    Mode36.Visibility = Visibility.Visible;
                    DeleteProfileButton.IsEnabled = true;
                    ProfileNameEditPanel.Visibility = Visibility.Visible;
                    ProfileNameBox.Text = profile.Name;
                    ProfileSettingsTitle.Visibility = Visibility.Collapsed;
                }
                else
                {
                    DeleteProfileButton.IsEnabled = false;
                    ProfileNameEditPanel.Visibility = Visibility.Collapsed;
                    ProfileSettingsTitle.Visibility = Visibility.Visible;
                    
                    if (profileName == "Heartopia")
                    {
                        Mode37.Visibility = Visibility.Visible;
                        Mode15.Visibility = Visibility.Visible;
                        Mode36.Visibility = Visibility.Collapsed;
                    }
                    else if (profileName == "Where Winds Meet")
                    {
                        Mode37.Visibility = Visibility.Collapsed;
                        Mode15.Visibility = Visibility.Collapsed;
                        Mode36.Visibility = Visibility.Visible;
                    }
                }

                ExeNameBox.Text = profile.ExeName;
                SwitchToWindowCheck.IsChecked = profile.SwitchToWindow;
                AutoPauseCheck.IsChecked = profile.AutoPause;

                ActiveProfileTextBlock.Text = profile.Name;
                if (profile.SelectedMode == "37") { Mode37.IsChecked = true; MainMode37.IsChecked = true; }
                else if (profile.SelectedMode == "36") { Mode36.IsChecked = true; MainMode36.IsChecked = true; }
                else if (profile.SelectedMode == "15") { Mode15.IsChecked = true; MainMode15.IsChecked = true; }
                UpdateMainModeVisibility(profileName);

                if (profileName == "Heartopia")
                {
                    LoadGridKeys(profile, Heartopia37Grid1, "Heartopia37Grid1");
                    LoadGridKeys(profile, Heartopia37Grid2, "Heartopia37Grid2");
                    LoadGridKeys(profile, Heartopia37Grid3, "Heartopia37Grid3");
                    LoadGridKeys(profile, Heartopia15Grid1, "Heartopia15Grid1");
                    LoadGridKeys(profile, Heartopia15Grid2, "Heartopia15Grid2");
                }
                else if (profileName == "Where Winds Meet")
                {
                    LoadGridKeys(profile, WWMGrid1, "WWMGrid1");
                    LoadGridKeys(profile, WWMGrid2, "WWMGrid2");
                    LoadGridKeys(profile, WWMGrid3, "WWMGrid3");
                }
                else
                {
                    LoadGridKeys(profile, WWMGrid1, "WWMGrid1");
                    LoadGridKeys(profile, WWMGrid2, "WWMGrid2");
                    LoadGridKeys(profile, WWMGrid3, "WWMGrid3");
                    LoadGridKeys(profile, Heartopia37Grid1, "Heartopia37Grid1");
                    LoadGridKeys(profile, Heartopia37Grid2, "Heartopia37Grid2");
                    LoadGridKeys(profile, Heartopia37Grid3, "Heartopia37Grid3");
                    LoadGridKeys(profile, Heartopia15Grid1, "Heartopia15Grid1");
                    LoadGridKeys(profile, Heartopia15Grid2, "Heartopia15Grid2");
                }
            }

            UpdateModeVisibility();
            _isLoading = wasLoading;
            _config.LastSelectedProfile = profileName;
            SwitchToGameWindow();
            SaveConfig();
        }

        private void ModeChanged(object sender, RoutedEventArgs e)
        {
            if (_isLoading) return;

            if (Mode37.IsChecked == true) MainMode37.IsChecked = true;
            else if (Mode36.IsChecked == true) MainMode36.IsChecked = true;
            else if (Mode15.IsChecked == true) MainMode15.IsChecked = true;

            UpdateModeVisibility();
            SaveConfig();
        }

        private void UpdateModeVisibility()
        {
            HeartopiaKeys37.Visibility = Visibility.Collapsed;
            HeartopiaKeys15.Visibility = Visibility.Collapsed;
            WWMKeys36.Visibility = Visibility.Collapsed;

            if (Mode37.IsChecked == true) HeartopiaKeys37.Visibility = Visibility.Visible;
            if (Mode15.IsChecked == true) HeartopiaKeys15.Visibility = Visibility.Visible;
            if (Mode36.IsChecked == true) WWMKeys36.Visibility = Visibility.Visible;
        }

        private void MainModeChanged(object sender, RoutedEventArgs e)
        {
            if (_isLoading) return;

            if (MainMode37.IsChecked == true) Mode37.IsChecked = true;
            else if (MainMode36.IsChecked == true) Mode36.IsChecked = true;
            else if (MainMode15.IsChecked == true) Mode15.IsChecked = true;

            UpdateModeVisibility();
            SaveConfig();
        }

        private void UpdateMainModeVisibility(string profileName)
        {
            if (MainMode37 == null || MainMode36 == null || MainMode15 == null) return;

            var profile = _config.Profiles.FirstOrDefault(p => p.Name == profileName);
            if (profile != null && !profile.IsDefault)
            {
                MainMode37.Visibility = Visibility.Visible;
                MainMode36.Visibility = Visibility.Visible;
                MainMode15.Visibility = Visibility.Visible;
            }
            else if (profileName == "Heartopia")
            {
                MainMode37.Visibility = Visibility.Visible;
                MainMode36.Visibility = Visibility.Collapsed;
                MainMode15.Visibility = Visibility.Visible;
            }
            else if (profileName == "Where Winds Meet")
            {
                MainMode37.Visibility = Visibility.Collapsed;
                MainMode36.Visibility = Visibility.Visible;
                MainMode15.Visibility = Visibility.Collapsed;
            }
            else
            {
                MainMode37.Visibility = Visibility.Collapsed;
                MainMode36.Visibility = Visibility.Collapsed;
                MainMode15.Visibility = Visibility.Collapsed;
            }
        }

        private int DetectMajorKey(IEnumerable<Note> notes)
        {
            int[] pitchCount = new int[12];
            foreach (var n in notes)
            {
                if (n.Channel != 9) // Melanchall uses 0-15, percussion is 9
                    pitchCount[n.NoteNumber % 12]++;
            }

            int bestRoot = 0;
            int bestScore = -1;
            int[] majorPattern = { 0, 2, 4, 5, 7, 9, 11 };

            for (int root = 0; root < 12; root++)
            {
                int score = 0;
                foreach (int interval in majorPattern)
                {
                    int note = (root + interval) % 12;
                    score += pitchCount[note];
                }

                if (score > bestScore)
                {
                    bestScore = score;
                    bestRoot = root;
                }
            }

            return bestRoot;
        }

        private int FoldTo15KeyScale(int note, int root)
        {
            int[] pattern = { 0, 2, 4, 5, 7, 9, 11 };
            int pitchClass = note % 12;

            int degree = -1;
            int smallestDistance = int.MaxValue;

            for (int i = 0; i < pattern.Length; i++)
            {
                int scaleNote = (root + pattern[i]) % 12;
                int distance = Math.Abs(scaleNote - pitchClass);

                if (distance > 6) distance = 12 - distance;

                if (distance < smallestDistance)
                {
                    smallestDistance = distance;
                    degree = i;
                }
            }

            int[] cMajor = { 0, 2, 4, 5, 7, 9, 11 };
            int min15 = 48;
            int max15 = 48 + 24;

            int octaveOffset = (note - min15) / 12;
            if (octaveOffset < 0) octaveOffset = 0;
            if (octaveOffset > 2) octaveOffset = 2;

            int mapped = min15 + (octaveOffset * 12) + cMajor[degree];

            while (mapped < min15) mapped += 12;
            while (mapped > max15) mapped -= 12;

            mapped += 12;

            return mapped;
        }

        private async void PlaylistListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var item = FindAnchestor<System.Windows.Controls.ListViewItem>((DependencyObject)e.OriginalSource);
            if (item != null)
            {
                var song = PlaylistListView.ItemContainerGenerator.ItemFromContainer(item) as MidiSong;
                if (song != null)
                {
                    await StartPlayback(song, false);
                }
            }
        }

        private void PlaylistListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (PlaylistListView.View is GridView gridView)
            {
                double totalFixedWidth = 0;
                for (int i = 1; i < gridView.Columns.Count; i++)
                {
                    totalFixedWidth += gridView.Columns[i].Width;
                }

                double newWidth = PlaylistListView.ActualWidth - totalFixedWidth - 6;
                if (newWidth > 100)
                {
                    gridView.Columns[0].Width = newWidth;
                }
            }
        }

        private void PlaylistListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
            System.Windows.Controls.ListView? listView = sender as System.Windows.Controls.ListView;
            if (listView == null)
                return;
            System.Windows.Controls.ListViewItem? item = FindAnchestor<System.Windows.Controls.ListViewItem>((DependencyObject)e.OriginalSource);

            if (item != null)
            {
                MidiSong? song = listView.ItemContainerGenerator.ItemFromContainer(item) as MidiSong;
                if (song != null)
                {
                    if (listView.SelectedItems.Contains(song))
                    {
                        _draggedSongs = listView.SelectedItems.Cast<MidiSong>().ToList();
                    }
                    else
                    {
                        _draggedSongs = new List<MidiSong> { song };
                    }
                }
            }
            else
            {
                _draggedSongs = null;
            }
        }

        private void PlaylistListView_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && _draggedSongs != null && _draggedSongs.Count > 0)
            {
                if ((SortByComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() != "Custom Order") return;

                System.Windows.Point mousePos = e.GetPosition(null);
                Vector diff = _dragStartPoint - mousePos;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    var orderedSongs = _draggedSongs.OrderBy(s => _masterPlaylist.IndexOf(s)).ToList();
                    System.Windows.DataObject dragData = new System.Windows.DataObject("SongsList", orderedSongs);
                    System.Windows.DragDrop.DoDragDrop(PlaylistListView, dragData, System.Windows.DragDropEffects.Move);
                    _draggedSongs = null;
                }
            }
        }

        private void PlaylistListView_Drop(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent("SongsList"))
            {
                var droppedSongs = e.Data.GetData("SongsList") as List<MidiSong>;
                if (droppedSongs == null || droppedSongs.Count == 0) return;

                System.Windows.Controls.ListView? listView = sender as System.Windows.Controls.ListView;
                if (listView == null)
                    return;
                System.Windows.Controls.ListViewItem? targetItem = FindAnchestor<System.Windows.Controls.ListViewItem>((DependencyObject)e.OriginalSource);
                MidiSong? targetSong = targetItem != null ? listView.ItemContainerGenerator.ItemFromContainer(targetItem) as MidiSong : null;

                bool ascending = SortDirectionButton.Content.ToString() == "▲";

                int targetIndex;
                if (targetSong != null)
                {
                    targetIndex = _masterPlaylist.IndexOf(targetSong);
                }
                else
                {
                    targetIndex = ascending ? _masterPlaylist.Count : 0;
                }

                foreach (var song in droppedSongs)
                {
                    int oldIdx = _masterPlaylist.IndexOf(song);
                    if (oldIdx != -1)
                    {
                        if (oldIdx < targetIndex) targetIndex--;
                        _masterPlaylist.Remove(song);
                    }
                }

                if (targetIndex < 0) targetIndex = 0;
                if (targetIndex > _masterPlaylist.Count) targetIndex = _masterPlaylist.Count;

                for (int i = 0; i < droppedSongs.Count; i++)
                {
                    _masterPlaylist.Insert(targetIndex + i, droppedSongs[i]);
                }

                UpdateCustomOrderIndices();
                SortPlaylist();
                
                listView.SelectedItems.Clear();
                foreach (var song in droppedSongs)
                {
                    listView.SelectedItems.Add(song);
                }

                SaveConfig();
            }
        }

        private static T? FindAnchestor<T>(DependencyObject current) where T : DependencyObject
        {
            do
            {
                if (current is T) return (T)current;
                current = VisualTreeHelper.GetParent(current);
            }
            while (current != null);
            return null;
        }

        private HwndSource? _hwndSource;
        private const int WM_HOTKEY = 0x0312;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private enum HotKeyId
        {
            PlayPause = 1,
            Stop = 2,
            Prev = 3,
            Next = 4,
            Preview = 5,
            ToggleMode = 6
        }

        private IntPtr _windowHandle;

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            _windowHandle = new WindowInteropHelper(this).Handle;
            _hwndSource = HwndSource.FromHwnd(_windowHandle);
            _hwndSource.AddHook(HwndHook);
            RegisterAllHotKeys();
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY)
            {
                int id = wParam.ToInt32();
                switch ((HotKeyId)id)
                {
                    case HotKeyId.PlayPause: TogglePlayPause(false); break;
                    case HotKeyId.Stop: StopPlayback(); break;
                    case HotKeyId.Prev: PrevButton_Click(this, new RoutedEventArgs()); break;
                    case HotKeyId.Next: NextButton_Click(this, new RoutedEventArgs()); break;
                    case HotKeyId.Preview: TogglePlayPause(true); break;
                    case HotKeyId.ToggleMode: TogglePlaybackMode(); break;
                }
                handled = true;
            }
            return IntPtr.Zero;
        }

        private void RegisterAllHotKeys()
        {
            if (_windowHandle == IntPtr.Zero) return;

            if (SettingsUI != null && SettingsUI.Visibility == Visibility.Visible) return;
            if (ProfilesUI != null && ProfilesUI.Visibility == Visibility.Visible) return;

            UnregisterAllHotKeys();
            RegisterShortcut(_config.Settings.ShortcutPlayPause, HotKeyId.PlayPause);
            RegisterShortcut(_config.Settings.ShortcutStop, HotKeyId.Stop);
            RegisterShortcut(_config.Settings.ShortcutPrev, HotKeyId.Prev);
            RegisterShortcut(_config.Settings.ShortcutNext, HotKeyId.Next);
            RegisterShortcut(_config.Settings.ShortcutPreview, HotKeyId.Preview);
            RegisterShortcut(_config.Settings.ShortcutToggleMode, HotKeyId.ToggleMode);

            UpdateTooltips();
        }

        private void UpdateTooltips()
        {
            PrevButton.ToolTip = _config.Settings.ShortcutPrev;
            PlayPauseButton.ToolTip = _config.Settings.ShortcutPlayPause;
            var previewBtn = (System.Windows.Controls.Button)FindName("PreviewButton");
            if (previewBtn != null) previewBtn.ToolTip = _config.Settings.ShortcutPreview;
            StopButton.ToolTip = _config.Settings.ShortcutStop;
            NextButton.ToolTip = _config.Settings.ShortcutNext;
        }

        private void UnregisterAllHotKeys()
        {
            if (_windowHandle == IntPtr.Zero) return;
            foreach (HotKeyId id in Enum.GetValues(typeof(HotKeyId)))
            {
                UnregisterHotKey(_windowHandle, (int)id);
            }
        }

        private void RegisterShortcut(string shortcut, HotKeyId id)
        {
            var tb = id switch
            {
                HotKeyId.PlayPause => ShortcutPlayPause,
                HotKeyId.Stop => ShortcutStop,
                HotKeyId.Prev => ShortcutPrev,
                HotKeyId.Next => ShortcutNext,
                HotKeyId.Preview => (System.Windows.Controls.TextBox)this.FindName("ShortcutPreview"),
                HotKeyId.ToggleMode => ShortcutToggleMode,
                _ => null
            };

            if (tb != null)
            {
                tb.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
                tb.ClearValue(System.Windows.Controls.Control.ForegroundProperty);
                tb.ToolTip = "Click and press keys to set shortcut. Backspace to clear.";
            }

            if (string.IsNullOrWhiteSpace(shortcut)) return;

            var parts = shortcut.Split('+');
            uint modifiers = 0;
            uint vk = 0;

            foreach (var part in parts)
            {
                string p = part.Trim().ToLower();
                if (p == "ctrl") modifiers |= 0x0002;
                else if (p == "alt") modifiers |= 0x0001;
                else if (p == "shift") modifiers |= 0x0004;
                else if (p == "win") modifiers |= 0x0008;
                else
                {
                    string trimmedPart = part.Trim();
                    
                    // Try to parse as WPF Key first (since we record with Key.ToString())
                    if (Enum.TryParse<Key>(trimmedPart, true, out Key wpfKey))
                    {
                        vk = (uint)KeyInterop.VirtualKeyFromKey(wpfKey);
                    }
                    // Fallback to WinForms Keys
                    else if (Enum.TryParse<Forms.Keys>(trimmedPart, true, out Forms.Keys formsKey))
                    {
                        vk = (uint)formsKey;
                    }
                    else if (trimmedPart.Length == 1)
                    {
                        char c = trimmedPart.ToUpper()[0];
                        Forms.Keys key = c switch
                        {
                            ',' => Forms.Keys.Oemcomma,
                            '.' => Forms.Keys.OemPeriod,
                            '/' => Forms.Keys.OemQuestion,
                            ';' => Forms.Keys.Oem1,
                            '\'' => Forms.Keys.Oem7,
                            '[' => Forms.Keys.OemOpenBrackets,
                            ']' => Forms.Keys.Oem6,
                            '\\' => Forms.Keys.Oem5,
                            '-' => Forms.Keys.OemMinus,
                            '=' => Forms.Keys.Oemplus,
                            '`' => Forms.Keys.Oemtilde,
                            _ => (Forms.Keys)c
                        };
                        vk = (uint)key;
                    }
                }
            }

            if (vk != 0 && _windowHandle != IntPtr.Zero)
            {
                bool success = RegisterHotKey(_windowHandle, (int)id, modifiers, vk);
                if (!success && tb != null)
                {
                    tb.Background = System.Windows.Media.Brushes.DarkRed;
                    tb.Foreground = System.Windows.Media.Brushes.White;
                    tb.ToolTip = "HotKey registration failed. It might be in use by another application.";
                }
            }
        }

        private void TogglePlaybackMode()
        {
            string nextMode = _config.Main.PlaybackMode switch
            {
                "PlayOnce" => "RepeatSong",
                "RepeatSong" => "RepeatPlaylist",
                "RepeatPlaylist" => "RandomPlaylist",
                _ => "PlayOnce"
            };
            
            _config.Main.PlaybackMode = nextMode;
            ApplyConfig();
            SaveConfig();
        }

        protected override void OnClosed(EventArgs e)
        {
            UnregisterAllHotKeys();
            _hwndSource?.RemoveHook(HwndHook);
            StopPlayback();
            base.OnClosed(e);
        }

        protected override void OnLocationChanged(EventArgs e)
        {
            base.OnLocationChanged(e);
            SaveConfig();
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            SaveConfig();
        }

        protected override void OnStateChanged(EventArgs e)
        {
            base.OnStateChanged(e);
            SaveConfig();
        }
        private async Task MaybeAutoCheckUpdates()
        {
            if (_config.Settings.AutoUpdateEnabled)
            {
                await CheckForUpdatesAsync(showResultIfUpToDate: false, autoPrompt: true);
            }
        }

        private async void CheckUpdates_Click(object sender, RoutedEventArgs e)
        {
            await CheckForUpdatesAsync(showResultIfUpToDate: true, autoPrompt: true);
        }

        private async Task CheckForUpdatesAsync(bool showResultIfUpToDate, bool autoPrompt)
        {
            var statusText = (System.Windows.Controls.TextBlock)this.FindName("UpdateStatusText");
            try
            {
                if (statusText != null) statusText.Text = "Checking releases...";
                using var http = new HttpClient();
                http.DefaultRequestHeaders.UserAgent.ParseAdd("PlayFlowMIDI-Updater");

                var url = $"https://api.github.com/repos/{RepoOwner}/{RepoName}/releases/latest";
                var resp = await http.GetAsync(url);
                if (!resp.IsSuccessStatusCode)
                {
                    if (statusText != null) statusText.Text = "Check failed";
                    if (showResultIfUpToDate)
                        System.Windows.MessageBox.Show($"Failed to check updates: {resp.StatusCode}", "Update", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var json = await resp.Content.ReadAsStringAsync();
                var latest = JsonDocument.Parse(json).RootElement;
                var latestTag = latest.GetProperty("tag_name").GetString();
                var assets = latest.GetProperty("assets");

                var current = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0.0";
                bool hasNewer = IsNewer(latestTag, current);

                if (!hasNewer)
                {
                    if (statusText != null) statusText.Text = "Up to date";
                    if (showResultIfUpToDate)
                        System.Windows.MessageBox.Show("You already have the latest version.", "Update", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var assetUrl = FindAssetUrl(assets, latestTag);
                if (assetUrl == null)
                {
                    if (statusText != null) statusText.Text = "Asset not found";
                    System.Windows.MessageBox.Show("New version found, but release asset is missing.", "Update", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var result = autoPrompt
                    ? System.Windows.MessageBox.Show($"New version {latestTag} available. Download and install now?", "Update", MessageBoxButton.YesNo, MessageBoxImage.Question)
                    : MessageBoxResult.No;

                if (autoPrompt && result != MessageBoxResult.Yes)
                {
                    if (statusText != null) statusText.Text = "Update skipped";
                    return;
                }

                if (statusText != null) statusText.Text = "Downloading...";
                var tmpZip = Path.Combine(Path.GetTempPath(), $"PlayFlowMIDI-{latestTag}.zip");
                using (var stream = await http.GetStreamAsync(assetUrl))
                using (var file = File.Create(tmpZip))
                {
                    await stream.CopyToAsync(file);
                }

                var appDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                if (statusText != null) statusText.Text = "Installing...";
                int currentPid = Process.GetCurrentProcess().Id;
                ScheduleUpdateInstall(tmpZip, appDir, currentPid);

                if (statusText != null) statusText.Text = "Ready to restart";
                System.Windows.MessageBox.Show("Update downloaded. The app will close and restart to finish installing.", "Update", MessageBoxButton.OK, MessageBoxImage.Information);
                System.Windows.Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                if (statusText != null) statusText.Text = "Update failed";
                System.Windows.MessageBox.Show($"Update failed: {ex.Message}", "Update", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static bool IsNewer(string? tag, string currentVersion)
        {
            if (string.IsNullOrWhiteSpace(tag)) return false;

            string normalized = tag.TrimStart('v', 'V');
            if (!Version.TryParse(normalized, out var latest)) return false;
            if (!Version.TryParse(currentVersion, out var current)) return true;
            return latest > current;
        }

        private static string? FindAssetUrl(JsonElement assets, string? latestTag)
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.GetProperty("name").GetString();
                if (!string.IsNullOrEmpty(name) && name.StartsWith(AssetPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    return asset.GetProperty("browser_download_url").GetString();
                }
            }
            return null;
        }

        private static void ScheduleUpdateInstall(string zipPath, string destinationFolder, int currentPid)
        {
            var updater = Path.Combine(Path.GetTempPath(), $"PlayFlowMIDI-updater-{Guid.NewGuid():N}.cmd");
            var tmpRoot = Path.Combine(Path.GetTempPath(), $"PlayFlowMIDI-update-{Guid.NewGuid():N}");

            var script = new StringBuilder();
            script.AppendLine("@echo off");
            script.AppendLine("setlocal");
            script.AppendLine($"set ZIP=\"{zipPath}\"");
            script.AppendLine($"set DEST=\"{destinationFolder}\"");
            script.AppendLine($"set TMPDEST=\"{tmpRoot}\"");
            script.AppendLine($"set PID={currentPid}");
            script.AppendLine("timeout /t 2 /nobreak >nul");
            script.AppendLine(":waitloop");
            script.AppendLine("tasklist /FI \"PID eq %PID%\" | findstr /C:\" %PID% \" >nul");
            script.AppendLine("if %errorlevel%==0 (");
            script.AppendLine("  timeout /t 1 /nobreak >nul");
            script.AppendLine("  goto waitloop");
            script.AppendLine(")");
            script.AppendLine("powershell -NoProfile -Command \"$zip='%ZIP%'; $tmp='%TMPDEST%'; $dest='%DEST%'; Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue; Expand-Archive -Force $zip $tmp; $src=$tmp; $items=Get-ChildItem -LiteralPath $src; if($items.Count -eq 1 -and $items[0].PSIsContainer){$src=$items[0].FullName}; if(!(Test-Path -LiteralPath $dest)){ New-Item -ItemType Directory -Path $dest | Out-Null }; robocopy $src $dest /E /PURGE /NFL /NDL /NJH /NJS | Out-Null\"");
            script.AppendLine("start \"\" \"%DEST%\\PlayFlowMIDI.exe\"");
            script.AppendLine("powershell -NoProfile -Command \"Remove-Item -Recurse -Force %TMPDEST% -ErrorAction SilentlyContinue\"");
            script.AppendLine("del /f /q %ZIP%");
            script.AppendLine("del /f /q \"%~f0\"");

            File.WriteAllText(updater, script.ToString());

            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c \"\"{updater}\"\"",
                CreateNoWindow = true,
                UseShellExecute = false
            });
        }
    }
}
