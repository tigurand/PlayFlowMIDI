# PlayFlowMIDI

**PlayFlowMIDI** is a universal auto-play tool for MIDI files, designed specifically for games featuring musical performance gameplay. It allows you to play complex MIDI compositions with ease, providing high-precision input simulation and extensive customization.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-brightgreen.svg)
![Framework](https://img.shields.io/badge/framework-.NET%2010-blueviolet.svg)

## ✨ Features

- 🎹 **Universal Compatibility**: Works with virtually any game that uses keyboard inputs for musical instruments.
- 🎮 **Pre-configured Profiles**:
  - **Heartopia**
  - **Where Winds Meet**
  - **Neverness to Everness**
- 🛠️ **Custom Game Support**: Easily add your own profiles for other games by defining key mappings and executable targets.
- ⏱️ **High Precision**: Powered by [DryWetMidi](https://github.com/melanchall/drywetmidi) for low-latency and accurate note timing.
- 🚀 **Advanced Playback Controls**:
  - **Speed & Pitch**: Real-time adjustment of playback speed and transposition.
  - **Octave Handling**: Intelligently **Wrap** or **Clamp** notes that fall outside the game's supported range.
  - **Track Management**: Toggle individual MIDI tracks on/off during playback.
- ⌨️ **Global Hotkeys**: Control playback (Play/Pause, Stop, Next, Previous) even while the game is focused.
- 🤖 **Smart Automation**:
  - **Auto-Pause**: Automatically pause playback when you alt-tab out of the game.
  - **Auto-Switch**: Bring the game window to the foreground automatically when starting a song.
- 🎛️ **Three Input Methods**:
  - **Legacy**: The old behavior — real input while focused, plain window messages otherwise.
  - **Focus-Free (default)**: Sends key messages directly to a background game window. Works only with games that process window messages.
  - **Background Injection (experimental, risky)**: Injects a small hook into the game process for true unfocused play with Raw Input/Unity games. May be detected by anti-cheat.

## 🎛️ Input Methods

Choose the method in **Settings → Input Method**. All modes honor the profile's **Automatically switch to game window** and **Automatically pause when game is inactive** options.

### Legacy

The old behavior: real keyboard input (`SendInput`) is used while the game is focused, and plain window messages are used while it is unfocused. Background play isn't supported for most games.

### Focus-Free (default)

Focus-Free sends Windows keyboard messages (`WM_KEYDOWN`/`WM_KEYUP`) and activation notifications (`WM_ACTIVATE`) directly to the game window without taking focus. It uses no DLL injection, hooks, or drivers. This is the default input method.

This only works with games whose input system listens to window messages. Many modern games — especially Unity games such as **Heartopia** — read Raw Input or ignore input entirely while unfocused, so Focus-Free produces no notes in those games. **NTE** is known to work with this style of background messaging, but results vary by game and build.

### Background Injection (experimental, risky)

Background Injection is the true unfocused mode for games that ignore window messages. It injects a small .NET Framework hook (`PlayFlowMIDI.BackgroundHook.dll` via `PlayFlowMIDI.Injector.exe` and EasyHook) into the 64-bit game process. Inside the game it:

- Makes the game believe it is focused (`GetForegroundWindow`/`GetActiveWindow`/`GetFocus`).
- Returns synthetic Raw Input keyboard events when PlayFlowMIDI posts `WM_INPUT`.
- Leaves normal input untouched when the mode is idle.

PlayFlowMIDI feeds key events through shared memory, so the game keeps playing while you work in another tab with no focus changes and no global keystrokes. This was tested with **Heartopia**.

Important:

- **It injects executable code into the game process.** Anti-cheat software may detect or block it, and your game account can be banned. The mode is opt-in and shows a warning when selected.
- Only 64-bit games are supported.
- For background play, disable **Automatically switch to game window** and **Automatically pause when game is inactive** in the profile. If Auto-Switch is enabled, the game is brought forward at song start; if Auto-Pause is enabled, playback pauses when it loses focus.
- The hook stays installed in the game process, disabled, until the game exits. This lets later PlayFlowMIDI sessions reuse it without unreliable re-injection. **Restart the game to fully remove the hook.**
- If injection fails, restart the game before trying again. A process that already had a failed/disposed hook cannot always be injected a second time.

Logs are written to `%LOCALAPPDATA%\PlayFlowMIDI\background-injector.log` and `background-hook.log`. The hook performs no anti-cheat bypass, hiding, or detection evasion.

If injection reports ready but no notes play, make sure the app and `PlayFlowMIDI.BackgroundHook.dll` come from the same build. Version 1.2.0.0 and later use a protocol handshake and write `payload protocol mismatch` to `background-injector.log` instead of silently doing nothing.

### Troubleshooting

If Focus-Free does not reach a game, launch PlayFlowMIDI with the environment variable `PLAYFLOWMIDI_DEBUG=1`. It writes target-window and focus diagnostics to `%LOCALAPPDATA%\PlayFlowMIDI\input-debug.log`. The most common causes are a wrong executable name in the profile, a game running at a different privilege level (run PlayFlowMIDI as administrator), or a game that only reads Raw Input while focused.

### A note about anti-cheat

Legacy and Focus-Free use only standard, documented Windows input APIs and do not inject code into games, install kernel drivers, hook game APIs, or bypass anti-cheat detection. **Background Injection does inject a hook into the game process** (no stealth or anti-cheat bypass). Window messages are the least invasive way to reach a background game, but **no method can guarantee that a game or its anti-cheat will not react to automation**. Some anti-cheats can block synthetic input or flag automation. Please follow the terms of service of the games you play.

## 🚀 Getting Started

### Prerequisites

- Windows 10/11
- [.NET 10 Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)

### Installation

1. Download the latest release from the [Releases](https://github.com/tigurand/PlayFlowMIDI/releases) page.
2. Extract the ZIP file to a folder of your choice.
3. Run `PlayFlowMIDI.exe` (may require Administrator privileges).

## 📖 How to Use

1. **Select a Profile**: Choose your game from the profile list (e.g., Heartopia).
2. **Add MIDI Files**: Click on the Load Folder or + button to load MIDI into playlist.
3. **Configure Settings**: Adjust speed, pitch, or octave handling if necessary.
4. **Play**: Press the Play button or use the global shortcut.

### Adding Custom Games

You can add support for new games by using the in-app profile menu. You just need to map MIDI notes to the corresponding keyboard keys used by the game.

## 🛠️ Development

If you want to build the project from source:

1. Clone the repository:
   ```bash
   git clone https://github.com/tigurand/PlayFlowMIDI.git
   ```
2. Open `PlayFlowMIDI.slnx` (or the project folder) in Visual Studio.
3. Ensure you have the .NET 10 SDK installed.
4. Build the solution in `Release` mode.

## 🤝 Contributing

Contributions are welcome! Whether it's reporting bugs, suggesting features, or adding new game profiles, feel free to open an issue or submit a pull request.

## 📜 License

This project is licensed under the MIT License.

## ⚠️ Disclaimer

This tool is intended for personal use and convenience. Please be aware of the Terms of Service of the games you play. Use of automation tools may be against the rules of some online games; use at your own risk.
