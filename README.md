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
  - **Background Play**: Supports sending inputs directly to game windows even when they are not in focus (might not work on all games).

## 🚀 Getting Started

### Prerequisites

- Windows 10/11
- [.NET 10 Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)

### Installation

1. Download the latest release from the [Releases](https://github.com/tigurand/PlayFlowMIDI/releases) page.
2. Extract the ZIP file to a folder of your choice.
3. Run `PlayFlowMIDI.exe`.

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
