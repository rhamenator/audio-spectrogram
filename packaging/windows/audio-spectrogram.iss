#define AppName "Audio Spectrogram"
#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif

[Setup]
AppId={{4BA376D1-0C7C-4B9E-B5BC-0FC47EA95135}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={autopf}\Audio Spectrogram
DefaultGroupName={#AppName}
OutputDir=..\..\artifacts
OutputBaseFilename=audio-spectrogram-{#AppVersion}-windows-x64-setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "..\..\build\Release\audio-spectrogram.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\build\Release\audio-spectrogram-info.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\build\Release\audio-spectrogram-cli.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Audio Spectrogram"; Filename: "{app}\audio-spectrogram.exe"
Name: "{group}\Audio Device Information"; Filename: "{app}\audio-spectrogram-info.exe"
