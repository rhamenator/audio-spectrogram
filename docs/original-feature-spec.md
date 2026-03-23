# Original Feature Spec

This document is a clean-room reconstruction of the legacy `gram.exe` behavior based on:

- binary metadata from `C:\gram\gram.exe`
- help index content from `C:\gram\gram.cnt`
- extracted help text from `C:\gram\gram.hlp`

## Confirmed identity

- Product: `Spectrogram`
- Executable: `Gram.exe`
- Version: `5.1.7.0`
- Platform: 32-bit Windows PE (`0x14C`)
- Primary APIs: `USER32`, `GDI32`, `WINMM`, `comdlg32`, `KERNEL32`

## Confirmed major modes

1. Analyze File
2. Scan File
3. Scan Input

## Confirmed capabilities

- Spectrogram display
- Scope display with line/bar modes
- Linear and logarithmic frequency scales
- Single-channel, dual-channel, and channel-specific analysis
- FFT sizes from 512 up to at least 16384
- Adjustable time stepping between FFT windows
- Spectrum averaging
- Color palette editing
- Frequency markers and cursor readouts
- Playback of whole file or visible window
- Saving audio and image files
- Printing
- Data logging
- Automatic analysis after scan stop
- Pitch detection in scan mode
- Frequency calibration

## Constraints and likely implementation details

- Legacy Win32/GDI rendering
- `WINMM` wave input/output for capture and playback
- PCM wave file support
- Custom binary `.ini` state file
- Strong emphasis on interactive desktop workflows

## Reimplementation strategy

We should split the reproduction effort into layers:

1. DSP core
   - WAV loading
   - channel routing
   - FFT
   - windowing
   - dB conversion
   - linear/log frequency mapping
   - palette mapping

2. Export and automation
   - command-line interface
   - image export
   - saved analysis presets

3. Interactive application
   - live input
   - playback
   - markers
   - UI

4. Improvements beyond the original
   - readable config files
   - higher bit-depth support
   - better export formats
   - portable cross-platform UI
   - optional SIMD / threaded FFT path

## Why not assembly first

Assembly is best reserved for specific hotspots after:

- the required behavior is understood
- the DSP math is validated
- a profiler identifies real bottlenecks

Using assembly first would make it much harder to match the original behavior quickly,
especially while the feature set is still being reconstructed.
