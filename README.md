# audio-spectrogram

`audio-spectrogram` is the new standalone evolution of the clean-room reimplementation work for
the legacy `gram.exe` / `Spectrogram 5.1.7.0` tool.

The project now has two fronts on top of a shared DSP core:

- `gram_repro`: file-based CLI spectrogram export
- `gram_live`: Win32 default-output loopback capture/playback with a scrolling spectrogram

Current capabilities:

- Load PCM `.wav` files
- Select left, right, or mixed audio
- Compute a Hann-windowed FFT spectrogram
- Export the result as a `.ppm` image
- Support linear, logarithmic, and octave-based frequency mapping
- Capture live audio from the Windows default output device via WASAPI loopback
- Play back captured audio through `WINMM`
- Save captured audio as `.wav`
- Show average FFT timing for quick profiling
- Select sample rate, FFT size, palette, minimum and maximum displayed frequency, frequency scaling, channel mode, display mode, spectrum amplitude mode, and grid visibility in the GUI
- Offer resampled analysis rates from `16000` through `96000` Hz in the live GUI
- Stage sample-rate, FFT, scale, and frequency-range changes behind an `Apply` button while simpler view/palette toggle changes update immediately
- Block unstable sample-rate/FFT combinations with an explicit apply-time error instead of failing silently
- Adjust display sensitivity from a single top-row slider with auto-ranging
- Show capture/source status in a bottom status bar instead of truncating it in the header
- Show current media metadata in the status bar when Windows exposes it
- Write crash/apply/capture diagnostics to `gram_live.log` beside the executable
- Show frequency scales on both sides of the live graph
- Switch between a scrolling spectrogram and an instantaneous spectrum view
- Switch between mono, stereo, and automatic channel display handling
- Freeze the current display without stopping loopback capture
- Show spectrum cursor readouts plus peak-hold and averaged overlays
- Add In/Out markers with looped playback of the marked region
- Export the live window as `.png`
- Export detected pitch frames as `.csv`
- Save and load GUI presets from simple preset files
- Control waterfall history length and scroll speed from the live GUI
- Add tuner, vectorscope, and room-response style display modes
- Add note-name guides, peak labels, harmonic hints, chord hints, and transient markers
- Toggle the musical note grid and beat/transient markers separately from the main graph grid
- Show tuner-specific instrument controls in the live window without duplicating the main transport buttons
- Include a separate `gram_audio_info` utility for inspecting Windows render-endpoint capabilities

## Why C++ first

Reproducing the original in assembly would be possible, but it would slow down the
reverse-engineering and iteration loop dramatically. C++ lets us isolate the DSP core
cleanly so we can later:

- port the core to C if desired
- optimize hotspots with SIMD or assembly
- add a GUI, live audio capture, playback, and better export formats

## Build

```powershell
cmake -S E:\audio-spectrogram -B E:\audio-spectrogram\build
cmake --build E:\audio-spectrogram\build --config Release
```

## Usage

CLI:

```powershell
E:\audio-spectrogram\build\Release\gram_repro.exe input.wav output.ppm --profile
```

Live app:

```powershell
E:\audio-spectrogram\build\Release\gram_live.exe
```

The live app listens to the default multimedia output endpoint, not to any single app window.

Example with explicit options:

```powershell
E:\audio-spectrogram\build\Release\gram_repro.exe `
  input.wav output.ppm `
  --fft 2048 `
  --hop 256 `
  --height 768 `
  --channel mix `
  --palette spectrum `
  --min-db -90 `
  --max-db 0 `
  --max-freq 12000 `
  --scale octave
```

## What is reproduced so far

- File-based audio analysis
- FFT-size-driven frequency resolution
- Linear, logarithmic, and octave display behavior
- Channel selection
- Adjustable display range in dB
- Spectrogram image generation
- Real-time input capture
- Real-time scrolling spectrogram display
- Real-time instantaneous spectrum display
- Horizontal-frequency instantaneous spectrum display with amplitude axes
- Stereo stacked spectrograms and stereo overlaid spectrum traces
- Frozen-display inspection with continued background capture
- Cursor frequency/level readouts in the live status bar
- Peak-hold and averaged spectrum overlays
- Marker lines with looped playback of the selected capture region
- PNG export of the live window
- CSV export of detected note frames
- Preset save/load for repeatable setups
- Adjustable waterfall history length and scroll speed
- Tuner, vectorscope, and room-response display modes
- Note-name guides, labeled peaks, harmonic hints, chord hints, and transient markers
- Lower minimum-frequency display support down to `0 Hz`
- Independent musical-note-grid and beat-marker toggles
- Instrument-aware tuner mode with its own start/stop controls
- Adjustable display ceiling and floor beyond the old `0 dB` / `-90 dB` limits
- Switchable graph grid with side frequency scales
- Playback of captured audio
- Wave save from the live session

## Improvements already added

- Deterministic CLI workflow that is easy to automate
- Dependency-free image output
- Better defaults for overlap and windowing than many older tools used
- Shared FFT planning with precomputed bit-reversal and twiddle factors
- Built-in timing output for FFT profiling

## Next steps

- Per-session crash logging for device/reconfigure failures that still evade reproduction
- Richer marker editing and drag-to-select regions directly in the graph
- Peak labeling, pitch detection, and calibration tools
- SIMD and assembly experiments on measured hotspots

See `docs/original-feature-spec.md` for the recovered legacy feature inventory.
