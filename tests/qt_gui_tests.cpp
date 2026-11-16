// SPDX-License-Identifier: GPL-3.0-only
// Headless checks of the Linux Qt front end (run with QT_QPA_PLATFORM=offscreen).
#include <QApplication>
#include <QTemporaryDir>

#include <cmath>
#include <cstdio>
#include <numbers>
#include <vector>

#include "main_window.h"

namespace {

int failures = 0;

void check(bool ok, const char* what) {
    std::printf("%s %s\n", ok ? "ok  " : "FAIL", what);
    if (!ok) {
        ++failures;
    }
}

std::vector<std::int16_t> tone(double hz, double amplitude, std::int64_t startFrame, int frames) {
    std::vector<std::int16_t> pcm(static_cast<std::size_t>(frames) * 2);
    for (int i = 0; i < frames; ++i) {
        const double t = static_cast<double>(startFrame + i) / gramqt::PcmCapture::kRate;
        const auto v = static_cast<std::int16_t>(std::lround(amplitude * 32767.0 * std::sin(2.0 * std::numbers::pi * hz * t)));
        pcm[static_cast<std::size_t>(i) * 2] = v;
        pcm[static_cast<std::size_t>(i) * 2 + 1] = v;
    }
    return pcm;
}

// Feed audio in 20 ms chunks and let the view run, like a live capture would.
void play(gramqt::PcmCapture& capture, gramqt::SpectrogramView& view, double hz, double amplitude, double seconds) {
    constexpr int chunk = gramqt::PcmCapture::kRate / 50;
    const int chunks = static_cast<int>(seconds * 50);
    for (int i = 0; i < chunks; ++i) {
        const auto pcm = tone(hz, amplitude, capture.totalFrames(), chunk);
        capture.feed(pcm.data(), chunk);
        view.advance();
    }
}

int brightestRow(const gram::Image& image, int x) {
    int best = 0;
    int bestScore = -1;
    for (int y = 0; y < image.height; ++y) {
        const auto px = image.pixels[static_cast<std::size_t>(y * image.width + x)];
        const int score = static_cast<int>((px >> 16) & 0xFF) + static_cast<int>((px >> 8) & 0xFF) + static_cast<int>(px & 0xFF);
        if (score > bestScore) {
            bestScore = score;
            best = y;
        }
    }
    return best;
}

} // namespace

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    gramqt::PcmCapture capture;
    gramqt::SpectrogramView view(&capture);
    view.resize(1000, 600);
    view.show();
    QApplication::processEvents();
    const auto plot = view.plotRect();

    // 1. Start the schedule, then play 3 s of a 1 kHz tone at -12 dBFS.
    play(capture, view, 1000.0, 0.25, 1.0);
    const auto startColumns = view.columns();
    play(capture, view, 1000.0, 0.25, 3.0);
    const double expected = 3.0 * plot.width() / view.settings().historySec;
    const auto made = static_cast<double>(view.columns() - startColumns);
    check(std::abs(made - expected) <= 2.0, "columns are spaced exactly history/width seconds apart on the sample clock");
    check(view.skipped() == 0, "no time was skipped");

    // 2. Level calibration: -12 dBFS tone reads within the Hann scalloping loss of its true level.
    check(view.peakDb() < -11.5 && view.peakDb() > -13.7, "a -12 dBFS tone reads about -12 dBFS (calibrated)");

    // 3. The tone sits on the 1 kHz row of the log axis.
    const auto& image = view.image();
    const int row = brightestRow(image, image.width - 1);
    const double expectedRow = (1.0 - view.position(1000.0)) * (image.height - 1);
    check(std::abs(row - expectedRow) <= 4.0, "the 1 kHz tone is drawn on the 1 kHz row");

    {
        const auto painted = view.grab().toImage().convertToFormat(QImage::Format_RGB32);
        const QRgb p = painted.pixel(plot.right() - 1, plot.top() + static_cast<int>(expectedRow));
        check(qRed(p) + qGreen(p) + qBlue(p) > 150, "the tone is visible in the painted widget");
    }

    // 4. Freeze stops the display but not the audio.
    const auto before = view.columns();
    view.setFrozen(true);
    play(capture, view, 1000.0, 0.25, 0.5);
    check(view.columns() == before, "frozen display does not advance");
    view.setFrozen(false);

    // 5. Changing the palette keeps the history.
    play(capture, view, 1000.0, 0.25, 1.0);
    auto settings = view.settings();
    settings.palette = gram::PaletteMode::Gray;
    view.setSettings(settings);
    check(view.columns() > 0 && brightestRow(view.image(), view.image().width - 1) > 0,
          "changing the palette keeps the history");

    // 6. A change of analysis restarts the display.
    settings.fftSize = 8192;
    view.setSettings(settings);
    check(view.columns() == 0, "changing the FFT size restarts the display");

    // 7. Spectrum view computes from the latest window.
    view.setMode(gramqt::ViewMode::Spectrum);
    play(capture, view, 440.0, 0.25, 0.5);
    QApplication::processEvents();
    check(view.columns() > 0, "spectrum view updates");
    view.setMode(gramqt::ViewMode::Spectrogram);

    // 8. A WAV file renders.
    QTemporaryDir dir;
    std::vector<std::int16_t> mono(48000 * 2);
    for (std::size_t i = 0; i < mono.size(); ++i) {
        mono[i] = static_cast<std::int16_t>(std::lround(8000.0 * std::sin(2.0 * std::numbers::pi * 2000.0 * static_cast<double>(i) / 48000.0)));
    }
    const auto path = dir.filePath(QStringLiteral("tone.wav"));
    gram::write_wav_mono_16(path.toStdString(), 48000, mono);
    QString error;
    check(view.loadWav(path, &error) && view.fileMode(), "a WAV file renders");
    check(!view.loadWav(dir.filePath(QStringLiteral("missing.wav")), &error) && !error.isEmpty(), "a missing file reports an error");

    // 8b. A restarted capture with no data must not be padded with the old wall-clock deficit.
    capture.clear();
    view.reset();
    view.setMode(gramqt::ViewMode::Spectrogram);
    view.advance();
    check(capture.totalFrames() < gramqt::PcmCapture::kRate, "restarting the capture does not pad a stale deficit");

    // 9. The whole window builds and grabs.
    gramqt::StartupOptions options;
    options.noCapture = true;
    gramqt::MainWindow window(options);
    window.show();
    QApplication::processEvents();
    check(!window.grab().isNull(), "the main window renders");

    std::printf("%s\n", failures == 0 ? "ALL PASSED" : "FAILED");
    return failures == 0 ? 0 : 1;
}
