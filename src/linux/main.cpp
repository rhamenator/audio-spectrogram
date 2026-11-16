// SPDX-License-Identifier: GPL-3.0-only
// Linux desktop front end for the audio-spectrogram DSP core.
#include <QApplication>
#include <QCommandLineParser>
#include <QImage>
#include <QTimer>

#include "main_window.h"

using gramqt::StartupOptions;
using gramqt::ViewMode;

namespace {

template <typename T>
bool pick(const QString& text, const std::initializer_list<std::pair<const char*, T>>& table, T& out) {
    for (const auto& [name, value] : table) {
        if (text.compare(QLatin1String(name), Qt::CaseInsensitive) == 0) {
            out = value;
            return true;
        }
    }
    return false;
}

} // namespace

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    QApplication::setApplicationName(QStringLiteral("audio-spectrogram"));
    QApplication::setApplicationDisplayName(QStringLiteral("Audio Spectrogram"));
    QApplication::setDesktopFileName(QStringLiteral("audio-spectrogram"));

    QCommandLineParser parser;
    parser.setApplicationDescription(
        QStringLiteral("Live spectrogram of the audio you are playing (or any input), or of a WAV file."));
    parser.addHelpOption();
    const auto opt = [&](const char* name, const char* description, const char* valueName = nullptr,
                         const char* def = nullptr) {
        QCommandLineOption option(QString::fromLatin1(name), QString::fromLatin1(description),
                                  valueName ? QString::fromLatin1(valueName) : QString(),
                                  def ? QString::fromLatin1(def) : QString());
        parser.addOption(option);
        return option;
    };
    const auto source = opt("source", "PulseAudio/PipeWire source (default: monitor of the default output)", "name");
    const auto wav = opt("wav", "Show a WAV file instead of capturing", "file");
    const auto fft = opt("fft", "FFT size, power of two (default 4096)", "n");
    const auto scale = opt("scale", "linear | log | octave (default log)", "mode");
    const auto palette = opt("palette", "spectrum | gray | blue | green | amber | ice", "name");
    const auto channel = opt("channel", "mix | left | right (default mix)", "mode");
    const auto history = opt("history", "Seconds across the plot (default 15)", "s");
    const auto minHz = opt("min-freq", "Lowest displayed frequency in Hz (default 20)", "hz");
    const auto maxHz = opt("max-freq", "Highest displayed frequency in Hz (default 20000)", "hz");
    const auto range = opt("range", "Displayed range in dB (default 90)", "db");
    const auto spectrum = opt("spectrum", "Start in the Spectrum view");
    const auto raw = opt("raw-levels", "Do not calibrate levels to dBFS (use the core's raw scale)");
    const auto fixed = opt("fixed-range", "Disable auto range (ceiling 0 dB)");
    const auto size = opt("size", "Window size WIDTHxHEIGHT (default 1100x640)", "WxH");
    const auto shot = opt("shot", "Save a screenshot of the display to this PNG and (with --exit-after-shot) quit", "file");
    const auto shotAfter = opt("shot-after", "Seconds to wait before the screenshot (default 9)", "s", "9");
    const auto exitAfter = opt("exit-after-shot", "Quit once the screenshot was written");
    parser.process(app);

    StartupOptions options;
    options.source = parser.value(source);
    options.wavPath = parser.value(wav);
    auto& s = options.settings;
    bool ok = true;
    if (parser.isSet(fft)) {
        s.fftSize = parser.value(fft).toInt(&ok);
    }
    if (parser.isSet(scale)) {
        pick<gram::FrequencyScale>(parser.value(scale),
                                   {{"linear", gram::FrequencyScale::Linear},
                                    {"log", gram::FrequencyScale::Logarithmic},
                                    {"logarithmic", gram::FrequencyScale::Logarithmic},
                                    {"octave", gram::FrequencyScale::Octave}},
                                   s.scale);
    }
    if (parser.isSet(palette)) {
        pick<gram::PaletteMode>(parser.value(palette),
                                {{"spectrum", gram::PaletteMode::Spectrum}, {"gray", gram::PaletteMode::Gray},
                                 {"blue", gram::PaletteMode::Blue}, {"green", gram::PaletteMode::Green},
                                 {"amber", gram::PaletteMode::Amber}, {"ice", gram::PaletteMode::Ice}},
                                s.palette);
    }
    if (parser.isSet(channel)) {
        pick<gram::ChannelMode>(parser.value(channel),
                                {{"mix", gram::ChannelMode::Mix}, {"left", gram::ChannelMode::Left},
                                 {"right", gram::ChannelMode::Right}},
                                s.channel);
    }
    if (parser.isSet(history)) {
        s.historySec = parser.value(history).toDouble();
    }
    if (parser.isSet(minHz)) {
        s.minHz = parser.value(minHz).toDouble();
    }
    if (parser.isSet(maxHz)) {
        s.maxHz = parser.value(maxHz).toDouble();
    }
    if (parser.isSet(range)) {
        s.sensitivityDb = parser.value(range).toDouble();
    }
    s.calibrateDbfs = !parser.isSet(raw);
    s.autoRange = !parser.isSet(fixed);
    if (parser.isSet(spectrum)) {
        options.mode = ViewMode::Spectrum;
    }
    if (parser.isSet(size)) {
        const auto parts = parser.value(size).split(QLatin1Char('x'));
        if (parts.size() == 2) {
            options.width = parts[0].toInt();
            options.height = parts[1].toInt();
        }
    }
    if (!ok || !gram::is_power_of_two(s.fftSize) || s.fftSize < 64 || s.fftSize > 65536) {
        qCritical("--fft must be a power of two between 64 and 65536");
        return 2;
    }

    gramqt::MainWindow window(options);
    window.show();
    if (parser.isSet(shot)) {
        const auto path = parser.value(shot);
        const int delayMs = static_cast<int>(parser.value(shotAfter).toDouble() * 1000.0);
        const bool quit = parser.isSet(exitAfter);
        QTimer::singleShot(delayMs, &app, [&window, path, quit, &app] {
            const bool saved = window.grab().toImage().convertToFormat(QImage::Format_RGB32).save(path);
            if (!saved) {
                qCritical("could not write %s", qPrintable(path));
            }
            if (quit) {
                app.exit(saved ? 0 : 3);
            }
        });
    }
    return app.exec();
}
