// SPDX-License-Identifier: GPL-3.0-only
#include "spectrogram_view.h"

#include <QFontMetrics>
#include <QPainter>
#include <QPainterPath>
#include <QResizeEvent>

#include <algorithm>
#include <cmath>

namespace gramqt {

namespace {

constexpr int kLeft = 70;
constexpr int kRight = 70;
constexpr int kTop = 10;
constexpr int kBottom = 42;
constexpr double kDbfsOffset = 12.0412; // 20 * log10(4): the Hann-windowed core reads a full-scale sine at -12 dB
constexpr double kStallSeconds = 0.25;
constexpr int kCatchUpBudget = 256;

// The core packs pixels as 0x00RRGGBB; make the alpha byte opaque so every paint path (including ARGB visuals) shows them.
QImage opaqueImage(const gram::Image& image) {
    QImage out(image.width, image.height, QImage::Format_ARGB32_Premultiplied);
    for (int y = 0; y < image.height; ++y) {
        auto* line = reinterpret_cast<std::uint32_t*>(out.scanLine(y));
        const auto* src = image.pixels.data() + static_cast<std::size_t>(y) * static_cast<std::size_t>(image.width);
        for (int x = 0; x < image.width; ++x) {
            line[x] = src[x] | 0xFF000000u;
        }
    }
    return out;
}

QString frequencyLabel(double hz) {
    if (hz >= 1000.0) {
        const double k = hz / 1000.0;
        return QStringLiteral("%1 kHz").arg(k == std::floor(k) ? QString::number(static_cast<int>(k))
                                                               : QString::number(k, 'f', 1));
    }
    return QStringLiteral("%1 Hz").arg(static_cast<int>(std::lround(hz)));
}

} // namespace

SpectrogramView::SpectrogramView(PcmCapture* capture, QWidget* parent)
    : QWidget(parent), capture_(capture) {
    setMinimumSize(420, 240);
    setAutoFillBackground(false);
    setAttribute(Qt::WA_OpaquePaintEvent);
    rebuildPlan();
    resetImage();
    timer_.setTimerType(Qt::PreciseTimer);
    timer_.setInterval(16);
    connect(&timer_, &QTimer::timeout, this, [this] {
        advance();
    });
    timer_.start();
}

QRect SpectrogramView::plotRect() const {
    return QRect(kLeft, kTop, std::max(100, width() - kLeft - kRight), std::max(80, height() - kTop - kBottom));
}

double SpectrogramView::framesPerColumn() const {
    const int w = std::max(1, image_.width);
    return PcmCapture::kRate * settings_.historySec / static_cast<double>(w);
}

double SpectrogramView::position(double frequency) const {
    const double nyquist = (fileMode_ ? fileRate_ : PcmCapture::kRate) / 2.0;
    const double lo = std::clamp(settings_.minHz, 0.0, nyquist);
    const double hi = settings_.maxHz > 0.0 ? std::min(settings_.maxHz, nyquist) : nyquist;
    switch (settings_.scale) {
    case gram::FrequencyScale::Linear:
        return hi > lo ? (frequency - lo) / (hi - lo) : 0.0;
    case gram::FrequencyScale::Logarithmic: {
        const double minLog = std::max(lo, 1.0);
        const double span = std::log10(std::max(hi, minLog)) - std::log10(minLog);
        return span > 0.0 ? (std::log10(std::max(frequency, 1.0)) - std::log10(minLog)) / span : 0.0;
    }
    case gram::FrequencyScale::Octave: {
        const double minOctave = std::max(lo, 1.0);
        const double span = std::log2(std::max(hi, minOctave) / minOctave);
        return span > 0.0 ? std::log2(std::max(frequency, 1.0) / minOctave) / span : 0.0;
    }
    }
    return 0.0;
}

gram::SpectrogramSettings SpectrogramView::coreSettings() const {
    gram::SpectrogramSettings core;
    core.fft_size = settings_.fftSize;
    core.hop_size = std::max(1, settings_.fftSize / 4);
    core.image_height = std::max(1, image_.height);
    core.min_db = ceilingDb_ - settings_.sensitivityDb;
    core.max_db = ceilingDb_;
    core.min_frequency = settings_.minHz;
    core.max_frequency = settings_.maxHz;
    core.frequency_scale = settings_.scale;
    core.palette_mode = settings_.palette;
    return core;
}

void SpectrogramView::rebuildPlan() {
    if (!gram::is_power_of_two(settings_.fftSize)) {
        settings_.fftSize = 4096;
    }
    plan_ = gram::FftPlan(settings_.fftSize);
    profiling_ = {};
}

void SpectrogramView::resetImage() {
    const auto plot = plotRect();
    image_ = gram::make_image(plot.width(), plot.height());
    history_.clear();
    started_ = false;
    columns_ = 0;
    skipped_ = 0;
    peakDb_ = -200.0;
    peakHold_.clear();
    ceilingDb_ = settings_.autoRange ? 0.0 : settings_.manualCeilingDb;
    update();
    emit statusChanged();
}

void SpectrogramView::reset() {
    fileMode_ = false;
    resetImage();
}

void SpectrogramView::setSettings(const ViewSettings& settings) {
    const bool geometryChanged = settings.fftSize != settings_.fftSize || settings.scale != settings_.scale ||
                                 settings.channel != settings_.channel || settings.minHz != settings_.minHz ||
                                 settings.maxHz != settings_.maxHz || settings.historySec != settings_.historySec ||
                                 settings.calibrateDbfs != settings_.calibrateDbfs;
    const bool fftChanged = settings.fftSize != settings_.fftSize;
    settings_ = settings;
    settings_.historySec = std::clamp(settings_.historySec, 2.0, 120.0);
    settings_.sensitivityDb = std::clamp(settings_.sensitivityDb, 12.0, 200.0);
    settings_.minHz = std::clamp(settings_.minHz, 0.0, PcmCapture::kRate / 2.0 - 1.0);
    if (settings_.maxHz <= settings_.minHz) {
        settings_.maxHz = std::min(PcmCapture::kRate / 2.0, settings_.minHz + 1.0);
    }
    if (fftChanged) {
        rebuildPlan();
    }
    if (geometryChanged) {
        resetImage(); // the stored dB history no longer matches the new analysis
    } else {
        if (!settings_.autoRange) {
            ceilingDb_ = settings_.manualCeilingDb;
        }
        rerenderHistory(); // palette / range changes keep the history
    }
}

void SpectrogramView::setMode(ViewMode mode) {
    mode_ = mode;
    spectrumClock_.invalidate();
    update();
}

void SpectrogramView::setFrozen(bool frozen) {
    frozen_ = frozen;
    if (!frozen_) {
        started_ = false; // resume from "now" instead of replaying the frozen interval
    }
    emit statusChanged();
}

void SpectrogramView::rerenderHistory() {
    if (image_.width <= 0) {
        return;
    }
    gram::clear_image(image_);
    const auto core = coreSettings();
    const int offset = image_.width - static_cast<int>(history_.size());
    for (std::size_t i = 0; i < history_.size(); ++i) {
        gram::set_image_column(image_, offset + static_cast<int>(i), history_[i], core);
    }
    update();
}

bool SpectrogramView::computeBins(std::int64_t endFrame) {
    if (!capture_ || !capture_->window(endFrame, settings_.fftSize, settings_.channel, samples_)) {
        return false;
    }
    plan_.compute_db(samples_.data(), bins_, -240.0, &profiling_);
    if (settings_.calibrateDbfs) {
        for (auto& v : bins_) {
            v += kDbfsOffset;
        }
    }
    return true;
}

void SpectrogramView::updateCeiling() {
    double top = -240.0;
    for (const double v : bins_) {
        top = std::max(top, v);
    }
    peakDb_ = top;
    if (settings_.autoRange) {
        const double target = std::clamp(std::ceil((top + 6.0) / 6.0) * 6.0, -120.0, 24.0);
        if (std::abs(target - ceilingDb_) >= 6.0) {
            ceilingDb_ = target;
        }
    }
}

// One display column from the dB bins: linear interpolation between bins at each row's centre, and the loudest bin
// when several bins share a row, so narrow peaks stay visible and low rows are not blocky.
std::vector<double> SpectrogramView::interpolatedColumn(const gram::SpectrogramSettings& core) const {
    const int rows = core.image_height;
    std::vector<double> column(static_cast<std::size_t>(rows), -240.0);
    if (bins_.size() < 2 || rows < 1) {
        return column;
    }
    const double nyquist = PcmCapture::kRate / 2.0;
    const double lo = std::clamp(core.min_frequency, 0.0, nyquist);
    const double hi = std::clamp(core.max_frequency, lo + 1.0, nyquist);
    const double last = static_cast<double>(bins_.size() - 1);
    const auto frequency = [&](double r) {
        r = std::clamp(r, 0.0, 1.0);
        if (core.frequency_scale == gram::FrequencyScale::Linear) {
            return lo + r * (hi - lo);
        }
        const double base = std::max(lo, 1.0);
        return base * std::pow(hi / base, r);
    };
    const auto binAt = [&](double r) { return std::clamp(frequency(r) * last / nyquist, 0.0, last); };
    for (int y = 0; y < rows; ++y) {
        const double r = rows == 1 ? 0.5 : static_cast<double>(rows - 1 - y) / (rows - 1);
        const double a = binAt(r - 0.5 / rows);
        const double b = binAt(r + 0.5 / rows);
        const double center = binAt(r);
        const auto first = static_cast<std::size_t>(center);
        const auto next = std::min(first + 1, bins_.size() - 1);
        double value = bins_[first] + (bins_[next] - bins_[first]) * (center - static_cast<double>(first));
        if (b - a > 1.0) {
            for (auto i = static_cast<std::size_t>(std::ceil(a)); i <= static_cast<std::size_t>(std::floor(b)); ++i) {
                value = std::max(value, bins_[i]);
            }
        }
        column[static_cast<std::size_t>(y)] = value;
    }
    return column;
}

void SpectrogramView::pushColumn(std::int64_t endFrame) {
    if (!computeBins(endFrame)) {
        return;
    }
    updateCeiling();
    const auto core = coreSettings();
    auto column = interpolatedColumn(core);
    gram::scroll_image_left(image_);
    gram::set_image_column(image_, image_.width - 1, column, core);
    history_.push_back(std::move(column));
    if (static_cast<int>(history_.size()) > image_.width) {
        history_.erase(history_.begin(), history_.end() - image_.width);
    }
    ++columns_;
}

void SpectrogramView::padStalls() {
    // A monitor source delivers nothing while no audio plays. Keep the sample clock moving with silence so the
    // time axis stays true; real data never has to catch up because it simply resumes at real time.
    if (!capture_ || !capture_->running()) {
        return;
    }
    const auto total = capture_->totalFrames();
    if (!padClock_.isValid() || total != padSeen_) {
        // Real data arrived (or the capture restarted): re-anchor the wall clock to the sample clock.
        padClock_.restart();
        padBase_ = total;
        padSeen_ = total;
        return;
    }
    const auto expected = padBase_ + static_cast<std::int64_t>(padClock_.nsecsElapsed() * PcmCapture::kRate / 1e9);
    const auto deficit = std::min<std::int64_t>(expected - total, PcmCapture::kCapacityFrames / 2);
    if (deficit > static_cast<std::int64_t>(kStallSeconds * PcmCapture::kRate) &&
        capture_->msSinceData() > static_cast<qint64>(kStallSeconds * 1000)) {
        capture_->padSilence(deficit);
        padBase_ = capture_->totalFrames();
        padSeen_ = padBase_;
        padClock_.restart();
    }
}

void SpectrogramView::advance() {
    if (frozen_ || fileMode_ || !capture_) {
        return;
    }
    padStalls();
    const std::int64_t total = capture_->totalFrames();
    if (mode_ == ViewMode::Spectrum) {
        if (!spectrumClock_.isValid() || spectrumClock_.elapsed() >= 33) {
            spectrumClock_.restart();
            if (total >= settings_.fftSize && computeBins(total)) {
                updateCeiling();
                if (peakHold_.size() != bins_.size()) {
                    peakHold_.assign(bins_.size(), -240.0);
                }
                for (std::size_t i = 0; i < bins_.size(); ++i) {
                    peakHold_[i] = std::max(peakHold_[i] - 0.5, bins_[i]);
                }
                ++columns_;
                update();
                emit statusChanged();
            }
        }
        return;
    }
    if (!started_) {
        if (total < settings_.fftSize) {
            return;
        }
        started_ = true;
        nextEnd_ = static_cast<double>(total);
        return;
    }
    const double step = framesPerColumn();
    int budget = kCatchUpBudget;
    bool drew = false;
    while (nextEnd_ <= static_cast<double>(total) && budget-- > 0) {
        pushColumn(static_cast<std::int64_t>(nextEnd_));
        nextEnd_ += step;
        drew = true;
    }
    if (nextEnd_ + PcmCapture::kRate < static_cast<double>(total)) {
        nextEnd_ = static_cast<double>(total) - step; // hopelessly behind: jump to now and count it
        ++skipped_;
    }
    if (drew) {
        update();
        emit statusChanged();
    }
}

bool SpectrogramView::loadWav(const QString& path, QString* error) {
    try {
        const auto wav = gram::load_wav(path.toStdString());
        const auto mono = gram::select_channel(wav, settings_.channel);
        if (mono.empty() || wav.sample_rate <= 0) {
            throw std::runtime_error("The file contains no audio.");
        }
        const auto plot = plotRect();
        gram::SpectrogramSettings core = coreSettings();
        core.image_height = plot.height();
        core.hop_size = std::max<int>(1, static_cast<int>(mono.size() / static_cast<std::size_t>(plot.width())));
        gram::ProfilingStats stats;
        // Auto-range from the loudest bin of the whole file.
        gram::FftPlan scan(settings_.fftSize);
        std::vector<double> bins;
        double top = -240.0;
        for (std::size_t s = 0; s + static_cast<std::size_t>(settings_.fftSize) <= mono.size();
             s += static_cast<std::size_t>(std::max(1, core.hop_size * 8))) {
            scan.compute_db(mono.data() + s, bins, -240.0);
            for (const double v : bins) {
                top = std::max(top, v + (settings_.calibrateDbfs ? kDbfsOffset : 0.0));
            }
        }
        if (settings_.autoRange) {
            ceilingDb_ = std::clamp(std::ceil((top + 6.0) / 6.0) * 6.0, -120.0, 24.0);
        } else {
            ceilingDb_ = settings_.manualCeilingDb;
        }
        core.min_db = ceilingDb_ - settings_.sensitivityDb - (settings_.calibrateDbfs ? kDbfsOffset : 0.0);
        core.max_db = ceilingDb_ - (settings_.calibrateDbfs ? kDbfsOffset : 0.0);
        core.max_frequency = std::min(settings_.maxHz, wav.sample_rate / 2.0);
        const auto image = gram::render_full_spectrogram(mono, wav.sample_rate, core, &stats);
        fileImage_ = opaqueImage(image);
        fileSeconds_ = static_cast<double>(mono.size()) / wav.sample_rate;
        fileRate_ = wav.sample_rate;
        fileName_ = path;
        fileMode_ = true;
        profiling_ = stats;
        update();
        emit statusChanged();
        return true;
    } catch (const std::exception& e) {
        if (error) {
            *error = QString::fromUtf8(e.what());
        }
        return false;
    }
}

void SpectrogramView::showLive() {
    fileMode_ = false;
    fileImage_ = QImage();
    resetImage();
}

QString SpectrogramView::statusText() const {
    if (fileMode_) {
        return QStringLiteral("File: %1 · %2 s · Resolution: %3 Hz")
            .arg(fileName_)
            .arg(fileSeconds_, 0, 'f', 1)
            .arg(PcmCapture::kRate / static_cast<double>(settings_.fftSize), 0, 'f', 3);
    }
    return QStringLiteral("Frames: %1 · FFT: %2 ms · Resolution: %3 Hz · Window: %4 ms · %5 dB range ending %6 dBFS%7")
        .arg(columns_)
        .arg(profiling_.average_fft_ms, 0, 'f', 2)
        .arg(PcmCapture::kRate / static_cast<double>(settings_.fftSize), 0, 'f', 3)
        .arg(1000.0 * settings_.fftSize / PcmCapture::kRate, 0, 'f', 1)
        .arg(settings_.sensitivityDb, 0, 'f', 0)
        .arg(ceilingDb_, 0, 'f', 0)
        .arg(skipped_ ? QStringLiteral(" · skipped %1 s").arg(skipped_) : QString());
}

std::vector<double> SpectrogramView::tickFrequencies() const {
    const double nyquist = (fileMode_ ? fileRate_ : PcmCapture::kRate) / 2.0;
    const double lo = std::clamp(settings_.minHz, 0.0, nyquist);
    const double hi = settings_.maxHz > 0.0 ? std::min(settings_.maxHz, nyquist) : nyquist;
    std::vector<double> ticks;
    if (settings_.scale == gram::FrequencyScale::Linear) {
        for (int i = 0; i <= 5; ++i) {
            ticks.push_back(lo + (hi - lo) * i / 5.0);
        }
        return ticks;
    }
    static const double base[] = {20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000};
    for (const double f : base) {
        if (f >= lo * 0.999 && f <= hi * 1.001) {
            ticks.push_back(f);
        }
    }
    if (ticks.size() < 2) {
        ticks = {lo, hi};
    }
    return ticks;
}

void SpectrogramView::drawAxes(QPainter& painter, const QRect& plot) const {
    painter.save();
    const QColor grid(57, 74, 99);
    const QColor text(150, 168, 196);
    painter.setPen(text);
    const QFontMetrics metrics(painter.font());
    for (const double f : tickFrequencies()) {
        const int y = plot.bottom() - static_cast<int>(std::lround(position(f) * (plot.height() - 1)));
        painter.setPen(grid);
        painter.drawLine(plot.left(), y, plot.right(), y);
        painter.setPen(text);
        const auto label = frequencyLabel(f);
        painter.drawText(plot.left() - 6 - metrics.horizontalAdvance(label), y + metrics.ascent() / 2 - 1, label);
        painter.drawText(plot.right() + 6, y + metrics.ascent() / 2 - 1, label);
    }
    painter.setPen(text);
    if (mode_ == ViewMode::Spectrogram) {
        const double span = fileMode_ ? fileSeconds_ : settings_.historySec;
        for (int i = 0; i <= 4; ++i) {
            const int x = plot.left() + (plot.width() - 1) * i / 4;
            const double t = fileMode_ ? span * i / 4.0 : -span * (4 - i) / 4.0;
            const auto label = fileMode_ ? QStringLiteral("%1 s").arg(t, 0, 'f', 1)
                                         : QStringLiteral("%1 s").arg(t, 0, 'f', t == std::floor(t) ? 0 : 2);
            painter.drawText(x - metrics.horizontalAdvance(label) / 2, plot.bottom() + metrics.height() + 2, label);
        }
    }
    painter.drawText(plot.left(), plot.bottom() + 2 * metrics.height() + 4,
                     QStringLiteral("%1 ... %2 dB%3")
                         .arg(std::lround(ceilingDb_ - settings_.sensitivityDb))
                         .arg(std::lround(ceilingDb_))
                         .arg(settings_.calibrateDbfs ? QStringLiteral("FS") : QStringLiteral(" (core)")));
    painter.setPen(QPen(QColor(110, 130, 160), 1));
    painter.drawRect(plot.adjusted(-1, -1, 0, 0));
    painter.restore();
}

void SpectrogramView::drawSpectrum(QPainter& painter, const QRect& plot) const {
    if (bins_.empty()) {
        return;
    }
    painter.save();
    painter.setClipRect(plot);
    painter.setRenderHint(QPainter::Antialiasing, true);
    const double floorDb = ceilingDb_ - settings_.sensitivityDb;
    const auto levelToY = [&](double db) {
        const double p = std::clamp((db - floorDb) / settings_.sensitivityDb, 0.0, 1.0);
        return plot.bottom() - p * (plot.height() - 1);
    };
    // Frequency runs left to right in this view; the horizontal position uses the same scale as the spectrogram axis.
    const auto draw = [&](const std::vector<double>& values, const QColor& color) {
        QPainterPath path;
        bool first = true;
        for (std::size_t i = 1; i < values.size(); ++i) {
            const double f = PcmCapture::kRate / 2.0 * static_cast<double>(i) / static_cast<double>(values.size() - 1);
            const double pos = position(f);
            if (pos < 0.0 || pos > 1.0) {
                continue;
            }
            const double x = plot.left() + pos * (plot.width() - 1);
            const double y = levelToY(values[i]);
            if (first) {
                path.moveTo(x, y);
                first = false;
            } else {
                path.lineTo(x, y);
            }
        }
        painter.setPen(QPen(color, 1.3));
        painter.drawPath(path);
    };
    draw(peakHold_, QColor(255, 196, 77, 150));
    draw(bins_, QColor(110, 220, 255));
    painter.restore();
}

void SpectrogramView::paintEvent(QPaintEvent*) {
    QPainter painter(this);
    painter.fillRect(rect(), QColor(17, 24, 39));
    const auto plot = plotRect();
    painter.fillRect(plot, Qt::black);
    if (fileMode_ && !fileImage_.isNull()) {
        painter.setRenderHint(QPainter::SmoothPixmapTransform, true);
        painter.drawImage(plot, fileImage_);
    } else if (mode_ == ViewMode::Spectrogram) {
        painter.drawImage(plot.topLeft(), opaqueImage(image_));
    } else {
        drawSpectrum(painter, plot);
    }
    drawAxes(painter, plot);
}

void SpectrogramView::resizeEvent(QResizeEvent* event) {
    QWidget::resizeEvent(event);
    resetImage();
}

} // namespace gramqt
