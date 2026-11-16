// SPDX-License-Identifier: GPL-3.0-only
#pragma once

#include <QImage>
#include <QElapsedTimer>
#include <QRect>
#include <QString>
#include <QTimer>
#include <QWidget>

#include <cstdint>
#include <vector>

#include "pcm_capture.h"
#include "spectrogram_core.h"

namespace gramqt {

enum class ViewMode { Spectrogram, Spectrum };

struct ViewSettings {
    int fftSize = 4096;
    gram::FrequencyScale scale = gram::FrequencyScale::Logarithmic;
    gram::PaletteMode palette = gram::PaletteMode::Spectrum;
    gram::ChannelMode channel = gram::ChannelMode::Mix;
    double minHz = 20.0;
    double maxHz = 20000.0;
    double sensitivityDb = 90.0;  // displayed range below the ceiling
    bool autoRange = true;        // ceiling follows the loudest bin in 6 dB steps
    double manualCeilingDb = 0.0; // used when autoRange is off
    double historySec = 15.0;     // time spanned by the plot width
    bool calibrateDbfs = true;    // shift the core's raw levels so a full-scale sine peaks at 0 dBFS
};

// Reference spectrogram display.
//
// Every pixel column corresponds to an exact moment of the audio: columns are spaced exactly
// (historySec / plotWidth) seconds apart on the sample clock and each column's FFT window ends at that
// sample position. If the application is late it computes the missing columns (it does not blend or
// extrapolate), so the picture depends only on the audio, not on frame timing.
class SpectrogramView : public QWidget {
    Q_OBJECT
public:
    explicit SpectrogramView(PcmCapture* capture, QWidget* parent = nullptr);

    void setSettings(const ViewSettings& settings);
    [[nodiscard]] const ViewSettings& settings() const { return settings_; }
    void setMode(ViewMode mode);
    [[nodiscard]] ViewMode mode() const { return mode_; }
    void setFrozen(bool frozen);
    [[nodiscard]] bool frozen() const { return frozen_; }

    void reset();
    bool loadWav(const QString& path, QString* error);
    void showLive();
    [[nodiscard]] bool fileMode() const { return fileMode_; }

    // Process every column that the audio received so far makes due. Called by the display timer; public so that
    // tests can drive it deterministically.
    void advance();

    [[nodiscard]] QRect plotRect() const;
    [[nodiscard]] const gram::Image& image() const { return image_; }
    [[nodiscard]] double ceilingDb() const { return ceilingDb_; }
    [[nodiscard]] double peakDb() const { return peakDb_; }
    [[nodiscard]] std::uint64_t columns() const { return columns_; }
    [[nodiscard]] std::uint64_t skipped() const { return skipped_; }
    [[nodiscard]] double framesPerColumn() const;
    [[nodiscard]] QString statusText() const;
    // Vertical position (0 = min frequency, 1 = max frequency) of a frequency under the current scale.
    [[nodiscard]] double position(double frequency) const;

signals:
    void statusChanged();

protected:
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;

private:
    void rebuildPlan();
    void resetImage();
    void rerenderHistory();
    [[nodiscard]] gram::SpectrogramSettings coreSettings() const;
    bool computeBins(std::int64_t endFrame);
    void pushColumn(std::int64_t endFrame);
    void updateCeiling();
    void padStalls();
    void drawAxes(QPainter& painter, const QRect& plot) const;
    void drawSpectrum(QPainter& painter, const QRect& plot) const;
    [[nodiscard]] std::vector<double> tickFrequencies() const;

    PcmCapture* capture_ = nullptr;
    ViewSettings settings_;
    ViewMode mode_ = ViewMode::Spectrogram;
    bool frozen_ = false;
    bool fileMode_ = false;
    QImage fileImage_;
    double fileSeconds_ = 0.0;
    QString fileName_;

    gram::FftPlan plan_{4096};
    gram::Image image_;
    std::vector<std::vector<double>> history_; // calibrated dB per displayed column (rows as in the core's column)
    std::vector<double> samples_;
    std::vector<double> bins_;   // calibrated dB bins of the most recent window
    std::vector<double> peakHold_;
    double ceilingDb_ = 0.0;
    double peakDb_ = -200.0;

    bool started_ = false;
    double nextEnd_ = 0.0;
    std::uint64_t columns_ = 0;
    std::uint64_t skipped_ = 0;
    gram::ProfilingStats profiling_;
    QElapsedTimer padClock_;
    std::int64_t padBase_ = 0;
    QElapsedTimer spectrumClock_;
    QTimer timer_;
};

} // namespace gramqt
