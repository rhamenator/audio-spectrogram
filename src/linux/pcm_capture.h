// SPDX-License-Identifier: GPL-3.0-only
#pragma once

#include <QByteArray>
#include <QElapsedTimer>
#include <QObject>
#include <QProcess>
#include <QString>

#include <cstdint>
#include <vector>

#include "spectrogram_core.h"

namespace gramqt {

struct SourceInfo {
    QString name;
    bool monitor = false;
};

// Stereo 48 kHz capture from PulseAudio/PipeWire through `parec`, kept in a ring buffer addressed by absolute frame
// number, so the display can ask for the window ending at any exact sample position.
class PcmCapture : public QObject {
    Q_OBJECT
public:
    static constexpr int kRate = 48000;
    static constexpr std::int64_t kCapacityFrames = 1 << 20; // about 21.8 s

    explicit PcmCapture(QObject* parent = nullptr);
    ~PcmCapture() override;

    static std::vector<SourceInfo> listSources();
    static QString defaultMonitorSource();

    bool start(const QString& source);
    void stop();
    [[nodiscard]] bool running() const;
    [[nodiscard]] const QString& source() const { return source_; }

    // Append interleaved stereo 16-bit frames (used by the capture reader, by file replay and by tests).
    void feed(const std::int16_t* stereo, std::size_t frames);
    void padSilence(std::int64_t frames);

    [[nodiscard]] std::int64_t totalFrames() const { return total_; }
    [[nodiscard]] qint64 msSinceData() const { return sinceData_.isValid() ? sinceData_.elapsed() : -1; }
    void clear();

    // Copy `count` mono samples (normalised to +-1) that end at absolute frame `endFrame` (exclusive).
    [[nodiscard]] bool window(std::int64_t endFrame, int count, gram::ChannelMode mode, std::vector<double>& out) const;

signals:
    void failed(const QString& message);

private:
    void onReadyRead();

    QProcess process_;
    QByteArray pending_;
    QString source_;
    std::vector<std::int16_t> ring_;
    std::int64_t total_ = 0;
    QElapsedTimer sinceData_;
    bool stopping_ = false;
};

} // namespace gramqt
