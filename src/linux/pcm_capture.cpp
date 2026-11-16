// SPDX-License-Identifier: GPL-3.0-only
#include "pcm_capture.h"

#include <QStringList>

#include <algorithm>
#include <csignal>
#include <cstring>

#include <sys/prctl.h>

namespace gramqt {

namespace {

QString runPactl(const QStringList& arguments) {
    QProcess process;
    process.start(QStringLiteral("pactl"), arguments);
    if (!process.waitForFinished(2500) || process.exitCode() != 0) {
        return {};
    }
    return QString::fromUtf8(process.readAllStandardOutput());
}

} // namespace

PcmCapture::PcmCapture(QObject* parent)
    : QObject(parent), ring_(static_cast<std::size_t>(kCapacityFrames) * 2, 0) {
    connect(&process_, &QProcess::readyReadStandardOutput, this, &PcmCapture::onReadyRead);
    connect(&process_, &QProcess::errorOccurred, this, [this](QProcess::ProcessError) {
        if (!stopping_) {
            emit failed(tr("Could not run parec (is pulseaudio-utils installed?)"));
        }
    });
    connect(&process_, &QProcess::finished, this, [this](int, QProcess::ExitStatus) {
        if (!stopping_) {
            emit failed(tr("Audio capture stopped unexpectedly"));
        }
    });
}

PcmCapture::~PcmCapture() {
    stop();
}

std::vector<SourceInfo> PcmCapture::listSources() {
    std::vector<SourceInfo> sources;
    const auto lines = runPactl({QStringLiteral("list"), QStringLiteral("short"), QStringLiteral("sources")})
                           .split(QLatin1Char('\n'), Qt::SkipEmptyParts);
    for (const auto& line : lines) {
        const auto fields = line.split(QLatin1Char('\t'));
        if (fields.size() >= 2) {
            sources.push_back({fields[1], fields[1].endsWith(QStringLiteral(".monitor"))});
        }
    }
    return sources;
}

QString PcmCapture::defaultMonitorSource() {
    const auto sink = runPactl({QStringLiteral("get-default-sink")}).trimmed();
    if (!sink.isEmpty()) {
        return sink + QStringLiteral(".monitor");
    }
    for (const auto& source : listSources()) {
        if (source.monitor) {
            return source.name;
        }
    }
    return {};
}

bool PcmCapture::start(const QString& source) {
    stop();
    if (source.isEmpty()) {
        return false;
    }
    source_ = source;
    stopping_ = false;
    clear();
    process_.setProgram(QStringLiteral("parec"));
    process_.setArguments({QStringLiteral("-d"), source, QStringLiteral("--raw"), QStringLiteral("--format=s16le"),
                           QStringLiteral("--rate=48000"), QStringLiteral("--channels=2"),
                           QStringLiteral("--latency-msec=20")});
    // parec must not outlive this application.
    process_.setChildProcessModifier([] { ::prctl(PR_SET_PDEATHSIG, SIGTERM); });
    process_.start();
    return true;
}

void PcmCapture::stop() {
    stopping_ = true;
    if (process_.state() != QProcess::NotRunning) {
        process_.terminate();
        if (!process_.waitForFinished(1000)) {
            process_.kill();
            process_.waitForFinished(1000);
        }
    }
    pending_.clear();
}

bool PcmCapture::running() const {
    return process_.state() != QProcess::NotRunning;
}

void PcmCapture::clear() {
    total_ = 0;
    pending_.clear();
    sinceData_.invalidate();
    std::fill(ring_.begin(), ring_.end(), std::int16_t{0});
}

void PcmCapture::onReadyRead() {
    pending_.append(process_.readAllStandardOutput());
    const auto frames = pending_.size() / 4;
    if (frames > 0) {
        feed(reinterpret_cast<const std::int16_t*>(pending_.constData()), static_cast<std::size_t>(frames));
        pending_.remove(0, frames * 4);
    }
}

void PcmCapture::feed(const std::int16_t* stereo, std::size_t frames) {
    for (std::size_t i = 0; i < frames; ++i) {
        const auto slot = static_cast<std::size_t>((total_ + static_cast<std::int64_t>(i)) % kCapacityFrames) * 2;
        ring_[slot] = stereo[i * 2];
        ring_[slot + 1] = stereo[i * 2 + 1];
    }
    total_ += static_cast<std::int64_t>(frames);
    sinceData_.restart();
}

void PcmCapture::padSilence(std::int64_t frames) {
    for (std::int64_t i = 0; i < frames; ++i) {
        const auto slot = static_cast<std::size_t>((total_ + i) % kCapacityFrames) * 2;
        ring_[slot] = 0;
        ring_[slot + 1] = 0;
    }
    total_ += frames; // deliberately does not restart sinceData_: padding is not real data
}

bool PcmCapture::window(std::int64_t endFrame, int count, gram::ChannelMode mode, std::vector<double>& out) const {
    const std::int64_t start = endFrame - count;
    if (count <= 0 || start < 0 || endFrame > total_ || start < total_ - kCapacityFrames) {
        return false;
    }
    out.resize(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) {
        const auto slot = static_cast<std::size_t>((start + i) % kCapacityFrames) * 2;
        const double left = ring_[slot] / 32768.0;
        const double right = ring_[slot + 1] / 32768.0;
        out[static_cast<std::size_t>(i)] =
            mode == gram::ChannelMode::Left ? left : mode == gram::ChannelMode::Right ? right : 0.5 * (left + right);
    }
    return true;
}

} // namespace gramqt
