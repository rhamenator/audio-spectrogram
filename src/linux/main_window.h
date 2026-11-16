// SPDX-License-Identifier: GPL-3.0-only
#pragma once

#include <QMainWindow>
#include <QString>

#include "pcm_capture.h"
#include "spectrogram_view.h"

class QCheckBox;
class QComboBox;
class QDoubleSpinBox;
class QLabel;
class QSpinBox;
class QTimer;

namespace gramqt {

struct StartupOptions {
    QString source;      // empty = default output monitor
    QString wavPath;     // open this file instead of capturing
    ViewSettings settings;
    ViewMode mode = ViewMode::Spectrogram;
    bool noCapture = false;
    int width = 1100;
    int height = 640;
};

class MainWindow : public QMainWindow {
    Q_OBJECT
public:
    explicit MainWindow(const StartupOptions& options, QWidget* parent = nullptr);
    ~MainWindow() override;

    [[nodiscard]] SpectrogramView* view() const { return view_; }
    [[nodiscard]] PcmCapture* capture() const { return capture_; }

private:
    void buildControls();
    void applyControls();
    void populateSources(const QString& preferred);
    void startCapture();
    void savePng();
    void openWav();
    void refreshStatus();

    PcmCapture* capture_ = nullptr;
    SpectrogramView* view_ = nullptr;
    QLabel* status_ = nullptr;
    QTimer* statusTimer_ = nullptr;

    QComboBox* mode_ = nullptr;
    QComboBox* source_ = nullptr;
    QComboBox* fft_ = nullptr;
    QComboBox* scale_ = nullptr;
    QComboBox* palette_ = nullptr;
    QComboBox* channel_ = nullptr;
    QDoubleSpinBox* minHz_ = nullptr;
    QDoubleSpinBox* maxHz_ = nullptr;
    QSpinBox* range_ = nullptr;
    QDoubleSpinBox* history_ = nullptr;
    QCheckBox* autoRange_ = nullptr;
    QCheckBox* calibrate_ = nullptr;
    QCheckBox* freeze_ = nullptr;
    bool building_ = false;
};

} // namespace gramqt
