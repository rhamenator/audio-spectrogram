// SPDX-License-Identifier: GPL-3.0-only
#include "main_window.h"

#include <QCheckBox>
#include <QComboBox>
#include <QDoubleSpinBox>
#include <QFileDialog>
#include <QGridLayout>
#include <QLabel>
#include <QMessageBox>
#include <QPushButton>
#include <QSignalBlocker>
#include <QSpinBox>
#include <QStandardPaths>
#include <QStatusBar>
#include <QTimer>
#include <QVBoxLayout>

namespace gramqt {

namespace {

int indexOfData(const QComboBox* box, int value) {
    return std::max(0, box->findData(value));
}

} // namespace

MainWindow::MainWindow(const StartupOptions& options, QWidget* parent) : QMainWindow(parent) {
    setWindowTitle(QStringLiteral("Audio Spectrogram"));
    capture_ = new PcmCapture(this);
    view_ = new SpectrogramView(capture_, this);
    view_->setSettings(options.settings);
    view_->setMode(options.mode);

    auto* central = new QWidget(this);
    auto* layout = new QVBoxLayout(central);
    layout->setContentsMargins(6, 6, 6, 4);
    layout->setSpacing(4);
    setCentralWidget(central);

    building_ = true;
    buildControls();
    layout->addWidget(view_, 1);
    status_ = new QLabel(this);
    status_->setTextFormat(Qt::PlainText);
    statusBar()->addWidget(status_, 1);
    populateSources(options.source);

    // Reflect the startup options in the controls.
    {
        const QSignalBlocker blocker(this);
        mode_->setCurrentIndex(options.mode == ViewMode::Spectrogram ? 0 : 1);
        fft_->setCurrentIndex(indexOfData(fft_, options.settings.fftSize));
        scale_->setCurrentIndex(indexOfData(scale_, static_cast<int>(options.settings.scale)));
        palette_->setCurrentIndex(indexOfData(palette_, static_cast<int>(options.settings.palette)));
        channel_->setCurrentIndex(indexOfData(channel_, static_cast<int>(options.settings.channel)));
        minHz_->setValue(options.settings.minHz);
        maxHz_->setValue(options.settings.maxHz);
        range_->setValue(static_cast<int>(options.settings.sensitivityDb));
        history_->setValue(options.settings.historySec);
        autoRange_->setChecked(options.settings.autoRange);
        calibrate_->setChecked(options.settings.calibrateDbfs);
    }
    building_ = false;

    resize(options.width, options.height);
    statusTimer_ = new QTimer(this);
    statusTimer_->setInterval(250);
    connect(statusTimer_, &QTimer::timeout, this, &MainWindow::refreshStatus);
    statusTimer_->start();
    connect(capture_, &PcmCapture::failed, this, [this](const QString& message) {
        status_->setText(message);
    });

    if (!options.wavPath.isEmpty()) {
        QString error;
        if (!view_->loadWav(options.wavPath, &error)) {
            status_->setText(error);
        }
    } else if (!options.noCapture) {
        startCapture();
    }
}

MainWindow::~MainWindow() {
    capture_->stop();
}

void MainWindow::buildControls() {
    auto* bar = new QWidget(this);
    auto* grid = new QGridLayout(bar);
    grid->setContentsMargins(0, 0, 0, 0);
    grid->setHorizontalSpacing(6);
    grid->setVerticalSpacing(3);
    static_cast<QVBoxLayout*>(centralWidget()->layout())->addWidget(bar);

    const auto add = [&](int row, int column, const QString& label, QWidget* widget) {
        auto* caption = new QLabel(label, bar);
        grid->addWidget(caption, row, column * 2, Qt::AlignRight);
        grid->addWidget(widget, row, column * 2 + 1);
    };
    const auto combo = [&](std::initializer_list<std::pair<QString, int>> items) {
        auto* box = new QComboBox(bar);
        for (const auto& item : items) {
            box->addItem(item.first, item.second);
        }
        return box;
    };

    mode_ = new QComboBox(bar);
    mode_->addItem(QStringLiteral("Spectrogram"), 0);
    mode_->addItem(QStringLiteral("Spectrum"), 1);
    source_ = new QComboBox(bar);
    source_->setMinimumContentsLength(24);
    source_->setSizeAdjustPolicy(QComboBox::AdjustToMinimumContentsLengthWithIcon);
    freeze_ = new QCheckBox(QStringLiteral("Freeze"), bar);
    auto* reset = new QPushButton(QStringLiteral("Reset"), bar);
    auto* open = new QPushButton(QStringLiteral("Open WAV..."), bar);
    auto* live = new QPushButton(QStringLiteral("Live"), bar);
    auto* save = new QPushButton(QStringLiteral("Save PNG..."), bar);
    add(0, 0, QStringLiteral("View"), mode_);
    add(0, 1, QStringLiteral("Source"), source_);
    grid->addWidget(freeze_, 0, 4);
    grid->addWidget(reset, 0, 5);
    grid->addWidget(open, 0, 6);
    grid->addWidget(live, 0, 7);
    grid->addWidget(save, 0, 8);

    fft_ = combo({{QStringLiteral("512"), 512}, {QStringLiteral("1024"), 1024}, {QStringLiteral("2048"), 2048},
                  {QStringLiteral("4096"), 4096}, {QStringLiteral("8192"), 8192}, {QStringLiteral("16384"), 16384}});
    scale_ = combo({{QStringLiteral("Linear"), static_cast<int>(gram::FrequencyScale::Linear)},
                    {QStringLiteral("Logarithmic"), static_cast<int>(gram::FrequencyScale::Logarithmic)},
                    {QStringLiteral("Octave"), static_cast<int>(gram::FrequencyScale::Octave)}});
    palette_ = combo({{QStringLiteral("Spectrum"), static_cast<int>(gram::PaletteMode::Spectrum)},
                      {QStringLiteral("Gray"), static_cast<int>(gram::PaletteMode::Gray)},
                      {QStringLiteral("Blue"), static_cast<int>(gram::PaletteMode::Blue)},
                      {QStringLiteral("Green"), static_cast<int>(gram::PaletteMode::Green)},
                      {QStringLiteral("Amber"), static_cast<int>(gram::PaletteMode::Amber)},
                      {QStringLiteral("Ice"), static_cast<int>(gram::PaletteMode::Ice)}});
    channel_ = combo({{QStringLiteral("Mix"), static_cast<int>(gram::ChannelMode::Mix)},
                      {QStringLiteral("Left"), static_cast<int>(gram::ChannelMode::Left)},
                      {QStringLiteral("Right"), static_cast<int>(gram::ChannelMode::Right)}});
    minHz_ = new QDoubleSpinBox(bar);
    minHz_->setRange(0, 23000);
    minHz_->setDecimals(0);
    minHz_->setSuffix(QStringLiteral(" Hz"));
    maxHz_ = new QDoubleSpinBox(bar);
    maxHz_->setRange(10, 24000);
    maxHz_->setDecimals(0);
    maxHz_->setSuffix(QStringLiteral(" Hz"));
    range_ = new QSpinBox(bar);
    range_->setRange(24, 180);
    range_->setSuffix(QStringLiteral(" dB"));
    history_ = new QDoubleSpinBox(bar);
    history_->setRange(2, 120);
    history_->setDecimals(0);
    history_->setSuffix(QStringLiteral(" s"));
    autoRange_ = new QCheckBox(QStringLiteral("Auto range"), bar);
    calibrate_ = new QCheckBox(QStringLiteral("Calibrate to dBFS"), bar);

    add(1, 0, QStringLiteral("FFT"), fft_);
    add(1, 1, QStringLiteral("Scale"), scale_);
    add(1, 2, QStringLiteral("Palette"), palette_);
    add(1, 3, QStringLiteral("Channel"), channel_);
    add(2, 0, QStringLiteral("Min"), minHz_);
    add(2, 1, QStringLiteral("Max"), maxHz_);
    add(2, 2, QStringLiteral("Range"), range_);
    add(2, 3, QStringLiteral("History"), history_);
    grid->addWidget(autoRange_, 2, 8);
    grid->addWidget(calibrate_, 1, 8);

    for (QComboBox* box : {fft_, scale_, palette_, channel_}) {
        connect(box, &QComboBox::currentIndexChanged, this, [this] { applyControls(); });
    }
    for (auto* spin : {minHz_, maxHz_, history_}) {
        connect(spin, &QDoubleSpinBox::editingFinished, this, [this] { applyControls(); });
    }
    connect(range_, &QSpinBox::editingFinished, this, [this] { applyControls(); });
    connect(autoRange_, &QCheckBox::toggled, this, [this] { applyControls(); });
    connect(calibrate_, &QCheckBox::toggled, this, [this] { applyControls(); });
    connect(mode_, &QComboBox::currentIndexChanged, this, [this](int index) {
        view_->setMode(index == 0 ? ViewMode::Spectrogram : ViewMode::Spectrum);
    });
    connect(freeze_, &QCheckBox::toggled, this, [this](bool on) { view_->setFrozen(on); });
    connect(reset, &QPushButton::clicked, this, [this] { view_->reset(); });
    connect(open, &QPushButton::clicked, this, [this] { openWav(); });
    connect(live, &QPushButton::clicked, this, [this] {
        view_->showLive();
        startCapture();
    });
    connect(save, &QPushButton::clicked, this, [this] { savePng(); });
    connect(source_, &QComboBox::currentIndexChanged, this, [this] {
        if (!building_) {
            startCapture();
        }
    });
}

void MainWindow::populateSources(const QString& preferred) {
    const QSignalBlocker blocker(source_);
    source_->clear();
    const auto monitor = PcmCapture::defaultMonitorSource();
    source_->addItem(QStringLiteral("Playback monitor (default output)"), monitor);
    for (const auto& source : PcmCapture::listSources()) {
        source_->addItem(source.name + (source.monitor ? QStringLiteral("  [monitor]") : QStringLiteral("  [input]")),
                         source.name);
    }
    if (!preferred.isEmpty()) {
        const int index = source_->findData(preferred);
        if (index >= 0) {
            source_->setCurrentIndex(index);
        } else {
            source_->addItem(preferred, preferred);
            source_->setCurrentIndex(source_->count() - 1);
        }
    }
}

void MainWindow::applyControls() {
    if (building_) {
        return;
    }
    ViewSettings s = view_->settings();
    s.fftSize = fft_->currentData().toInt();
    s.scale = static_cast<gram::FrequencyScale>(scale_->currentData().toInt());
    s.palette = static_cast<gram::PaletteMode>(palette_->currentData().toInt());
    s.channel = static_cast<gram::ChannelMode>(channel_->currentData().toInt());
    s.minHz = minHz_->value();
    s.maxHz = maxHz_->value();
    s.sensitivityDb = range_->value();
    s.historySec = history_->value();
    s.autoRange = autoRange_->isChecked();
    s.calibrateDbfs = calibrate_->isChecked();
    view_->setSettings(s);
    refreshStatus();
}

void MainWindow::startCapture() {
    const auto source = source_->currentData().toString();
    if (source.isEmpty()) {
        status_->setText(QStringLiteral("No audio source found (is PulseAudio/PipeWire running?)"));
        return;
    }
    view_->reset();
    capture_->start(source);
}

void MainWindow::savePng() {
    const auto dir = QStandardPaths::writableLocation(QStandardPaths::PicturesLocation);
    const auto path = QFileDialog::getSaveFileName(this, QStringLiteral("Save spectrogram"),
                                                   dir + QStringLiteral("/spectrogram.png"),
                                                   QStringLiteral("PNG image (*.png)"));
    if (!path.isEmpty() && !view_->grab().toImage().convertToFormat(QImage::Format_RGB32).save(path)) {
        QMessageBox::warning(this, QStringLiteral("Save failed"), QStringLiteral("Could not write %1").arg(path));
    }
}

void MainWindow::openWav() {
    const auto path = QFileDialog::getOpenFileName(this, QStringLiteral("Open WAV file"), QString(),
                                                   QStringLiteral("WAV audio (*.wav)"));
    if (path.isEmpty()) {
        return;
    }
    QString error;
    if (!view_->loadWav(path, &error)) {
        QMessageBox::warning(this, QStringLiteral("Cannot open file"), error);
    }
}

void MainWindow::refreshStatus() {
    status_->setText(view_->statusText() + (capture_->running() && !view_->fileMode()
                                                ? QStringLiteral(" · capturing %1").arg(capture_->source())
                                                : QString()));
}

} // namespace gramqt
