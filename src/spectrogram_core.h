#pragma once

#include <complex>
#include <cstdint>
#include <string>
#include <vector>

namespace gram {

enum class ChannelMode {
    Left,
    Right,
    Mix
};

enum class PaletteMode {
    Spectrum,
    Gray,
    Blue,
    Green,
    Amber,
    Ice
};

enum class FrequencyScale {
    Linear,
    Logarithmic,
    Octave
};

struct WavData {
    int sample_rate = 0;
    int channels = 0;
    int bits_per_sample = 0;
    std::vector<double> samples;
};

struct SpectrogramSettings {
    int fft_size = 2048;
    int hop_size = 512;
    int image_height = 768;
    double min_db = -90.0;
    double max_db = 0.0;
    double min_frequency = 0.0;
    double max_frequency = 0.0;
    FrequencyScale frequency_scale = FrequencyScale::Linear;
    PaletteMode palette_mode = PaletteMode::Spectrum;
};

struct ProfilingStats {
    double last_fft_ms = 0.0;
    double average_fft_ms = 0.0;
    std::uint64_t fft_calls = 0;
};

struct Image {
    int width = 0;
    int height = 0;
    std::vector<std::uint32_t> pixels;
};

class FftPlan {
public:
    explicit FftPlan(int size = 2048);

    [[nodiscard]] int size() const noexcept;
    [[nodiscard]] int bin_count() const noexcept;

    void compute_db(
        const double* samples,
        std::vector<double>& out_db_bins,
        double floor_db,
        ProfilingStats* profiling = nullptr) const;

private:
    int size_ = 0;
    std::vector<std::size_t> bit_reverse_;
    std::vector<double> window_;
    std::vector<std::complex<double>> roots_;
};

[[nodiscard]] bool is_power_of_two(int value);
[[nodiscard]] WavData load_wav(const std::string& path);
[[nodiscard]] std::vector<double> select_channel(const WavData& wav, ChannelMode mode);
[[nodiscard]] std::vector<double> render_column(
    const std::vector<double>& db_bins,
    int sample_rate,
    const SpectrogramSettings& settings);
[[nodiscard]] Image make_image(int width, int height);
void clear_image(Image& image, std::uint32_t color = 0);
void scroll_image_left(Image& image);
void set_image_column(
    Image& image,
    int x,
    const std::vector<double>& row_db,
    const SpectrogramSettings& settings);
[[nodiscard]] Image render_full_spectrogram(
    const std::vector<double>& samples,
    int sample_rate,
    const SpectrogramSettings& settings,
    ProfilingStats* profiling = nullptr);
void write_ppm(const std::string& path, const Image& image);
void write_wav_mono_16(
    const std::string& path,
    int sample_rate,
    const std::vector<std::int16_t>& samples);

} // namespace gram
