#include "spectrogram_core.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <complex>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <stdexcept>
#include <string_view>

namespace gram {

namespace {

constexpr double kPi = 3.14159265358979323846;

[[noreturn]] void fail(const std::string& message) {
    throw std::runtime_error(message);
}

std::uint16_t read_u16(std::istream& input) {
    unsigned char bytes[2] {};
    input.read(reinterpret_cast<char*>(bytes), 2);
    if (!input) {
        fail("Unexpected end of file while reading 16-bit value.");
    }
    return static_cast<std::uint16_t>(bytes[0] | (bytes[1] << 8));
}

std::uint32_t read_u32(std::istream& input) {
    unsigned char bytes[4] {};
    input.read(reinterpret_cast<char*>(bytes), 4);
    if (!input) {
        fail("Unexpected end of file while reading 32-bit value.");
    }
    return static_cast<std::uint32_t>(
        bytes[0] |
        (bytes[1] << 8) |
        (bytes[2] << 16) |
        (bytes[3] << 24));
}

std::uint32_t pack_bgrx(std::uint8_t r, std::uint8_t g, std::uint8_t b) {
    return static_cast<std::uint32_t>(b) |
        (static_cast<std::uint32_t>(g) << 8) |
        (static_cast<std::uint32_t>(r) << 16);
}

std::uint32_t lerp_color(std::uint32_t left, std::uint32_t right, double t) {
    const auto unpack = [](std::uint32_t color, int shift) -> int {
        return static_cast<int>((color >> shift) & 0xFFu);
    };

    const auto mix = [t](int a, int b) -> std::uint8_t {
        const auto value = static_cast<long long>(std::lround(a + (b - a) * t));
        return static_cast<std::uint8_t>(std::clamp<long long>(value, 0, 255));
    };

    const auto b = mix(unpack(left, 0), unpack(right, 0));
    const auto g = mix(unpack(left, 8), unpack(right, 8));
    const auto r = mix(unpack(left, 16), unpack(right, 16));
    return pack_bgrx(r, g, b);
}

double normalize_level(double db, const SpectrogramSettings& settings) {
    const double clamped = std::clamp(db, settings.min_db, settings.max_db);
    return (clamped - settings.min_db) / (settings.max_db - settings.min_db);
}

std::uint32_t colorize(double level, PaletteMode palette_mode) {
    struct Stop {
        double t;
        std::uint32_t color;
    };

    static const std::vector<Stop> spectrum {
        {0.00, pack_bgrx(0, 0, 0)},
        {0.12, pack_bgrx(88, 0, 128)},
        {0.26, pack_bgrx(75, 0, 130)},
        {0.40, pack_bgrx(0, 70, 200)},
        {0.55, pack_bgrx(0, 170, 220)},
        {0.70, pack_bgrx(0, 200, 90)},
        {0.82, pack_bgrx(255, 220, 0)},
        {0.92, pack_bgrx(255, 120, 0)},
        {1.00, pack_bgrx(255, 0, 0)}
    };
    static const std::vector<Stop> gray {
        {0.00, pack_bgrx(0, 0, 0)},
        {1.00, pack_bgrx(255, 255, 255)}
    };
    static const std::vector<Stop> blue {
        {0.00, pack_bgrx(0, 0, 0)},
        {1.00, pack_bgrx(90, 170, 255)}
    };
    static const std::vector<Stop> green {
        {0.00, pack_bgrx(0, 0, 0)},
        {1.00, pack_bgrx(80, 255, 120)}
    };
    static const std::vector<Stop> amber {
        {0.00, pack_bgrx(0, 0, 0)},
        {1.00, pack_bgrx(255, 190, 40)}
    };
    static const std::vector<Stop> ice {
        {0.00, pack_bgrx(0, 0, 10)},
        {0.25, pack_bgrx(0, 40, 120)},
        {0.60, pack_bgrx(0, 180, 220)},
        {1.00, pack_bgrx(240, 255, 255)}
    };

    const std::vector<Stop>* stops = &spectrum;
    if (palette_mode == PaletteMode::Gray) {
        stops = &gray;
    } else if (palette_mode == PaletteMode::Blue) {
        stops = &blue;
    } else if (palette_mode == PaletteMode::Green) {
        stops = &green;
    } else if (palette_mode == PaletteMode::Amber) {
        stops = &amber;
    } else if (palette_mode == PaletteMode::Ice) {
        stops = &ice;
    }

    if (level <= stops->front().t) {
        return stops->front().color;
    }
    if (level >= stops->back().t) {
        return stops->back().color;
    }

    for (std::size_t index = 1; index < stops->size(); ++index) {
        if (level <= (*stops)[index].t) {
            const auto& left = (*stops)[index - 1];
            const auto& right = (*stops)[index];
            const double local = (level - left.t) / (right.t - left.t);
            return lerp_color(left.color, right.color, local);
        }
    }

    return stops->back().color;
}

} // namespace

bool is_power_of_two(int value) {
    return value > 0 && (value & (value - 1)) == 0;
}

FftPlan::FftPlan(int size)
    : size_(size),
      bit_reverse_(static_cast<std::size_t>(size)),
      window_(static_cast<std::size_t>(size)),
      roots_(static_cast<std::size_t>(size / 2)) {
    if (!is_power_of_two(size_)) {
        fail("FFT size must be a power of two.");
    }

    int levels = 0;
    for (int value = size_; value > 1; value >>= 1) {
        ++levels;
    }

    for (int index = 0; index < size_; ++index) {
        std::size_t reversed = 0;
        for (int bit = 0; bit < levels; ++bit) {
            reversed = (reversed << 1) | static_cast<std::size_t>((index >> bit) & 1);
        }
        bit_reverse_[static_cast<std::size_t>(index)] = reversed;
        window_[static_cast<std::size_t>(index)] =
            0.5 - 0.5 * std::cos((2.0 * kPi * index) / (size_ - 1));
    }

    for (int index = 0; index < size_ / 2; ++index) {
        const double angle = -2.0 * kPi * static_cast<double>(index) / static_cast<double>(size_);
        roots_[static_cast<std::size_t>(index)] =
            std::complex<double>(std::cos(angle), std::sin(angle));
    }
}

int FftPlan::size() const noexcept {
    return size_;
}

int FftPlan::bin_count() const noexcept {
    return size_ / 2 + 1;
}

void FftPlan::compute_db(
    const double* samples,
    std::vector<double>& out_db_bins,
    double floor_db,
    ProfilingStats* profiling) const {
    const auto started = std::chrono::steady_clock::now();

    std::vector<std::complex<double>> working(static_cast<std::size_t>(size_));
    for (int index = 0; index < size_; ++index) {
        working[bit_reverse_[static_cast<std::size_t>(index)]] =
            std::complex<double>(samples[index] * window_[static_cast<std::size_t>(index)], 0.0);
    }

    for (int length = 2; length <= size_; length <<= 1) {
        const int half = length / 2;
        const int step = size_ / length;
        for (int start = 0; start < size_; start += length) {
            for (int offset = 0; offset < half; ++offset) {
                const auto twiddle = roots_[static_cast<std::size_t>(offset * step)];
                const auto even = working[static_cast<std::size_t>(start + offset)];
                const auto odd =
                    working[static_cast<std::size_t>(start + offset + half)] * twiddle;
                working[static_cast<std::size_t>(start + offset)] = even + odd;
                working[static_cast<std::size_t>(start + offset + half)] = even - odd;
            }
        }
    }

    const int bins = bin_count();
    out_db_bins.assign(static_cast<std::size_t>(bins), floor_db);
    for (int bin = 0; bin < bins; ++bin) {
        const double magnitude = std::abs(working[static_cast<std::size_t>(bin)]);
        const double normalized = magnitude / static_cast<double>(size_);
        out_db_bins[static_cast<std::size_t>(bin)] =
            20.0 * std::log10(std::max(normalized, 1.0e-12));
    }

    if (profiling != nullptr) {
        using millis = std::chrono::duration<double, std::milli>;
        profiling->last_fft_ms = millis(std::chrono::steady_clock::now() - started).count();
        ++profiling->fft_calls;
        profiling->average_fft_ms +=
            (profiling->last_fft_ms - profiling->average_fft_ms) /
            static_cast<double>(profiling->fft_calls);
    }
}

WavData load_wav(const std::string& path) {
    std::ifstream input(path, std::ios::binary);
    if (!input) {
        fail("Unable to open input file: " + path);
    }

    char riff[4] {};
    input.read(riff, 4);
    const auto riff_size = read_u32(input);
    (void)riff_size;
    char wave[4] {};
    input.read(wave, 4);

    if (std::string_view(riff, 4) != "RIFF" || std::string_view(wave, 4) != "WAVE") {
        fail("Input file is not a RIFF/WAVE file.");
    }

    bool have_fmt = false;
    bool have_data = false;
    std::uint16_t audio_format = 0;
    std::uint16_t channels = 0;
    std::uint32_t sample_rate = 0;
    std::uint16_t bits_per_sample = 0;
    std::vector<std::uint8_t> raw_data;

    while (input && (!have_fmt || !have_data)) {
        char chunk_id[4] {};
        input.read(chunk_id, 4);
        if (!input) {
            break;
        }

        const auto chunk_size = read_u32(input);
        const std::string_view id(chunk_id, 4);

        if (id == "fmt ") {
            audio_format = read_u16(input);
            channels = read_u16(input);
            sample_rate = read_u32(input);
            const auto byte_rate = read_u32(input);
            (void)byte_rate;
            const auto block_align = read_u16(input);
            (void)block_align;
            bits_per_sample = read_u16(input);

            if (chunk_size > 16) {
                input.seekg(static_cast<std::streamoff>(chunk_size - 16), std::ios::cur);
            }
            have_fmt = true;
        } else if (id == "data") {
            raw_data.resize(chunk_size);
            input.read(reinterpret_cast<char*>(raw_data.data()), static_cast<std::streamsize>(chunk_size));
            if (!input) {
                fail("Unexpected end of file while reading sample data.");
            }
            have_data = true;
        } else {
            input.seekg(static_cast<std::streamoff>(chunk_size), std::ios::cur);
        }

        if ((chunk_size & 1u) != 0u) {
            input.seekg(1, std::ios::cur);
        }
    }

    if (!have_fmt || !have_data) {
        fail("Wave file is missing required fmt/data chunks.");
    }
    if (audio_format != 1) {
        fail("Only uncompressed PCM wave files are supported.");
    }
    if (channels < 1 || channels > 2) {
        fail("Only mono and stereo PCM wave files are supported.");
    }
    if (bits_per_sample != 8 && bits_per_sample != 16) {
        fail("Only 8-bit and 16-bit PCM wave files are supported.");
    }

    WavData wav;
    wav.sample_rate = static_cast<int>(sample_rate);
    wav.channels = static_cast<int>(channels);
    wav.bits_per_sample = static_cast<int>(bits_per_sample);

    const int bytes_per_sample = bits_per_sample / 8;
    const int frame_size = bytes_per_sample * wav.channels;
    if (raw_data.size() % static_cast<std::size_t>(frame_size) != 0) {
        fail("Wave data size is not aligned to sample frames.");
    }

    const auto frame_count = raw_data.size() / static_cast<std::size_t>(frame_size);
    wav.samples.resize(frame_count * static_cast<std::size_t>(wav.channels));

    std::size_t cursor = 0;
    for (std::size_t frame = 0; frame < frame_count; ++frame) {
        for (int channel = 0; channel < wav.channels; ++channel) {
            double sample = 0.0;
            if (bits_per_sample == 8) {
                sample = (static_cast<int>(raw_data[cursor++]) - 128) / 128.0;
            } else {
                const auto value = static_cast<std::int16_t>(
                    raw_data[cursor] | (raw_data[cursor + 1] << 8));
                cursor += 2;
                sample = std::clamp(static_cast<double>(value) / 32768.0, -1.0, 1.0);
            }
            wav.samples[frame * static_cast<std::size_t>(wav.channels) +
                static_cast<std::size_t>(channel)] = sample;
        }
    }

    return wav;
}

std::vector<double> select_channel(const WavData& wav, ChannelMode mode) {
    const auto frame_count = wav.samples.size() / static_cast<std::size_t>(wav.channels);
    std::vector<double> selected(frame_count);

    for (std::size_t frame = 0; frame < frame_count; ++frame) {
        const double left = wav.samples[frame * static_cast<std::size_t>(wav.channels)];
        const double right = wav.channels == 2
            ? wav.samples[frame * static_cast<std::size_t>(wav.channels) + 1]
            : left;

        if (mode == ChannelMode::Left) {
            selected[frame] = left;
        } else if (mode == ChannelMode::Right) {
            selected[frame] = right;
        } else {
            selected[frame] = 0.5 * (left + right);
        }
    }

    return selected;
}

std::vector<double> render_column(
    const std::vector<double>& db_bins,
    int sample_rate,
    const SpectrogramSettings& settings) {
    const double nyquist = sample_rate / 2.0;
    const double min_frequency = std::clamp(settings.min_frequency, 0.0, nyquist);
    const double max_frequency = settings.max_frequency > 0.0
        ? std::min(settings.max_frequency, nyquist)
        : nyquist;
    const int bins = static_cast<int>(db_bins.size());

    std::vector<double> column(static_cast<std::size_t>(settings.image_height), settings.min_db);
    for (int y = 0; y < settings.image_height; ++y) {
        double position = 0.0;
        if (settings.image_height > 1) {
            position = static_cast<double>(y) / static_cast<double>(settings.image_height - 1);
        }

        double target_frequency = 0.0;
        if (settings.frequency_scale == FrequencyScale::Linear) {
            target_frequency = min_frequency + position * (max_frequency - min_frequency);
        } else if (settings.frequency_scale == FrequencyScale::Logarithmic) {
            const double min_log_freq = std::max(min_frequency, 1.0);
            const double span =
                std::log10(std::max(max_frequency, min_log_freq)) - std::log10(min_log_freq);
            target_frequency = std::pow(10.0, std::log10(min_log_freq) + position * span);
        } else {
            const double min_octave_freq = std::max(min_frequency, 1.0);
            const double top = std::max(max_frequency, min_octave_freq);
            const double octave_span = std::log2(top / min_octave_freq);
            target_frequency = min_octave_freq * std::pow(2.0, position * octave_span);
        }

        const auto rounded = static_cast<long long>(
            std::lround((target_frequency / nyquist) * static_cast<double>(bins - 1)));
        const int bin = static_cast<int>(
            std::clamp<long long>(rounded, 0, static_cast<long long>(bins - 1)));
        column[static_cast<std::size_t>(settings.image_height - 1 - y)] =
            db_bins[static_cast<std::size_t>(bin)];
    }

    return column;
}

Image make_image(int width, int height) {
    Image image;
    image.width = std::max(width, 1);
    image.height = std::max(height, 1);
    image.pixels.assign(
        static_cast<std::size_t>(image.width * image.height),
        pack_bgrx(0, 0, 0));
    return image;
}

void clear_image(Image& image, std::uint32_t color) {
    std::fill(image.pixels.begin(), image.pixels.end(), color);
}

void scroll_image_left(Image& image) {
    if (image.width <= 1 || image.height <= 0) {
        return;
    }

    for (int y = 0; y < image.height; ++y) {
        auto* row = image.pixels.data() + static_cast<std::size_t>(y * image.width);
        std::move(row + 1, row + image.width, row);
        row[image.width - 1] = pack_bgrx(0, 0, 0);
    }
}

void set_image_column(
    Image& image,
    int x,
    const std::vector<double>& row_db,
    const SpectrogramSettings& settings) {
    if (x < 0 || x >= image.width || image.height <= 0) {
        return;
    }

    const int rows = std::min(image.height, static_cast<int>(row_db.size()));
    for (int y = 0; y < rows; ++y) {
        image.pixels[static_cast<std::size_t>(y * image.width + x)] =
            colorize(normalize_level(row_db[static_cast<std::size_t>(y)], settings), settings.palette_mode);
    }
}

Image render_full_spectrogram(
    const std::vector<double>& samples,
    int sample_rate,
    const SpectrogramSettings& settings,
    ProfilingStats* profiling) {
    if (samples.size() < static_cast<std::size_t>(settings.fft_size)) {
        fail("Input is shorter than the selected FFT size.");
    }

    FftPlan plan(settings.fft_size);
    const auto frame_count =
        1 + (samples.size() - static_cast<std::size_t>(settings.fft_size)) /
        static_cast<std::size_t>(settings.hop_size);

    Image image = make_image(static_cast<int>(frame_count), settings.image_height);
    std::vector<double> db_bins;

    for (std::size_t frame = 0; frame < frame_count; ++frame) {
        const auto* window = samples.data() + frame * static_cast<std::size_t>(settings.hop_size);
        plan.compute_db(window, db_bins, settings.min_db, profiling);
        set_image_column(
            image,
            static_cast<int>(frame),
            render_column(db_bins, sample_rate, settings),
            settings);
    }

    return image;
}

void write_ppm(const std::string& path, const Image& image) {
    std::ofstream output(path, std::ios::binary);
    if (!output) {
        fail("Unable to open output file: " + path);
    }

    output << "P6\n" << image.width << ' ' << image.height << "\n255\n";
    for (const auto color : image.pixels) {
        const char rgb[3] {
            static_cast<char>((color >> 16) & 0xFFu),
            static_cast<char>((color >> 8) & 0xFFu),
            static_cast<char>(color & 0xFFu)
        };
        output.write(rgb, 3);
    }

    if (!output) {
        fail("Failed while writing output image.");
    }
}

void write_wav_mono_16(
    const std::string& path,
    int sample_rate,
    const std::vector<std::int16_t>& samples) {
    std::ofstream output(path, std::ios::binary);
    if (!output) {
        fail("Unable to open output wave file: " + path);
    }

    const std::uint32_t data_size =
        static_cast<std::uint32_t>(samples.size() * sizeof(std::int16_t));
    const std::uint32_t chunk_size = 36u + data_size;
    const std::uint16_t channels = 1;
    const std::uint16_t bits_per_sample = 16;
    const std::uint16_t block_align = channels * (bits_per_sample / 8);
    const std::uint32_t byte_rate =
        static_cast<std::uint32_t>(sample_rate) * static_cast<std::uint32_t>(block_align);

    output.write("RIFF", 4);
    output.write(reinterpret_cast<const char*>(&chunk_size), sizeof(chunk_size));
    output.write("WAVE", 4);
    output.write("fmt ", 4);

    const std::uint32_t fmt_size = 16;
    const std::uint16_t format = 1;
    output.write(reinterpret_cast<const char*>(&fmt_size), sizeof(fmt_size));
    output.write(reinterpret_cast<const char*>(&format), sizeof(format));
    output.write(reinterpret_cast<const char*>(&channels), sizeof(channels));
    output.write(reinterpret_cast<const char*>(&sample_rate), sizeof(sample_rate));
    output.write(reinterpret_cast<const char*>(&byte_rate), sizeof(byte_rate));
    output.write(reinterpret_cast<const char*>(&block_align), sizeof(block_align));
    output.write(reinterpret_cast<const char*>(&bits_per_sample), sizeof(bits_per_sample));

    output.write("data", 4);
    output.write(reinterpret_cast<const char*>(&data_size), sizeof(data_size));
    output.write(reinterpret_cast<const char*>(samples.data()), static_cast<std::streamsize>(data_size));

    if (!output) {
        fail("Failed while writing output wave file.");
    }
}

} // namespace gram
