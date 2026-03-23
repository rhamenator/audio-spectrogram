#include "spectrogram_core.h"

#include <cstdlib>
#include <exception>
#include <iostream>
#include <stdexcept>
#include <string>
#include <string_view>
#include <vector>

namespace {

struct CliOptions {
    std::string input_path;
    std::string output_path;
    gram::SpectrogramSettings settings;
    gram::ChannelMode channel_mode = gram::ChannelMode::Mix;
    bool show_profile = false;
};

[[noreturn]] void fail(const std::string& message) {
    throw std::runtime_error(message);
}

int parse_positive_int(const std::string& text, const char* option_name) {
    try {
        const int value = std::stoi(text);
        if (value <= 0) {
            fail(std::string(option_name) + " must be greater than zero.");
        }
        return value;
    } catch (const std::exception&) {
        fail(std::string("Invalid integer for ") + option_name + ".");
    }
}

double parse_double_value(const std::string& text, const char* option_name) {
    try {
        return std::stod(text);
    } catch (const std::exception&) {
        fail(std::string("Invalid number for ") + option_name + ".");
    }
}

gram::ChannelMode parse_channel(std::string_view value) {
    if (value == "left") {
        return gram::ChannelMode::Left;
    }
    if (value == "right") {
        return gram::ChannelMode::Right;
    }
    if (value == "mix") {
        return gram::ChannelMode::Mix;
    }
    fail("Invalid --channel value. Expected left, right, or mix.");
}

gram::PaletteMode parse_palette(std::string_view value) {
    if (value == "spectrum") {
        return gram::PaletteMode::Spectrum;
    }
    if (value == "gray") {
        return gram::PaletteMode::Gray;
    }
    if (value == "blue") {
        return gram::PaletteMode::Blue;
    }
    if (value == "green") {
        return gram::PaletteMode::Green;
    }
    if (value == "amber") {
        return gram::PaletteMode::Amber;
    }
    if (value == "ice") {
        return gram::PaletteMode::Ice;
    }
    fail("Invalid --palette value. Expected spectrum, gray, blue, green, amber, or ice.");
}

gram::FrequencyScale parse_scale(std::string_view value) {
    if (value == "linear") {
        return gram::FrequencyScale::Linear;
    }
    if (value == "log" || value == "logarithmic") {
        return gram::FrequencyScale::Logarithmic;
    }
    if (value == "octave" || value == "octaves") {
        return gram::FrequencyScale::Octave;
    }
    fail("Invalid --scale value. Expected linear, log, or octave.");
}

void print_usage() {
    std::cout
        << "Usage:\n"
        << "  gram_repro <input.wav> <output.ppm> [options]\n\n"
        << "Options:\n"
        << "  --fft <n>         FFT size, power of two (default: 2048)\n"
        << "  --hop <n>         Hop size in samples (default: 512)\n"
        << "  --height <n>      Output image height (default: 768)\n"
        << "  --min-db <n>      Minimum displayed dB (default: -90)\n"
        << "  --max-db <n>      Maximum displayed dB (default: 0)\n"
        << "  --max-freq <n>    Upper frequency limit in Hz (default: Nyquist)\n"
        << "  --channel <mode>  left | right | mix (default: mix)\n"
        << "  --palette <name>  spectrum | gray | blue | green | amber | ice\n"
        << "  --scale <mode>    linear | log | octave (default: linear)\n"
        << "  --profile         Print average FFT timing\n"
        << "  --help            Show this help text\n";
}

CliOptions parse_args(int argc, char** argv) {
    if (argc < 2) {
        print_usage();
        std::exit(1);
    }

    std::vector<std::string> args(argv + 1, argv + argc);
    if (args.size() == 1 && args[0] == "--help") {
        print_usage();
        std::exit(0);
    }
    if (args.size() < 2) {
        fail("Expected at least an input and output path.");
    }

    CliOptions options;
    options.input_path = args[0];
    options.output_path = args[1];

    for (std::size_t index = 2; index < args.size(); ++index) {
        const std::string& arg = args[index];

        auto require_value = [&](const char* name) -> const std::string& {
            if (index + 1 >= args.size()) {
                fail(std::string("Missing value for ") + name + ".");
            }
            ++index;
            return args[index];
        };

        if (arg == "--fft") {
            options.settings.fft_size = parse_positive_int(require_value("--fft"), "--fft");
        } else if (arg == "--hop") {
            options.settings.hop_size = parse_positive_int(require_value("--hop"), "--hop");
        } else if (arg == "--height") {
            options.settings.image_height = parse_positive_int(require_value("--height"), "--height");
        } else if (arg == "--min-db") {
            options.settings.min_db = parse_double_value(require_value("--min-db"), "--min-db");
        } else if (arg == "--max-db") {
            options.settings.max_db = parse_double_value(require_value("--max-db"), "--max-db");
        } else if (arg == "--max-freq") {
            options.settings.max_frequency = parse_double_value(require_value("--max-freq"), "--max-freq");
        } else if (arg == "--channel") {
            options.channel_mode = parse_channel(require_value("--channel"));
        } else if (arg == "--palette") {
            options.settings.palette_mode = parse_palette(require_value("--palette"));
        } else if (arg == "--scale") {
            options.settings.frequency_scale = parse_scale(require_value("--scale"));
        } else if (arg == "--log-freq") {
            options.settings.frequency_scale = gram::FrequencyScale::Logarithmic;
        } else if (arg == "--profile") {
            options.show_profile = true;
        } else if (arg == "--help") {
            print_usage();
            std::exit(0);
        } else {
            fail("Unknown argument: " + arg);
        }
    }

    if (!gram::is_power_of_two(options.settings.fft_size)) {
        fail("--fft must be a power of two.");
    }
    if (options.settings.hop_size > options.settings.fft_size) {
        fail("--hop must be less than or equal to --fft.");
    }
    if (options.settings.image_height < 64) {
        fail("--height must be at least 64 pixels.");
    }
    if (options.settings.min_db >= options.settings.max_db) {
        fail("--min-db must be lower than --max-db.");
    }

    return options;
}

} // namespace

int main(int argc, char** argv) {
    try {
        const CliOptions options = parse_args(argc, argv);
        const gram::WavData wav = gram::load_wav(options.input_path);
        const auto selected = gram::select_channel(wav, options.channel_mode);
        gram::ProfilingStats profiling;
        const auto image = gram::render_full_spectrogram(
            selected,
            wav.sample_rate,
            options.settings,
            options.show_profile ? &profiling : nullptr);
        gram::write_ppm(options.output_path, image);

        std::cout
            << "Wrote spectrogram image to " << options.output_path << "\n"
            << "Sample rate: " << wav.sample_rate << " Hz\n"
            << "Channels in source: " << wav.channels << "\n"
            << "FFT size: " << options.settings.fft_size << "\n"
            << "Hop size: " << options.settings.hop_size << "\n";

        if (options.show_profile) {
            std::cout
                << "FFT calls: " << profiling.fft_calls << "\n"
                << "Average FFT time: " << profiling.average_fft_ms << " ms\n";
        }

        return 0;
    } catch (const std::exception& error) {
        std::cerr << "gram_repro: " << error.what() << '\n';
        return 1;
    }
}
