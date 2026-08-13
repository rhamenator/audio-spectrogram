#include "spectrogram_core.h"

#include <cassert>
#include <cmath>
#include <filesystem>
#include <stdexcept>
#include <vector>

int main() {
    using namespace gram;

    assert(is_power_of_two(1));
    assert(is_power_of_two(2048));
    assert(!is_power_of_two(0));
    assert(!is_power_of_two(12));

    bool rejected = false;
    try {
        FftPlan invalid(12);
    } catch (const std::runtime_error&) {
        rejected = true;
    }
    assert(rejected);

    WavData stereo;
    stereo.sample_rate = 48'000;
    stereo.channels = 2;
    stereo.bits_per_sample = 16;
    stereo.samples = {1.0, -1.0, 0.5, 0.25};
    assert((select_channel(stereo, ChannelMode::Left) == std::vector<double>{1.0, 0.5}));
    assert((select_channel(stereo, ChannelMode::Right) == std::vector<double>{-1.0, 0.25}));
    const auto mixed = select_channel(stereo, ChannelMode::Mix);
    assert(std::abs(mixed[0]) < 1.0e-12);
    assert(std::abs(mixed[1] - 0.375) < 1.0e-12);

    auto image = make_image(3, 1);
    image.pixels = {1, 2, 3};
    scroll_image_left(image);
    assert(image.pixels[0] == 2 && image.pixels[1] == 3 && image.pixels[2] == 0);

    SpectrogramSettings settings;
    settings.fft_size = 8;
    settings.hop_size = 4;
    settings.image_height = 4;
    ProfilingStats profiling;
    const std::vector<double> samples(12, 0.25);
    const auto rendered = render_full_spectrogram(samples, 8'000, settings, &profiling);
    assert(rendered.width == 2);
    assert(rendered.height == 4);
    assert(rendered.pixels.size() == 8);
    assert(profiling.fft_calls == 2);

    const auto temp = std::filesystem::temp_directory_path() / "gram-core-test.wav";
    write_wav_mono_16(temp.string(), 8'000, {0, 16'384, -16'384});
    const auto loaded = load_wav(temp.string());
    std::filesystem::remove(temp);
    assert(loaded.sample_rate == 8'000);
    assert(loaded.channels == 1);
    assert(loaded.samples.size() == 3);
    assert(std::abs(loaded.samples[1] - 0.5) < 1.0e-6);

    return 0;
}
