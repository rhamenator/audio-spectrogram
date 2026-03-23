#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include "spectrogram_core.h"

#include <windows.h>
#include <windowsx.h>
#include <commdlg.h>
#include <commctrl.h>
#include <mmsystem.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <mmreg.h>
#include <wincodec.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Media.Control.h>

#include <algorithm>
#include <cmath>
#include <cstdlib>
#include <cstdint>
#include <cstring>
#include <exception>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "windowscodecs.lib")

namespace {

constexpr int kControlsHeight = 186;
constexpr UINT_PTR kPollTimer = 1;
constexpr UINT kPollIntervalMs = 30;
constexpr int kButtonWidth = 92;
constexpr int kButtonHeight = 28;
constexpr double kAnalysisFloorDb = -180.0;
constexpr double kPeakHoldDecayDbPerFrame = 0.18;
constexpr double kAverageBlendAlpha = 0.16;

enum ControlId : int {
    StartButton = 1001,
    StopButton = 1002,
    PlayButton = 1003,
    SaveButton = 1004,
    ApplyButton = 1005,
    FreezeButton = 1006,
    LoopButton = 1007,
    MarkInButton = 1008,
    MarkOutButton = 1009,
    SavePngButton = 1010,
    LoadPresetButton = 1011,
    SavePresetButton = 1012,
    SaveCsvButton = 1013,
    SampleRateCombo = 1101,
    FftCombo = 1102,
    PaletteCombo = 1103,
    ScaleCombo = 1104,
    MaxFreqCombo = 1105,
    DisplayCombo = 1106,
    GridCheck = 1107,
    ChannelCombo = 1108,
    AmplitudeCombo = 1109,
    MaxDbSlider = 1110,
    RangeSlider = 1111,
    MonoChannelCombo = 1112,
    PeakHoldCheck = 1113,
    AverageCheck = 1114,
    HistoryCombo = 1115,
    ScrollSpeedCombo = 1116,
    MinFreqCombo = 1117,
    MusicGridCheck = 1118,
    BeatMarksCheck = 1119,
    InstrumentCombo = 1120,
    TunerStartButton = 1121,
    TunerStopButton = 1122,
};

enum class DisplayMode {
    Spectrogram,
    Spectrum,
    Tuner,
    Vectorscope,
    Room
};

enum class ChannelDisplayMode {
    Auto,
    Mono,
    Stereo
};

enum class SpectrumAmplitudeMode {
    Decibels,
    Linear
};

enum class MonoChannelMode {
    Mix,
    Left,
    Right
};

struct SpectrumPeak {
    double frequency = 0.0;
    double level_db = -180.0;
    int midi_note = 0;
    double cents = 0.0;
};

struct MusicStatusFields {
    std::wstring pitch;
    std::wstring cents;
    std::wstring chord;
};

template <typename T>
void safe_release(T*& value) {
    if (value != nullptr) {
        value->Release();
        value = nullptr;
    }
}

std::wstring widen(const char* text) {
    if (text == nullptr) {
        return L"";
    }

    const int count = MultiByteToWideChar(CP_UTF8, 0, text, -1, nullptr, 0);
    if (count <= 1) {
        return std::wstring(text, text + std::strlen(text));
    }

    std::wstring wide(static_cast<std::size_t>(count - 1), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text, -1, wide.data(), count);
    return wide;
}

std::string narrow_system_path(const wchar_t* text) {
    if (text == nullptr) {
        return {};
    }

    const int count = WideCharToMultiByte(CP_ACP, 0, text, -1, nullptr, 0, nullptr, nullptr);
    if (count <= 1) {
        return {};
    }

    std::string narrow(static_cast<std::size_t>(count - 1), '\0');
    WideCharToMultiByte(CP_ACP, 0, text, -1, narrow.data(), count, nullptr, nullptr);
    return narrow;
}

bool is_float_format(const WAVEFORMATEX* format) {
    if (format == nullptr) {
        return false;
    }
    if (format->wFormatTag == WAVE_FORMAT_IEEE_FLOAT) {
        return true;
    }
    if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        const auto* extensible = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(format);
        return extensible->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
    }
    return false;
}

bool is_pcm_format(const WAVEFORMATEX* format) {
    if (format == nullptr) {
        return false;
    }
    if (format->wFormatTag == WAVE_FORMAT_PCM) {
        return true;
    }
    if (format->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        const auto* extensible = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(format);
        return extensible->SubFormat == KSDATAFORMAT_SUBTYPE_PCM;
    }
    return false;
}

int parse_combo_int(HWND combo) {
    wchar_t buffer[64] {};
    GetWindowTextW(combo, buffer, 63);
    return _wtoi(buffer);
}

double parse_combo_double(HWND combo) {
    wchar_t buffer[64] {};
    GetWindowTextW(combo, buffer, 63);
    wchar_t* end = nullptr;
    const double value = wcstod(buffer, &end);
    if (end == buffer) {
        return 0.0;
    }
    return value;
}

double parse_max_freq(HWND combo) {
    const LRESULT selection = SendMessageW(combo, CB_GETCURSEL, 0, 0);
    if (selection != CB_ERR) {
        wchar_t selected[64] {};
        SendMessageW(combo, CB_GETLBTEXT, static_cast<WPARAM>(selection), reinterpret_cast<LPARAM>(selected));
        if (wcscmp(selected, L"Auto") == 0) {
            return 0.0;
        }
        return static_cast<double>(_wtoi(selected));
    }

    wchar_t buffer[64] {};
    GetWindowTextW(combo, buffer, 63);
    if (wcscmp(buffer, L"Auto") == 0) {
        return 0.0;
    }
    return static_cast<double>(_wtoi(buffer));
}

void set_combo_to_text(HWND combo, const wchar_t* text) {
    const LRESULT count = SendMessageW(combo, CB_GETCOUNT, 0, 0);
    for (LRESULT index = 0; index < count; ++index) {
        wchar_t buffer[64] {};
        SendMessageW(combo, CB_GETLBTEXT, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(buffer));
        if (wcscmp(buffer, text) == 0) {
            SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(index), 0);
            return;
        }
    }
    SetWindowTextW(combo, text);
}

void update_average_bins(std::vector<double>& average_bins, const std::vector<double>& current_bins) {
    if (current_bins.empty()) {
        return;
    }
    if (average_bins.size() != current_bins.size()) {
        average_bins = current_bins;
        return;
    }
    for (std::size_t index = 0; index < current_bins.size(); ++index) {
        average_bins[index] =
            average_bins[index] * (1.0 - kAverageBlendAlpha) +
            current_bins[index] * kAverageBlendAlpha;
    }
}

void update_peak_hold_bins(std::vector<double>& peak_hold_bins, const std::vector<double>& current_bins) {
    if (current_bins.empty()) {
        return;
    }
    if (peak_hold_bins.size() != current_bins.size()) {
        peak_hold_bins = current_bins;
        return;
    }
    for (std::size_t index = 0; index < current_bins.size(); ++index) {
        peak_hold_bins[index] =
            (std::max)(current_bins[index], peak_hold_bins[index] - kPeakHoldDecayDbPerFrame);
    }
}

std::wstring format_db_text(double value) {
    std::wostringstream stream;
    stream << std::fixed << std::setprecision(1) << value << L" dB";
    return stream.str();
}

std::wstring trim_for_status(const std::wstring& text, std::size_t max_chars) {
    if (text.size() <= max_chars) {
        return text;
    }
    if (max_chars <= 1) {
        return text.substr(0, max_chars);
    }
    return text.substr(0, max_chars - 1) + L"...";
}

std::wstring fixed_status_text(const std::wstring& text, std::size_t width) {
    const std::wstring trimmed = trim_for_status(text, width);
    if (trimmed.size() >= width) {
        return trimmed;
    }
    return trimmed + std::wstring(width - trimmed.size(), L' ');
}

std::wstring module_directory() {
    wchar_t path[MAX_PATH] {};
    GetModuleFileNameW(nullptr, path, MAX_PATH);
    return std::filesystem::path(path).parent_path().wstring();
}

std::wstring live_log_path() {
    return module_directory() + L"\\gram_live.log";
}

std::wstring format_hresult(HRESULT value) {
    wchar_t* buffer = nullptr;
    const DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    const DWORD count = FormatMessageW(
        flags,
        nullptr,
        static_cast<DWORD>(value),
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        reinterpret_cast<LPWSTR>(&buffer),
        0,
        nullptr);
    std::wstring text;
    if (count > 0 && buffer != nullptr) {
        text.assign(buffer, buffer + count);
        while (!text.empty() && (text.back() == L'\r' || text.back() == L'\n' || text.back() == L' ' || text.back() == L'\t')) {
            text.pop_back();
        }
        LocalFree(buffer);
    }
    std::wostringstream stream;
    stream << L"0x" << std::hex << std::uppercase << static_cast<unsigned long>(value);
    if (!text.empty()) {
        stream << L" (" << text << L")";
    }
    return stream.str();
}

void append_live_log(const std::wstring& context, const std::wstring& message) {
    std::wofstream output(live_log_path(), std::ios::app);
    if (!output) {
        return;
    }
    SYSTEMTIME now {};
    GetLocalTime(&now);
    output << std::setfill(L'0')
           << std::setw(4) << now.wYear << L'-'
           << std::setw(2) << now.wMonth << L'-'
           << std::setw(2) << now.wDay << L' '
           << std::setw(2) << now.wHour << L':'
           << std::setw(2) << now.wMinute << L':'
           << std::setw(2) << now.wSecond << L'.'
           << std::setw(3) << now.wMilliseconds
           << L" [" << context << L"] "
           << message << L"\n";
}

void show_logged_error(HWND hwnd, const std::wstring& context, const std::wstring& message) {
    append_live_log(context, message);
    std::wstring dialog = message + L"\n\nSee " + live_log_path();
    MessageBoxW(hwnd, dialog.c_str(), L"gram_live", MB_ICONERROR);
}

void show_logged_hresult(HWND hwnd, const std::wstring& context, const std::wstring& message, HRESULT value) {
    show_logged_error(hwnd, context, message + L"\nHRESULT: " + format_hresult(value));
}

double frequency_from_midi(int midi_note) {
    return 440.0 * std::pow(2.0, (static_cast<double>(midi_note) - 69.0) / 12.0);
}

int midi_from_frequency(double frequency) {
    if (frequency <= 0.0) {
        return 69;
    }
    return static_cast<int>(std::lround(69.0 + 12.0 * std::log2(frequency / 440.0)));
}

std::wstring note_name_from_midi(int midi_note) {
    static const wchar_t* names[] {
        L"C", L"C#", L"D", L"D#", L"E", L"F", L"F#", L"G", L"G#", L"A", L"A#", L"B"
    };
    const int index = ((midi_note % 12) + 12) % 12;
    const int octave = midi_note / 12 - 1;
    std::wostringstream stream;
    stream << names[index] << octave;
    return stream.str();
}

std::wstring describe_note_from_frequency(double frequency, double* cents_out = nullptr) {
    if (frequency <= 0.0 || !std::isfinite(frequency)) {
        if (cents_out != nullptr) {
            *cents_out = 0.0;
        }
        return L"--";
    }
    const int midi = midi_from_frequency(frequency);
    const double reference = frequency_from_midi(midi);
    const double cents = 1200.0 * std::log2(frequency / reference);
    if (cents_out != nullptr) {
        *cents_out = cents;
    }
    return note_name_from_midi(midi);
}

std::wstring format_signed_cents(double cents) {
    std::wostringstream stream;
    stream << std::showpos << std::fixed << std::setprecision(1) << cents << L" cents";
    return stream.str();
}

std::wstring pitch_class_name(int midi_note) {
    static const wchar_t* names[] {
        L"C", L"C#", L"D", L"D#", L"E", L"F", L"F#", L"G", L"G#", L"A", L"A#", L"B"
    };
    return names[((midi_note % 12) + 12) % 12];
}

std::wstring chord_hint_from_midis(const std::vector<int>& midis) {
    if (midis.size() < 2) {
        return L"";
    }
    bool present[12] {};
    for (const int midi : midis) {
        present[((midi % 12) + 12) % 12] = true;
    }
    for (int root = 0; root < 12; ++root) {
        if (!present[root]) {
            continue;
        }
        const bool minor = present[(root + 3) % 12] && present[(root + 7) % 12];
        const bool major = present[(root + 4) % 12] && present[(root + 7) % 12];
        const bool power = present[(root + 7) % 12];
        if (major) {
            return pitch_class_name(root) + L" major";
        }
        if (minor) {
            return pitch_class_name(root) + L" minor";
        }
        if (power) {
            return pitch_class_name(root) + L"5";
        }
    }
    return L"";
}

LONG WINAPI unhandled_exception_filter(EXCEPTION_POINTERS* exception) {
    std::wostringstream message;
    message << L"Unhandled exception";
    if (exception != nullptr && exception->ExceptionRecord != nullptr) {
        message << L" 0x" << std::hex << std::uppercase
                << exception->ExceptionRecord->ExceptionCode
                << L" at 0x" << reinterpret_cast<std::uintptr_t>(exception->ExceptionRecord->ExceptionAddress);
    }
    append_live_log(L"unhandled_exception", message.str());
    std::wstring dialog = message.str() + L"\n\nSee " + live_log_path();
    MessageBoxW(nullptr, dialog.c_str(), L"gram_live", MB_ICONERROR);
    return EXCEPTION_EXECUTE_HANDLER;
}

[[noreturn]] void terminate_handler() {
    append_live_log(L"terminate", L"gram_live terminated unexpectedly.");
    std::wstring dialog = L"gram_live terminated unexpectedly.\n\nSee " + live_log_path();
    MessageBoxW(nullptr, dialog.c_str(), L"gram_live", MB_ICONERROR);
    std::abort();
}

class LiveApp;

void apply_settings_with_diagnostics(LiveApp* app);

class LiveApp {
public:
    ~LiveApp();
    bool create(HINSTANCE instance, int show);

    friend void apply_settings_with_diagnostics(LiveApp* app);

private:
    static LRESULT CALLBACK window_proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam);
    LRESULT handle_message(UINT message, WPARAM wparam, LPARAM lparam);

    void create_controls();
    void populate_controls();
    void layout_controls();
    void resize_image();
    [[nodiscard]] int status_bar_height() const;
    RECT graph_rect() const;
    RECT plot_rect() const;
    void sync_settings_from_controls();
    void rebuild_image_from_history();
    void refresh_auto_levels(bool force);

    void start_capture();
    void stop_capture();
    void consume_loopback_packets();
    void append_capture_chunk(const BYTE* data, UINT32 frames, bool silent);
    void append_analysis_sample(double mix_sample, double left_sample, double right_sample);
    void process_available_analysis_frames();
    void trim_analysis_buffer();

    void start_playback();
    void start_loop_playback();
    void stop_playback();
    void save_capture();
    void save_png();
    void save_preset();
    void load_preset();
    void save_note_csv();
    [[nodiscard]] std::vector<std::wstring> active_instrument_notes() const;

    void update_status();
    void update_level_labels();
    void mark_settings_dirty();
    void apply_live_settings_change(int control_id);
    void update_channel_controls();
    void update_transport_controls();
    void update_view_controls();
    void update_now_playing();
    void update_cursor_readout(POINT point, bool force_clear);
    void capture_freeze_snapshot();
    [[nodiscard]] double base_history_seconds() const;
    [[nodiscard]] double effective_scroll_multiplier() const;
    [[nodiscard]] double visible_history_seconds() const;
    [[nodiscard]] std::size_t marker_sample_from_cursor() const;
    [[nodiscard]] double normalized_for_frequency(double frequency) const;
    [[nodiscard]] const std::vector<double>& selected_bins(bool frozen = false) const;
    [[nodiscard]] std::vector<SpectrumPeak> collect_top_peaks(const std::vector<double>& bins, int count) const;
    [[nodiscard]] double current_stereo_correlation() const;
    [[nodiscard]] MusicStatusFields current_music_status() const;
    [[nodiscard]] std::wstring current_music_summary() const;
    void draw_spectrogram_markers(HDC dc, const RECT& plot);
    void draw_note_guides(HDC dc, const RECT& plot, bool horizontal_axis);
    void draw_peak_labels(HDC dc, const RECT& plot, const std::vector<SpectrumPeak>& peaks);
    void draw_tuner(HDC dc, const RECT& outer, const RECT& plot);
    void draw_vectorscope(HDC dc, const RECT& outer, const RECT& plot);
    void draw_room(HDC dc, const RECT& outer, const RECT& plot);
    void update_tuner_controls();
    [[nodiscard]] bool control_requires_apply(int control_id) const;
    void paint(HDC dc);
    void clear_for_new_capture();
    double frequency_for_position(double normalized) const;
    [[nodiscard]] double max_display_frequency() const;
    [[nodiscard]] bool use_stereo_display() const;
    [[nodiscard]] int spectrogram_pane_count() const;
    [[nodiscard]] RECT spectrogram_pane_rect(const RECT& plot, int index) const;
    [[nodiscard]] double spectrum_level_at_position(const std::vector<double>& db_bins, double normalized) const;
    [[nodiscard]] double normalized_amplitude(double db_value) const;
    [[nodiscard]] std::wstring format_amplitude_label(double normalized) const;
    [[nodiscard]] std::wstring format_frequency_label(double frequency) const;
    void draw_spectrogram_axes(HDC dc, const RECT& outer, const RECT& plot, const wchar_t* title);
    void draw_spectrum_axes(HDC dc, const RECT& outer, const RECT& plot);
    void draw_spectrum_trace(HDC dc, const RECT& plot, const std::vector<double>& db_bins, COLORREF color, int line_width, int style);
    void draw_spectrum(HDC dc, const RECT& outer, const RECT& plot);
    void draw_spectrogram_image(HDC dc, const RECT& plot, const gram::Image& image);

    HWND hwnd_ = nullptr;
    HWND start_button_ = nullptr;
    HWND stop_button_ = nullptr;
    HWND play_button_ = nullptr;
    HWND save_button_ = nullptr;
    HWND apply_button_ = nullptr;
    HWND freeze_button_ = nullptr;
    HWND loop_button_ = nullptr;
    HWND mark_in_button_ = nullptr;
    HWND mark_out_button_ = nullptr;
    HWND save_png_button_ = nullptr;
    HWND load_preset_button_ = nullptr;
    HWND save_preset_button_ = nullptr;
    HWND save_csv_button_ = nullptr;
    HWND status_bar_ = nullptr;
    HWND max_db_label_ = nullptr;
    HWND range_label_ = nullptr;
    HWND max_db_slider_ = nullptr;
    HWND range_slider_ = nullptr;
    HWND sample_rate_label_ = nullptr;
    HWND fft_label_ = nullptr;
    HWND palette_label_ = nullptr;
    HWND scale_label_ = nullptr;
    HWND min_freq_label_ = nullptr;
    HWND max_freq_label_ = nullptr;
    HWND channel_label_ = nullptr;
    HWND amplitude_label_ = nullptr;
    HWND mono_channel_label_ = nullptr;
    HWND sample_rate_combo_ = nullptr;
    HWND fft_combo_ = nullptr;
    HWND palette_combo_ = nullptr;
    HWND scale_combo_ = nullptr;
    HWND min_freq_combo_ = nullptr;
    HWND max_freq_combo_ = nullptr;
    HWND display_label_ = nullptr;
    HWND display_combo_ = nullptr;
    HWND grid_check_ = nullptr;
    HWND channel_combo_ = nullptr;
    HWND amplitude_combo_ = nullptr;
    HWND mono_channel_combo_ = nullptr;
    HWND peak_hold_check_ = nullptr;
    HWND average_check_ = nullptr;
    HWND history_label_ = nullptr;
    HWND history_combo_ = nullptr;
    HWND scroll_speed_label_ = nullptr;
    HWND scroll_speed_combo_ = nullptr;
    HWND music_grid_check_ = nullptr;
    HWND beat_marks_check_ = nullptr;
    HWND instrument_label_ = nullptr;
    HWND instrument_combo_ = nullptr;
    HWND tuner_start_button_ = nullptr;
    HWND tuner_stop_button_ = nullptr;

    bool com_initialized_ = false;
    IMMDeviceEnumerator* enumerator_ = nullptr;
    IMMDevice* device_ = nullptr;
    IAudioClient* audio_client_ = nullptr;
    IAudioCaptureClient* capture_client_ = nullptr;
    WAVEFORMATEX* mix_format_ = nullptr;
    winrt::Windows::Media::Control::GlobalSystemMediaTransportControlsSessionManager media_session_manager_ { nullptr };

    gram::SpectrogramSettings settings_ {};
    int analysis_sample_rate_ = 44100;
    int capture_sample_rate_ = 0;
    int capture_channels_ = 0;
    int last_stream_sample_rate_ = 44100;
    int last_stream_channels_ = 1;
    gram::FftPlan fft_plan_ {2048};
    gram::ProfilingStats profiling_ {};
    gram::Image image_ = gram::make_image(640, 360);
    gram::Image left_image_ = gram::make_image(640, 360);
    gram::Image right_image_ = gram::make_image(640, 360);
    gram::Image frozen_image_ = gram::make_image(640, 360);
    gram::Image frozen_left_image_ = gram::make_image(640, 360);
    gram::Image frozen_right_image_ = gram::make_image(640, 360);

    HWAVEOUT wave_out_ = nullptr;
    WAVEHDR output_header_ {};
    std::vector<std::int16_t> playback_samples_;
    bool playing_ = false;
    bool loop_playback_ = false;

    bool capturing_ = false;
    std::vector<std::int16_t> captured_samples_;
    std::vector<std::int16_t> captured_left_samples_;
    std::vector<std::int16_t> captured_right_samples_;
    std::vector<double> analysis_buffer_;
    std::vector<double> left_analysis_buffer_;
    std::vector<double> right_analysis_buffer_;
    std::vector<double> resample_buffer_;
    std::vector<double> left_resample_buffer_;
    std::vector<double> right_resample_buffer_;
    std::vector<double> db_bins_;
    std::vector<double> db_bins_left_;
    std::vector<double> db_bins_right_;
    std::vector<double> peak_hold_bins_;
    std::vector<double> peak_hold_bins_left_;
    std::vector<double> peak_hold_bins_right_;
    std::vector<double> average_bins_;
    std::vector<double> average_bins_left_;
    std::vector<double> average_bins_right_;
    std::vector<double> frozen_db_bins_;
    std::vector<double> frozen_db_bins_left_;
    std::vector<double> frozen_db_bins_right_;
    std::vector<double> frozen_peak_hold_bins_;
    std::vector<double> frozen_peak_hold_bins_left_;
    std::vector<double> frozen_peak_hold_bins_right_;
    std::vector<double> frozen_average_bins_;
    std::vector<double> frozen_average_bins_left_;
    std::vector<double> frozen_average_bins_right_;
    std::size_t analysis_cursor_ = 0;
    double resample_position_ = 0.0;
    DisplayMode display_mode_ = DisplayMode::Spectrogram;
    ChannelDisplayMode channel_display_mode_ = ChannelDisplayMode::Auto;
    SpectrumAmplitudeMode spectrum_amplitude_mode_ = SpectrumAmplitudeMode::Decibels;
    MonoChannelMode mono_channel_mode_ = MonoChannelMode::Mix;
    bool show_grid_ = true;
    bool show_note_guides_ = false;
    bool show_beat_marks_ = false;
    bool show_peak_hold_ = true;
    bool show_average_ = true;
    bool freeze_display_ = false;
    bool graph_dirty_ = true;
    bool settings_dirty_ = false;
    bool auto_levels_ready_ = false;
    bool tracking_mouse_ = false;
    bool cursor_sample_valid_ = false;
    ULONGLONG next_status_tick_ = 0;
    ULONGLONG next_metadata_tick_ = 0;
    ULONGLONG next_auto_level_tick_ = 0;
    double sensitivity_db_ = 90.0;
    double history_seconds_ = 12.0;
    double scroll_speed_ = 1.0;
    double scroll_credit_ = 0.0;
    std::size_t marker_in_sample_ = 0;
    std::size_t marker_out_sample_ = 0;
    bool marker_in_set_ = false;
    bool marker_out_set_ = false;
    std::size_t cursor_sample_index_ = 0;
    std::vector<std::size_t> transient_markers_;
    double last_frame_rms_ = 0.0;
    double stereo_correlation_ = 0.0;
    bool reference_available_ = false;
    std::wstring instrument_mode_ = L"chromatic";

    std::wstring status_text_ = L"Idle";
    std::wstring now_playing_text_;
    std::wstring cursor_text_ = L"Cursor: off";
    std::vector<int> available_sample_rates_ { 16000, 22050, 24000, 32000, 44100, 48000, 88200, 96000 };
};

void apply_settings_with_diagnostics(LiveApp* app) {
    try {
        auto text_of = [](HWND control) {
            wchar_t buffer[128] {};
            GetWindowTextW(control, buffer, 127);
            return std::wstring(buffer);
        };
        std::wostringstream attempt;
        attempt << L"sample_rate=" << text_of(app->sample_rate_combo_)
                << L", fft=" << text_of(app->fft_combo_)
                << L", scale=" << text_of(app->scale_combo_)
                << L", min_freq=" << text_of(app->min_freq_combo_)
                << L", max_freq=" << text_of(app->max_freq_combo_)
                << L", display=" << text_of(app->display_combo_);
        append_live_log(L"apply.begin", attempt.str());
        app->sync_settings_from_controls();
        append_live_log(L"apply.ok", attempt.str());
    } catch (const std::exception& error) {
        const auto error_message = widen(error.what());
        show_logged_error(app->hwnd_, L"apply.failed", error_message);
    }
}

LiveApp::~LiveApp() {
    stop_playback();
    stop_capture();
    if (com_initialized_) {
        CoUninitialize();
    }
}

bool LiveApp::create(HINSTANCE instance, int show) {
    SetUnhandledExceptionFilter(&unhandled_exception_filter);
    std::set_terminate(&terminate_handler);
    settings_.fft_size = 2048;
    settings_.hop_size = 512;
    settings_.image_height = 360;
    settings_.min_db = -90.0;
    settings_.max_db = 0.0;
    settings_.max_frequency = 12000.0;
    settings_.palette_mode = gram::PaletteMode::Spectrum;

    if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
        return false;
    }
    com_initialized_ = true;
    winrt::init_apartment(winrt::apartment_type::single_threaded);

    INITCOMMONCONTROLSEX controls {};
    controls.dwSize = sizeof(controls);
    controls.dwICC = ICC_BAR_CLASSES | ICC_WIN95_CLASSES;
    InitCommonControlsEx(&controls);

    const wchar_t* class_name = L"GramLiveWindow";
    WNDCLASSEXW window_class {};
    window_class.cbSize = sizeof(window_class);
    window_class.lpfnWndProc = &LiveApp::window_proc;
    window_class.hInstance = instance;
    window_class.lpszClassName = class_name;
    window_class.hCursor = LoadCursor(nullptr, IDC_ARROW);
    window_class.hbrBackground = nullptr;

    if (RegisterClassExW(&window_class) == 0) {
        return false;
    }

    hwnd_ = CreateWindowExW(
        0,
        class_name,
        L"gram_next - Live Music Analyzer",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        1260,
        760,
        nullptr,
        nullptr,
        instance,
        this);

    if (hwnd_ == nullptr) {
        return false;
    }

    ShowWindow(hwnd_, show);
    UpdateWindow(hwnd_);
    return true;
}

LRESULT CALLBACK LiveApp::window_proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
    LiveApp* app = nullptr;
    if (message == WM_NCCREATE) {
        auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
        app = static_cast<LiveApp*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(app));
        app->hwnd_ = hwnd;
    } else {
        app = reinterpret_cast<LiveApp*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (app != nullptr) {
        try {
            return app->handle_message(message, wparam, lparam);
        } catch (const winrt::hresult_error& error) {
            const std::wstring text = error.message().c_str();
            show_logged_error(hwnd, L"window_proc.hresult", text);
            PostQuitMessage(1);
            return 0;
        } catch (const std::exception& error) {
            const std::wstring text = widen(error.what());
            show_logged_error(hwnd, L"window_proc.exception", text);
            PostQuitMessage(1);
            return 0;
        } catch (...) {
            show_logged_error(hwnd, L"window_proc.unknown", L"An unexpected error occurred.");
            PostQuitMessage(1);
            return 0;
        }
    }

    return DefWindowProcW(hwnd, message, wparam, lparam);
}

LRESULT LiveApp::handle_message(UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
    case WM_CREATE:
        create_controls();
        populate_controls();
        sync_settings_from_controls();
        layout_controls();
        resize_image();
        SetTimer(hwnd_, kPollTimer, kPollIntervalMs, nullptr);
        return 0;

    case WM_SIZE:
        layout_controls();
        if (capturing_) {
            resize_image();
            rebuild_image_from_history();
        } else {
            graph_dirty_ = true;
        }
        {
            const RECT rect = graph_rect();
            InvalidateRect(hwnd_, &rect, FALSE);
        }
        return 0;

    case WM_GETMINMAXINFO: {
        auto* info = reinterpret_cast<MINMAXINFO*>(lparam);
        info->ptMinTrackSize.x = 1280;
        info->ptMinTrackSize.y = 760;
        return 0;
    }

    case WM_ERASEBKGND:
        return 1;

    case WM_HSCROLL:
        if (reinterpret_cast<HWND>(lparam) == max_db_slider_) {
            apply_live_settings_change(MaxDbSlider);
        }
        return 0;

    case WM_MOUSEMOVE: {
        TRACKMOUSEEVENT event {};
        event.cbSize = sizeof(event);
        event.dwFlags = TME_LEAVE;
        event.hwndTrack = hwnd_;
        if (!tracking_mouse_) {
            TrackMouseEvent(&event);
            tracking_mouse_ = true;
        }
        POINT point { GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam) };
        update_cursor_readout(point, false);
        return 0;
    }

    case WM_MOUSELEAVE:
        tracking_mouse_ = false;
        update_cursor_readout({}, true);
        return 0;

    case WM_COMMAND: {
        const int id = LOWORD(wparam);
        const int code = HIWORD(wparam);
        if (id == StartButton) {
            if (settings_dirty_) {
                apply_settings_with_diagnostics(this);
                if (settings_dirty_) {
                    return 0;
                }
            }
            start_capture();
        } else if (id == StopButton) {
            stop_playback();
            stop_capture();
        } else if (id == PlayButton) {
            loop_playback_ = false;
            start_playback();
        } else if (id == LoopButton) {
            start_loop_playback();
        } else if (id == SaveButton) {
            save_capture();
        } else if (id == SavePngButton) {
            save_png();
        } else if (id == SaveCsvButton) {
            save_note_csv();
        } else if (id == LoadPresetButton) {
            load_preset();
        } else if (id == SavePresetButton) {
            save_preset();
        } else if (id == FreezeButton && code == BN_CLICKED) {
            freeze_display_ = !freeze_display_;
            SetWindowTextW(freeze_button_, freeze_display_ ? L"Unfreeze" : L"Freeze");
            if (freeze_display_) {
                capture_freeze_snapshot();
            }
            graph_dirty_ = true;
            update_status();
        } else if (id == MarkInButton) {
            marker_in_sample_ = marker_sample_from_cursor();
            marker_in_set_ = true;
            graph_dirty_ = true;
            update_status();
        } else if (id == MarkOutButton) {
            marker_out_sample_ = marker_sample_from_cursor();
            marker_out_set_ = true;
            graph_dirty_ = true;
            update_status();
        } else if (id == ApplyButton) {
            apply_settings_with_diagnostics(this);
        } else if (
            (id == SampleRateCombo || id == FftCombo || id == PaletteCombo || id == ScaleCombo ||
                id == DisplayCombo || id == ChannelCombo || id == AmplitudeCombo || id == MonoChannelCombo ||
                id == HistoryCombo || id == ScrollSpeedCombo || id == InstrumentCombo) &&
            code == CBN_SELCHANGE) {
            apply_live_settings_change(id);
        } else if ((id == MaxFreqCombo || id == MinFreqCombo) &&
            (code == CBN_SELCHANGE || code == CBN_EDITCHANGE || code == CBN_EDITUPDATE || code == CBN_KILLFOCUS)) {
            apply_live_settings_change(id);
        } else if ((id == GridCheck || id == MusicGridCheck || id == BeatMarksCheck) && code == BN_CLICKED) {
            apply_live_settings_change(id);
        } else if (id == PeakHoldCheck && code == BN_CLICKED) {
            show_peak_hold_ = SendMessageW(peak_hold_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;
            graph_dirty_ = true;
            update_status();
        } else if (id == AverageCheck && code == BN_CLICKED) {
            show_average_ = SendMessageW(average_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;
            graph_dirty_ = true;
            update_status();
        }
        return 0;
    }

    case WM_TIMER: {
        if (wparam == kPollTimer && capturing_) {
            consume_loopback_packets();
        }
        const ULONGLONG now = GetTickCount64();
        if (now >= next_auto_level_tick_) {
            refresh_auto_levels(false);
            next_auto_level_tick_ = now + 2000;
        }
        if (now >= next_metadata_tick_) {
            update_now_playing();
            next_metadata_tick_ = now + 1000;
        }
        if (now >= next_status_tick_) {
            update_status();
            next_status_tick_ = now + 250;
        }
        if (graph_dirty_) {
            const RECT rect = graph_rect();
            InvalidateRect(hwnd_, &rect, FALSE);
            graph_dirty_ = false;
        }
        return 0;
    }

    case WM_PAINT: {
        PAINTSTRUCT paint_struct {};
        const HDC dc = BeginPaint(hwnd_, &paint_struct);
        paint(dc);
        EndPaint(hwnd_, &paint_struct);
        return 0;
    }

    case MM_WOM_DONE:
        if (loop_playback_) {
            stop_playback();
            start_loop_playback();
        } else {
            stop_playback();
        }
        return 0;

    case WM_DESTROY:
        KillTimer(hwnd_, kPollTimer);
        PostQuitMessage(0);
        return 0;

    default:
        return DefWindowProcW(hwnd_, message, wparam, lparam);
    }
}

void LiveApp::create_controls() {
    start_button_ = CreateWindowExW(0, L"BUTTON", L"Start", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(StartButton), nullptr, nullptr);
    stop_button_ = CreateWindowExW(0, L"BUTTON", L"Stop", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(StopButton), nullptr, nullptr);
    play_button_ = CreateWindowExW(0, L"BUTTON", L"Play", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(PlayButton), nullptr, nullptr);
    save_button_ = CreateWindowExW(0, L"BUTTON", L"Save WAV", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth + 24, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(SaveButton), nullptr, nullptr);
    freeze_button_ = CreateWindowExW(0, L"BUTTON", L"Freeze", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(FreezeButton), nullptr, nullptr);
    apply_button_ = CreateWindowExW(0, L"BUTTON", L"Apply", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(ApplyButton), nullptr, nullptr);
    loop_button_ = CreateWindowExW(0, L"BUTTON", L"Loop Play", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(LoopButton), nullptr, nullptr);
    mark_in_button_ = CreateWindowExW(0, L"BUTTON", L"Mark In", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(MarkInButton), nullptr, nullptr);
    mark_out_button_ = CreateWindowExW(0, L"BUTTON", L"Mark Out", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(MarkOutButton), nullptr, nullptr);
    save_png_button_ = CreateWindowExW(0, L"BUTTON", L"Save PNG", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(SavePngButton), nullptr, nullptr);
    save_csv_button_ = CreateWindowExW(0, L"BUTTON", L"Save CSV", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(SaveCsvButton), nullptr, nullptr);
    load_preset_button_ = CreateWindowExW(0, L"BUTTON", L"Load Preset", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth + 10, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(LoadPresetButton), nullptr, nullptr);
    save_preset_button_ = CreateWindowExW(0, L"BUTTON", L"Save Preset", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        0, 0, kButtonWidth + 10, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(SavePresetButton), nullptr, nullptr);
    tuner_start_button_ = CreateWindowExW(0, L"BUTTON", L"Tuner Start", WS_CHILD | BS_PUSHBUTTON,
        0, 0, kButtonWidth + 10, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(TunerStartButton), nullptr, nullptr);
    tuner_stop_button_ = CreateWindowExW(0, L"BUTTON", L"Tuner Stop", WS_CHILD | BS_PUSHBUTTON,
        0, 0, kButtonWidth + 10, kButtonHeight, hwnd_, reinterpret_cast<HMENU>(TunerStopButton), nullptr, nullptr);

    max_db_label_ = CreateWindowExW(0, L"STATIC", L"Sensitivity", WS_CHILD | WS_VISIBLE, 0, 0, 180, 18, hwnd_, nullptr, nullptr, nullptr);
    max_db_slider_ = CreateWindowExW(
        0, TRACKBAR_CLASSW, nullptr,
        WS_CHILD | WS_VISIBLE | TBS_HORZ | TBS_AUTOTICKS,
        0, 0, 260, 32, hwnd_, reinterpret_cast<HMENU>(MaxDbSlider), nullptr, nullptr);
    range_label_ = nullptr;
    range_slider_ = nullptr;

    status_bar_ = CreateWindowExW(
        0, STATUSCLASSNAMEW, nullptr,
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
        0, 0, 0, 0, hwnd_, nullptr, nullptr, nullptr);

    sample_rate_label_ = CreateWindowExW(0, L"STATIC", L"Sample Rate", WS_CHILD | WS_VISIBLE, 0, 0, 90, 18, hwnd_, nullptr, nullptr, nullptr);
    sample_rate_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 90, 200, hwnd_, reinterpret_cast<HMENU>(SampleRateCombo), nullptr, nullptr);
    fft_label_ = CreateWindowExW(0, L"STATIC", L"FFT", WS_CHILD | WS_VISIBLE, 0, 0, 55, 18, hwnd_, nullptr, nullptr, nullptr);
    fft_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 90, 200, hwnd_, reinterpret_cast<HMENU>(FftCombo), nullptr, nullptr);
    palette_label_ = CreateWindowExW(0, L"STATIC", L"Palette", WS_CHILD | WS_VISIBLE, 0, 0, 65, 18, hwnd_, nullptr, nullptr, nullptr);
    palette_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 95, 200, hwnd_, reinterpret_cast<HMENU>(PaletteCombo), nullptr, nullptr);
    scale_label_ = CreateWindowExW(0, L"STATIC", L"Scale", WS_CHILD | WS_VISIBLE, 0, 0, 55, 18, hwnd_, nullptr, nullptr, nullptr);
    scale_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 105, 200, hwnd_, reinterpret_cast<HMENU>(ScaleCombo), nullptr, nullptr);
    min_freq_label_ = CreateWindowExW(0, L"STATIC", L"Min Freq", WS_CHILD | WS_VISIBLE, 0, 0, 70, 18, hwnd_, nullptr, nullptr, nullptr);
    min_freq_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWN,
        0, 0, 90, 200, hwnd_, reinterpret_cast<HMENU>(MinFreqCombo), nullptr, nullptr);
    max_freq_label_ = CreateWindowExW(0, L"STATIC", L"Max Freq", WS_CHILD | WS_VISIBLE, 0, 0, 70, 18, hwnd_, nullptr, nullptr, nullptr);
    max_freq_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWN,
        0, 0, 95, 200, hwnd_, reinterpret_cast<HMENU>(MaxFreqCombo), nullptr, nullptr);
    channel_label_ = CreateWindowExW(0, L"STATIC", L"Layout", WS_CHILD | WS_VISIBLE, 0, 0, 70, 18, hwnd_, nullptr, nullptr, nullptr);
    channel_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 95, 200, hwnd_, reinterpret_cast<HMENU>(ChannelCombo), nullptr, nullptr);
    mono_channel_label_ = CreateWindowExW(0, L"STATIC", L"Channel", WS_CHILD | WS_VISIBLE, 0, 0, 80, 18, hwnd_, nullptr, nullptr, nullptr);
    mono_channel_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 105, 200, hwnd_, reinterpret_cast<HMENU>(MonoChannelCombo), nullptr, nullptr);
    display_label_ = CreateWindowExW(0, L"STATIC", L"Display", WS_CHILD | WS_VISIBLE, 0, 0, 60, 18, hwnd_, nullptr, nullptr, nullptr);
    display_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 120, 200, hwnd_, reinterpret_cast<HMENU>(DisplayCombo), nullptr, nullptr);
    amplitude_label_ = CreateWindowExW(0, L"STATIC", L"Amplitude", WS_CHILD | WS_VISIBLE, 0, 0, 70, 18, hwnd_, nullptr, nullptr, nullptr);
    amplitude_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 95, 200, hwnd_, reinterpret_cast<HMENU>(AmplitudeCombo), nullptr, nullptr);
    grid_check_ = CreateWindowExW(0, L"BUTTON", L"Grid", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        0, 0, 70, 22, hwnd_, reinterpret_cast<HMENU>(GridCheck), nullptr, nullptr);
    peak_hold_check_ = CreateWindowExW(0, L"BUTTON", L"Peak Hold", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        0, 0, 96, 22, hwnd_, reinterpret_cast<HMENU>(PeakHoldCheck), nullptr, nullptr);
    average_check_ = CreateWindowExW(0, L"BUTTON", L"Average", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        0, 0, 88, 22, hwnd_, reinterpret_cast<HMENU>(AverageCheck), nullptr, nullptr);
    history_label_ = CreateWindowExW(0, L"STATIC", L"History", WS_CHILD | WS_VISIBLE, 0, 0, 60, 18, hwnd_, nullptr, nullptr, nullptr);
    history_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 92, 180, hwnd_, reinterpret_cast<HMENU>(HistoryCombo), nullptr, nullptr);
    scroll_speed_label_ = CreateWindowExW(0, L"STATIC", L"Scroll", WS_CHILD | WS_VISIBLE, 0, 0, 55, 18, hwnd_, nullptr, nullptr, nullptr);
    scroll_speed_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 92, 180, hwnd_, reinterpret_cast<HMENU>(ScrollSpeedCombo), nullptr, nullptr);
    music_grid_check_ = CreateWindowExW(0, L"BUTTON", L"Music Grid", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        0, 0, 92, 22, hwnd_, reinterpret_cast<HMENU>(MusicGridCheck), nullptr, nullptr);
    beat_marks_check_ = CreateWindowExW(0, L"BUTTON", L"Beat Marks", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        0, 0, 92, 22, hwnd_, reinterpret_cast<HMENU>(BeatMarksCheck), nullptr, nullptr);
    instrument_label_ = CreateWindowExW(0, L"STATIC", L"Instrument", WS_CHILD | WS_VISIBLE, 0, 0, 70, 18, hwnd_, nullptr, nullptr, nullptr);
    instrument_combo_ = CreateWindowExW(0, L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST,
        0, 0, 140, 220, hwnd_, reinterpret_cast<HMENU>(InstrumentCombo), nullptr, nullptr);

    SendMessageW(max_db_slider_, TBM_SETRANGE, TRUE, MAKELPARAM(24, 180));
    SendMessageW(max_db_slider_, TBM_SETTICFREQ, 12, 0);
    SendMessageW(max_db_slider_, TBM_SETPAGESIZE, 0, 12);
    SendMessageW(max_db_slider_, TBM_SETPOS, TRUE, static_cast<LPARAM>(90));
    SendMessageW(peak_hold_check_, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(average_check_, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(music_grid_check_, BM_SETCHECK, BST_UNCHECKED, 0);
    SendMessageW(beat_marks_check_, BM_SETCHECK, BST_UNCHECKED, 0);
    EnableWindow(apply_button_, FALSE);
    update_transport_controls();
}

void LiveApp::populate_controls() {
    const wchar_t* fft_sizes[] = { L"512", L"1024", L"2048", L"4096", L"8192" };
    const wchar_t* palettes[] = { L"spectrum", L"gray", L"blue", L"green", L"amber", L"ice" };
    const wchar_t* scales[] = { L"linear", L"logarithmic", L"octave" };
    const wchar_t* min_freqs[] = { L"0", L"1", L"5", L"10", L"20", L"27.5", L"50", L"100" };
    const wchar_t* max_freqs[] = { L"Auto", L"800", L"1000", L"1500", L"2000", L"4000", L"8000", L"12000", L"16000", L"20000", L"24000", L"30000" };
    const wchar_t* channels[] = { L"auto", L"mono", L"stereo" };
    const wchar_t* mono_channels[] = { L"mix", L"left", L"right", L"Stereo" };
    const wchar_t* displays[] = { L"spectrogram", L"spectrum", L"tuner", L"vectorscope", L"room" };
    const wchar_t* amplitudes[] = { L"dB", L"linear" };
    const wchar_t* histories[] = { L"5", L"10", L"15", L"30", L"60" };
    const wchar_t* scroll_speeds[] = { L"0.5", L"1.0", L"2.0", L"4.0" };
    const wchar_t* instruments[] = { L"chromatic", L"guitar", L"bass", L"ukulele", L"violin", L"viola", L"cello", L"mandolin" };

    for (int rate : available_sample_rates_) {
        const std::wstring text = std::to_wstring(rate);
        SendMessageW(sample_rate_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str()));
    }
    for (const auto* item : fft_sizes) {
        SendMessageW(fft_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : palettes) {
        SendMessageW(palette_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : scales) {
        SendMessageW(scale_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : min_freqs) {
        SendMessageW(min_freq_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : max_freqs) {
        SendMessageW(max_freq_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : channels) {
        SendMessageW(channel_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : mono_channels) {
        SendMessageW(mono_channel_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : displays) {
        SendMessageW(display_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : amplitudes) {
        SendMessageW(amplitude_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : histories) {
        SendMessageW(history_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : scroll_speeds) {
        SendMessageW(scroll_speed_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    for (const auto* item : instruments) {
        SendMessageW(instrument_combo_, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }

    if (std::find(available_sample_rates_.begin(), available_sample_rates_.end(), 44100) != available_sample_rates_.end()) {
        set_combo_to_text(sample_rate_combo_, L"44100");
    } else if (!available_sample_rates_.empty()) {
        const std::wstring text = std::to_wstring(available_sample_rates_.front());
        set_combo_to_text(sample_rate_combo_, text.c_str());
    }
    set_combo_to_text(fft_combo_, L"2048");
    set_combo_to_text(palette_combo_, L"spectrum");
    set_combo_to_text(scale_combo_, L"linear");
    set_combo_to_text(min_freq_combo_, L"0");
    set_combo_to_text(max_freq_combo_, L"12000");
    set_combo_to_text(channel_combo_, L"auto");
    set_combo_to_text(mono_channel_combo_, L"mix");
    set_combo_to_text(display_combo_, L"spectrogram");
    set_combo_to_text(amplitude_combo_, L"dB");
    set_combo_to_text(history_combo_, L"15");
    set_combo_to_text(scroll_speed_combo_, L"1.0");
    set_combo_to_text(instrument_combo_, L"chromatic");
    SendMessageW(grid_check_, BM_SETCHECK, BST_CHECKED, 0);
    update_level_labels();
    update_channel_controls();
    update_view_controls();
}

void LiveApp::layout_controls() {
    const int row1 = 8;
    const int row1b = 42;
    const int row2_label = 76;
    const int row2_combo = 92;
    const int row3_label = 120;
    const int row3_combo = 136;
    RECT client {};
    GetClientRect(hwnd_, &client);

    MoveWindow(start_button_, 12, row1, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(stop_button_, 12 + (kButtonWidth + 8), row1, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(play_button_, 12 + 2 * (kButtonWidth + 8), row1, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(save_button_, 12 + 3 * (kButtonWidth + 8), row1, kButtonWidth + 24, kButtonHeight, TRUE);
    const int freeze_left = 12 + 3 * (kButtonWidth + 8) + (kButtonWidth + 40);
    MoveWindow(freeze_button_, freeze_left, row1, kButtonWidth, kButtonHeight, TRUE);
    const int apply_left = freeze_left + kButtonWidth + 8;
    MoveWindow(apply_button_, apply_left, row1, kButtonWidth, kButtonHeight, TRUE);

    const int slider_left = apply_left + kButtonWidth + 28;
    const int slider_label_width = 120;
    const int slider_width = (std::max)(static_cast<int>(client.right - slider_left - slider_label_width - 20), 220);
    MoveWindow(max_db_label_, slider_left, row1 - 1, slider_label_width + 20, 18, TRUE);
    MoveWindow(max_db_slider_, slider_left + slider_label_width, row1 - 4, slider_width, 28, TRUE);

    MoveWindow(loop_button_, 12, row1b, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(mark_in_button_, 12 + (kButtonWidth + 8), row1b, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(mark_out_button_, 12 + 2 * (kButtonWidth + 8), row1b, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(save_png_button_, 12 + 3 * (kButtonWidth + 8), row1b, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(save_csv_button_, 12 + 4 * (kButtonWidth + 8), row1b, kButtonWidth, kButtonHeight, TRUE);
    MoveWindow(load_preset_button_, 12 + 5 * (kButtonWidth + 8), row1b, kButtonWidth + 10, kButtonHeight, TRUE);
    MoveWindow(save_preset_button_, 12 + 6 * (kButtonWidth + 8) + 10, row1b, kButtonWidth + 10, kButtonHeight, TRUE);
    MoveWindow(instrument_label_, client.right - 232, row1b + 4, 70, 18, TRUE);
    MoveWindow(instrument_combo_, client.right - 156, row1b, 144, 220, TRUE);

    MoveWindow(sample_rate_combo_, 12, row2_combo, 90, 250, TRUE);
    MoveWindow(fft_combo_, 116, row2_combo, 90, 250, TRUE);
    MoveWindow(palette_combo_, 220, row2_combo, 95, 250, TRUE);
    MoveWindow(scale_combo_, 329, row2_combo, 105, 250, TRUE);
    MoveWindow(min_freq_combo_, 448, row2_combo, 90, 250, TRUE);
    MoveWindow(max_freq_combo_, 552, row2_combo, 95, 250, TRUE);
    MoveWindow(channel_combo_, 661, row2_combo, 95, 250, TRUE);
    MoveWindow(display_combo_, 770, row2_combo, 120, 250, TRUE);
    MoveWindow(sample_rate_label_, 12, row2_label, 90, 16, TRUE);
    MoveWindow(fft_label_, 116, row2_label, 60, 16, TRUE);
    MoveWindow(palette_label_, 220, row2_label, 70, 16, TRUE);
    MoveWindow(scale_label_, 329, row2_label, 55, 16, TRUE);
    MoveWindow(min_freq_label_, 448, row2_label, 70, 16, TRUE);
    MoveWindow(max_freq_label_, 552, row2_label, 70, 16, TRUE);
    MoveWindow(channel_label_, 661, row2_label, 70, 16, TRUE);
    MoveWindow(display_label_, 770, row2_label, 60, 16, TRUE);

    MoveWindow(amplitude_combo_, 12, row3_combo, 95, 250, TRUE);
    MoveWindow(mono_channel_combo_, 130, row3_combo, 105, 250, TRUE);
    MoveWindow(grid_check_, 249, row3_combo - 1, 70, 24, TRUE);
    MoveWindow(peak_hold_check_, 336, row3_combo - 1, 96, 24, TRUE);
    MoveWindow(average_check_, 446, row3_combo - 1, 88, 24, TRUE);
    MoveWindow(history_combo_, 560, row3_combo, 92, 220, TRUE);
    MoveWindow(scroll_speed_combo_, 675, row3_combo, 92, 220, TRUE);
    MoveWindow(music_grid_check_, 790, row3_combo - 1, 92, 24, TRUE);
    MoveWindow(beat_marks_check_, 896, row3_combo - 1, 92, 24, TRUE);
    MoveWindow(amplitude_label_, 12, row3_label, 65, 16, TRUE);
    MoveWindow(mono_channel_label_, 130, row3_label, 80, 16, TRUE);
    MoveWindow(history_label_, 560, row3_label, 55, 16, TRUE);
    MoveWindow(scroll_speed_label_, 675, row3_label, 55, 16, TRUE);

    SendMessageW(status_bar_, WM_SIZE, 0, 0);
    const int parts[] { -1 };
    SendMessageW(status_bar_, SB_SETPARTS, static_cast<WPARAM>(std::size(parts)), reinterpret_cast<LPARAM>(parts));
    update_view_controls();
}

void LiveApp::resize_image() {
    const RECT plot = plot_rect();
    const int width = (std::max)(static_cast<int>(plot.right - plot.left), 320);
    const int pane_count = spectrogram_pane_count();
    const int pane_gap = pane_count > 1 ? 12 : 0;
    const int height = (std::max)(static_cast<int>((plot.bottom - plot.top - pane_gap) / pane_count), 120);
    settings_.image_height = height;
    image_ = gram::make_image(width, height);
    left_image_ = gram::make_image(width, height);
    right_image_ = gram::make_image(width, height);
    gram::clear_image(image_);
    gram::clear_image(left_image_);
    gram::clear_image(right_image_);
    graph_dirty_ = true;
}

int LiveApp::status_bar_height() const {
    if (status_bar_ == nullptr) {
        return 0;
    }

    RECT rect {};
    if (!GetWindowRect(status_bar_, &rect)) {
        return 0;
    }
    return static_cast<int>(rect.bottom - rect.top);
}

RECT LiveApp::graph_rect() const {
    RECT client {};
    GetClientRect(hwnd_, &client);
    const int bottom_margin = status_bar_height() > 0 ? status_bar_height() + 6 : 10;
    RECT rect { 10, kControlsHeight, client.right - 10, client.bottom - bottom_margin };
    return rect;
}

RECT LiveApp::plot_rect() const {
    RECT rect = graph_rect();
    rect.left += 64;
    rect.right -= 64;
    rect.top += 10;
    rect.bottom -= (display_mode_ == DisplayMode::Spectrum || display_mode_ == DisplayMode::Room) ? 36 : 10;
    return rect;
}

void LiveApp::clear_for_new_capture() {
    captured_samples_.clear();
    captured_left_samples_.clear();
    captured_right_samples_.clear();
    analysis_buffer_.clear();
    left_analysis_buffer_.clear();
    right_analysis_buffer_.clear();
    resample_buffer_.clear();
    left_resample_buffer_.clear();
    right_resample_buffer_.clear();
    db_bins_.clear();
    db_bins_left_.clear();
    db_bins_right_.clear();
    peak_hold_bins_.clear();
    peak_hold_bins_left_.clear();
    peak_hold_bins_right_.clear();
    average_bins_.clear();
    average_bins_left_.clear();
    average_bins_right_.clear();
    frozen_db_bins_.clear();
    frozen_db_bins_left_.clear();
    frozen_db_bins_right_.clear();
    frozen_peak_hold_bins_.clear();
    frozen_peak_hold_bins_left_.clear();
    frozen_peak_hold_bins_right_.clear();
    frozen_average_bins_.clear();
    frozen_average_bins_left_.clear();
    frozen_average_bins_right_.clear();
    analysis_cursor_ = 0;
    resample_position_ = 0.0;
    scroll_credit_ = 0.0;
    profiling_ = {};
    auto_levels_ready_ = false;
    settings_.max_db = 0.0;
    settings_.min_db = settings_.max_db - sensitivity_db_;
    gram::clear_image(image_);
    gram::clear_image(left_image_);
    gram::clear_image(right_image_);
    gram::clear_image(frozen_image_);
    gram::clear_image(frozen_left_image_);
    gram::clear_image(frozen_right_image_);
    transient_markers_.clear();
    last_frame_rms_ = 0.0;
    stereo_correlation_ = 0.0;
    reference_available_ = false;
    graph_dirty_ = true;
    next_status_tick_ = 0;
}

void LiveApp::sync_settings_from_controls() {
    const int new_sample_rate = parse_combo_int(sample_rate_combo_);
    if (new_sample_rate <= 0) {
        throw std::runtime_error("Choose a valid analysis sample rate before applying settings.");
    }
    if (new_sample_rate == 11025) {
        throw std::runtime_error("11025 Hz is temporarily disabled in the live app because it is still causing capture startup failures.");
    }
    const int new_fft = parse_combo_int(fft_combo_);
    if (new_fft <= 0) {
        throw std::runtime_error("Choose a valid FFT size before applying settings.");
    }
    const double fft_window_seconds =
        static_cast<double>(new_fft) / static_cast<double>(new_sample_rate);
    if (fft_window_seconds > 0.25) {
        throw std::runtime_error("That sample-rate/FFT combination is temporarily blocked because the FFT window is too long for stable live capture on this build.");
    }
    const bool fft_changed = new_fft != settings_.fft_size;
    const bool sample_rate_changed = new_sample_rate != analysis_sample_rate_;
    const bool pipeline_changed = fft_changed || sample_rate_changed;
    const bool restart_capture = capturing_ && pipeline_changed;
    if (restart_capture) {
        stop_capture();
    }

    settings_.fft_size = new_fft;
    settings_.hop_size = new_fft / 4;
    const double requested_min_freq = parse_combo_double(min_freq_combo_);
    settings_.min_frequency = std::clamp(requested_min_freq, 0.0, 1000.0);
    const double requested_max_freq = parse_max_freq(max_freq_combo_);
    if (requested_max_freq <= 0.0) {
        settings_.max_frequency = 0.0;
    } else {
        settings_.max_frequency = std::clamp(requested_max_freq, 800.0, 30000.0);
    }
    if (settings_.max_frequency > 0.0 && settings_.min_frequency >= settings_.max_frequency) {
        throw std::runtime_error("Minimum frequency must be lower than maximum frequency.");
    }
    sensitivity_db_ = static_cast<double>(static_cast<int>(SendMessageW(max_db_slider_, TBM_GETPOS, 0, 0)));
    history_seconds_ = std::clamp(parse_combo_double(history_combo_), 5.0, 60.0);
    scroll_speed_ = std::clamp(parse_combo_double(scroll_speed_combo_), 0.5, 4.0);
    scroll_credit_ = 0.0;
    update_level_labels();

    wchar_t palette_text[32] {};
    GetWindowTextW(palette_combo_, palette_text, 31);
    if (wcscmp(palette_text, L"gray") == 0) {
        settings_.palette_mode = gram::PaletteMode::Gray;
    } else if (wcscmp(palette_text, L"blue") == 0) {
        settings_.palette_mode = gram::PaletteMode::Blue;
    } else if (wcscmp(palette_text, L"green") == 0) {
        settings_.palette_mode = gram::PaletteMode::Green;
    } else if (wcscmp(palette_text, L"amber") == 0) {
        settings_.palette_mode = gram::PaletteMode::Amber;
    } else if (wcscmp(palette_text, L"ice") == 0) {
        settings_.palette_mode = gram::PaletteMode::Ice;
    } else {
        settings_.palette_mode = gram::PaletteMode::Spectrum;
    }

    wchar_t scale_text[32] {};
    GetWindowTextW(scale_combo_, scale_text, 31);
    if (wcscmp(scale_text, L"logarithmic") == 0) {
        settings_.frequency_scale = gram::FrequencyScale::Logarithmic;
    } else if (wcscmp(scale_text, L"octave") == 0) {
        settings_.frequency_scale = gram::FrequencyScale::Octave;
    } else {
        settings_.frequency_scale = gram::FrequencyScale::Linear;
    }

    wchar_t display_text[32] {};
    GetWindowTextW(display_combo_, display_text, 31);
    if (wcscmp(display_text, L"spectrum") == 0) {
        display_mode_ = DisplayMode::Spectrum;
    } else if (wcscmp(display_text, L"tuner") == 0) {
        display_mode_ = DisplayMode::Tuner;
    } else if (wcscmp(display_text, L"vectorscope") == 0) {
        display_mode_ = DisplayMode::Vectorscope;
    } else if (wcscmp(display_text, L"room") == 0) {
        display_mode_ = DisplayMode::Room;
    } else {
        display_mode_ = DisplayMode::Spectrogram;
    }

    show_grid_ = SendMessageW(grid_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;
    show_note_guides_ = SendMessageW(music_grid_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;
    show_beat_marks_ = SendMessageW(beat_marks_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;
    show_peak_hold_ = SendMessageW(peak_hold_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;
    show_average_ = SendMessageW(average_check_, BM_GETCHECK, 0, 0) == BST_CHECKED;

    wchar_t channel_text[32] {};
    GetWindowTextW(channel_combo_, channel_text, 31);
    if (wcscmp(channel_text, L"mono") == 0) {
        channel_display_mode_ = ChannelDisplayMode::Mono;
    } else if (wcscmp(channel_text, L"stereo") == 0) {
        channel_display_mode_ = ChannelDisplayMode::Stereo;
    } else {
        channel_display_mode_ = ChannelDisplayMode::Auto;
    }

    wchar_t mono_channel_text[32] {};
    GetWindowTextW(mono_channel_combo_, mono_channel_text, 31);
    if (wcscmp(mono_channel_text, L"left") == 0) {
        mono_channel_mode_ = MonoChannelMode::Left;
    } else if (wcscmp(mono_channel_text, L"right") == 0) {
        mono_channel_mode_ = MonoChannelMode::Right;
    } else {
        mono_channel_mode_ = MonoChannelMode::Mix;
    }

    wchar_t amplitude_text[32] {};
    GetWindowTextW(amplitude_combo_, amplitude_text, 31);
    if (wcscmp(amplitude_text, L"linear") == 0) {
        spectrum_amplitude_mode_ = SpectrumAmplitudeMode::Linear;
    } else {
        spectrum_amplitude_mode_ = SpectrumAmplitudeMode::Decibels;
    }

    wchar_t instrument_text[32] {};
    GetWindowTextW(instrument_combo_, instrument_text, 31);
    instrument_mode_ = instrument_text;

    if (sample_rate_changed) {
        analysis_sample_rate_ = new_sample_rate;
    }

    resize_image();
    fft_plan_ = gram::FftPlan(settings_.fft_size);
    graph_dirty_ = true;
    auto_levels_ready_ = false;
    refresh_auto_levels(true);
    update_channel_controls();
    update_view_controls();

    if (pipeline_changed) {
        if (!restart_capture) {
            clear_for_new_capture();
        }
    } else {
        rebuild_image_from_history();
    }

    if (restart_capture) {
        start_capture();
    }
    if (freeze_display_) {
        capture_freeze_snapshot();
    }
    settings_dirty_ = false;
    EnableWindow(apply_button_, FALSE);
    update_status();
}

void LiveApp::rebuild_image_from_history() {
    gram::clear_image(image_);
    gram::clear_image(left_image_);
    gram::clear_image(right_image_);
    analysis_buffer_.clear();
    left_analysis_buffer_.clear();
    right_analysis_buffer_.clear();
    peak_hold_bins_.clear();
    peak_hold_bins_left_.clear();
    peak_hold_bins_right_.clear();
    average_bins_.clear();
    average_bins_left_.clear();
    average_bins_right_.clear();
    analysis_cursor_ = 0;
    scroll_credit_ = 0.0;
    profiling_ = {};
    graph_dirty_ = true;

    if (captured_samples_.empty()) {
        return;
    }

    const std::size_t needed_samples =
        static_cast<std::size_t>(settings_.fft_size) +
        static_cast<std::size_t>(settings_.hop_size) * static_cast<std::size_t>((std::max)(image_.width - 1, 0));
    const std::size_t take =
        (std::min)(needed_samples, captured_samples_.size());
    const std::size_t start = captured_samples_.size() - take;

    analysis_buffer_.reserve(take);
    left_analysis_buffer_.reserve(take);
    right_analysis_buffer_.reserve(take);
    for (std::size_t index = start; index < captured_samples_.size(); ++index) {
        analysis_buffer_.push_back(static_cast<double>(captured_samples_[index]) / 32768.0);
        left_analysis_buffer_.push_back(static_cast<double>(captured_left_samples_[index]) / 32768.0);
        right_analysis_buffer_.push_back(static_cast<double>(captured_right_samples_[index]) / 32768.0);
    }

    process_available_analysis_frames();
}

void LiveApp::refresh_auto_levels(bool force) {
    if (display_mode_ != DisplayMode::Spectrum) {
        if (!auto_levels_ready_ || force) {
            settings_.max_db = 0.0;
            settings_.min_db = settings_.max_db - sensitivity_db_;
            auto_levels_ready_ = true;
            graph_dirty_ = true;
        }
        return;
    }

    double peak = -120.0;
    auto fold_peak = [&peak](const std::vector<double>& bins) {
        if (!bins.empty()) {
            peak = (std::max)(peak, *std::max_element(bins.begin(), bins.end()));
        }
    };
    fold_peak(db_bins_);
    fold_peak(db_bins_left_);
    fold_peak(db_bins_right_);

    if (peak <= -119.0) {
        if (!auto_levels_ready_) {
            settings_.max_db = 0.0;
            settings_.min_db = settings_.max_db - sensitivity_db_;
        }
        return;
    }

    const double target_max = std::clamp(peak + 6.0, -72.0, 24.0);
    const bool changed =
        !auto_levels_ready_ ||
        std::abs(target_max - settings_.max_db) >= 1.0 ||
        std::abs((settings_.max_db - settings_.min_db) - sensitivity_db_) >= 0.5;
    settings_.max_db = target_max;
    settings_.min_db = settings_.max_db - sensitivity_db_;
    auto_levels_ready_ = true;

    if (force || changed) {
        graph_dirty_ = true;
        if (freeze_display_) {
            capture_freeze_snapshot();
        }
    }
}

void LiveApp::start_capture() {
    if (capturing_) {
        return;
    }

    clear_for_new_capture();

    const HRESULT create_enumerator = CoCreateInstance(
        __uuidof(MMDeviceEnumerator),
        nullptr,
        CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        reinterpret_cast<void**>(&enumerator_));
    if (FAILED(create_enumerator)) {
        show_logged_hresult(hwnd_, L"start_capture.CoCreateInstance", L"Unable to create MMDevice enumerator.", create_enumerator);
        return;
    }

    const HRESULT default_endpoint = enumerator_->GetDefaultAudioEndpoint(eRender, eMultimedia, &device_);
    if (FAILED(default_endpoint)) {
        show_logged_hresult(hwnd_, L"start_capture.GetDefaultAudioEndpoint", L"Unable to access the default output device.", default_endpoint);
        stop_capture();
        return;
    }

    const HRESULT activate_audio_client =
        device_->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&audio_client_));
    if (FAILED(activate_audio_client)) {
        show_logged_hresult(hwnd_, L"start_capture.ActivateAudioClient", L"Unable to activate the default output device.", activate_audio_client);
        stop_capture();
        return;
    }

    const HRESULT mix_format_result = audio_client_->GetMixFormat(&mix_format_);
    if (FAILED(mix_format_result)) {
        show_logged_hresult(hwnd_, L"start_capture.GetMixFormat", L"Unable to query the device mix format.", mix_format_result);
        stop_capture();
        return;
    }

    capture_sample_rate_ = static_cast<int>(mix_format_->nSamplesPerSec);
    capture_channels_ = static_cast<int>(mix_format_->nChannels);
    last_stream_sample_rate_ = capture_sample_rate_;
    last_stream_channels_ = (std::max)(1, (std::min)(capture_channels_, 2));

    const REFERENCE_TIME buffer_duration = 10000000;
    const HRESULT initialize_client = audio_client_->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK,
        buffer_duration,
        0,
        mix_format_,
        nullptr);
    if (FAILED(initialize_client)) {
        show_logged_hresult(hwnd_, L"start_capture.Initialize", L"Unable to initialize loopback capture on the default output device.", initialize_client);
        stop_capture();
        return;
    }

    const HRESULT capture_service =
        audio_client_->GetService(__uuidof(IAudioCaptureClient), reinterpret_cast<void**>(&capture_client_));
    if (FAILED(capture_service)) {
        show_logged_hresult(hwnd_, L"start_capture.GetService", L"Unable to create the audio capture client.", capture_service);
        stop_capture();
        return;
    }

    const HRESULT start_client = audio_client_->Start();
    if (FAILED(start_client)) {
        show_logged_hresult(hwnd_, L"start_capture.Start", L"Unable to start loopback capture.", start_client);
        stop_capture();
        return;
    }

    capturing_ = true;
    update_channel_controls();
    update_transport_controls();
    update_view_controls();
    next_auto_level_tick_ = GetTickCount64() + 2000;
    update_status();
    next_status_tick_ = GetTickCount64() + 250;
    graph_dirty_ = true;
}

void LiveApp::stop_capture() {
    capturing_ = false;
    if (audio_client_ != nullptr) {
        audio_client_->Stop();
    }
    if (mix_format_ != nullptr) {
        CoTaskMemFree(mix_format_);
        mix_format_ = nullptr;
    }
    safe_release(capture_client_);
    safe_release(audio_client_);
    safe_release(device_);
    safe_release(enumerator_);
    capture_sample_rate_ = 0;
    capture_channels_ = 0;
    resample_buffer_.clear();
    left_resample_buffer_.clear();
    right_resample_buffer_.clear();
    resample_position_ = 0.0;
    update_channel_controls();
    update_transport_controls();
    update_view_controls();
    next_auto_level_tick_ = GetTickCount64() + 2000;
    update_status();
    next_status_tick_ = GetTickCount64() + 250;
    graph_dirty_ = true;
}

void LiveApp::consume_loopback_packets() {
    if (!capturing_ || capture_client_ == nullptr) {
        return;
    }

    UINT32 packet_size = 0;
    while (SUCCEEDED(capture_client_->GetNextPacketSize(&packet_size)) && packet_size > 0) {
        BYTE* data = nullptr;
        UINT32 frames = 0;
        DWORD flags = 0;
        if (FAILED(capture_client_->GetBuffer(&data, &frames, &flags, nullptr, nullptr))) {
            break;
        }

        append_capture_chunk(data, frames, (flags & AUDCLNT_BUFFERFLAGS_SILENT) != 0);
        capture_client_->ReleaseBuffer(frames);

        if (FAILED(capture_client_->GetNextPacketSize(&packet_size))) {
            break;
        }
    }
}

void LiveApp::append_capture_chunk(const BYTE* data, UINT32 frames, bool silent) {
    if (mix_format_ == nullptr || frames == 0) {
        return;
    }

    const int channels = (std::max)(capture_channels_, 1);
    const int bytes_per_frame = (std::max)(static_cast<int>(mix_format_->nBlockAlign), 1);

    for (UINT32 frame = 0; frame < frames; ++frame) {
        double mono = 0.0;
        double left = 0.0;
        double right = 0.0;
        if (!silent && data != nullptr) {
            if (is_float_format(mix_format_)) {
                const auto* values = reinterpret_cast<const float*>(data + frame * bytes_per_frame);
                left = static_cast<double>(values[0]);
                right = channels > 1 ? static_cast<double>(values[1]) : left;
                for (int channel = 0; channel < channels; ++channel) {
                    mono += values[channel];
                }
                mono /= static_cast<double>(channels);
            } else if (is_pcm_format(mix_format_)) {
                if (mix_format_->wBitsPerSample == 16) {
                    const auto* values = reinterpret_cast<const std::int16_t*>(data + frame * bytes_per_frame);
                    left = static_cast<double>(values[0]) / 32768.0;
                    right = channels > 1 ? static_cast<double>(values[1]) / 32768.0 : left;
                    for (int channel = 0; channel < channels; ++channel) {
                        mono += static_cast<double>(values[channel]) / 32768.0;
                    }
                    mono /= static_cast<double>(channels);
                } else if (mix_format_->wBitsPerSample == 32) {
                    const auto* values = reinterpret_cast<const std::int32_t*>(data + frame * bytes_per_frame);
                    left = static_cast<double>(values[0]) / 2147483648.0;
                    right = channels > 1 ? static_cast<double>(values[1]) / 2147483648.0 : left;
                    for (int channel = 0; channel < channels; ++channel) {
                        mono += static_cast<double>(values[channel]) / 2147483648.0;
                    }
                    mono /= static_cast<double>(channels);
                }
            }
        }

        resample_buffer_.push_back(std::clamp(mono, -1.0, 1.0));
        left_resample_buffer_.push_back(std::clamp(left, -1.0, 1.0));
        right_resample_buffer_.push_back(std::clamp(right, -1.0, 1.0));
    }

    if (capture_sample_rate_ <= 0 || analysis_sample_rate_ <= 0) {
        return;
    }

    const double step = static_cast<double>(capture_sample_rate_) / static_cast<double>(analysis_sample_rate_);
    while (resample_position_ + 1.0 < static_cast<double>(resample_buffer_.size())) {
        const auto left_index = static_cast<std::size_t>(resample_position_);
        const double fraction = resample_position_ - static_cast<double>(left_index);
        const double mix_left = resample_buffer_[left_index];
        const double mix_right = resample_buffer_[left_index + 1];
        const double left_left = left_resample_buffer_[left_index];
        const double left_right = left_resample_buffer_[left_index + 1];
        const double right_left = right_resample_buffer_[left_index];
        const double right_right = right_resample_buffer_[left_index + 1];
        append_analysis_sample(
            mix_left + (mix_right - mix_left) * fraction,
            left_left + (left_right - left_left) * fraction,
            right_left + (right_right - right_left) * fraction);
        resample_position_ += step;
    }

    const std::size_t consumed = static_cast<std::size_t>(resample_position_);
    if (consumed > 0) {
        resample_buffer_.erase(
            resample_buffer_.begin(),
            resample_buffer_.begin() + static_cast<std::ptrdiff_t>(consumed));
        left_resample_buffer_.erase(
            left_resample_buffer_.begin(),
            left_resample_buffer_.begin() + static_cast<std::ptrdiff_t>(consumed));
        right_resample_buffer_.erase(
            right_resample_buffer_.begin(),
            right_resample_buffer_.begin() + static_cast<std::ptrdiff_t>(consumed));
        resample_position_ -= static_cast<double>(consumed);
    }
}

void LiveApp::append_analysis_sample(double mix_sample, double left_sample, double right_sample) {
    const auto mix = std::clamp(mix_sample, -1.0, 1.0);
    const auto left = std::clamp(left_sample, -1.0, 1.0);
    const auto right = std::clamp(right_sample, -1.0, 1.0);
    analysis_buffer_.push_back(mix);
    left_analysis_buffer_.push_back(left);
    right_analysis_buffer_.push_back(right);
    captured_samples_.push_back(static_cast<std::int16_t>(std::lround(mix * 32767.0)));
    captured_left_samples_.push_back(static_cast<std::int16_t>(std::lround(left * 32767.0)));
    captured_right_samples_.push_back(static_cast<std::int16_t>(std::lround(right * 32767.0)));
    process_available_analysis_frames();
    trim_analysis_buffer();
}

void LiveApp::process_available_analysis_frames() {
    while (analysis_cursor_ + static_cast<std::size_t>(settings_.fft_size) <= analysis_buffer_.size()) {
        double rms = 0.0;
        double left_energy = 0.0;
        double right_energy = 0.0;
        double cross = 0.0;
        for (int sample = 0; sample < settings_.fft_size; ++sample) {
            const double mix = analysis_buffer_[analysis_cursor_ + static_cast<std::size_t>(sample)];
            const double left = left_analysis_buffer_[analysis_cursor_ + static_cast<std::size_t>(sample)];
            const double right = right_analysis_buffer_[analysis_cursor_ + static_cast<std::size_t>(sample)];
            rms += mix * mix;
            left_energy += left * left;
            right_energy += right * right;
            cross += left * right;
        }
        rms = std::sqrt(rms / static_cast<double>(settings_.fft_size));
        if (left_energy > 1e-9 && right_energy > 1e-9) {
            stereo_correlation_ = cross / std::sqrt(left_energy * right_energy);
        }
        const std::size_t frame_end_sample = captured_samples_.size();
        if (rms > 0.01 && rms > last_frame_rms_ * 1.45) {
            transient_markers_.push_back(frame_end_sample);
            while (transient_markers_.size() > 128) {
                transient_markers_.erase(transient_markers_.begin());
            }
        }
        last_frame_rms_ = rms;

        fft_plan_.compute_db(
            analysis_buffer_.data() + analysis_cursor_,
            db_bins_,
            kAnalysisFloorDb,
            &profiling_);
        fft_plan_.compute_db(
            left_analysis_buffer_.data() + analysis_cursor_,
            db_bins_left_,
            kAnalysisFloorDb,
            nullptr);
        fft_plan_.compute_db(
            right_analysis_buffer_.data() + analysis_cursor_,
            db_bins_right_,
            kAnalysisFloorDb,
            nullptr);
        update_peak_hold_bins(peak_hold_bins_, db_bins_);
        update_peak_hold_bins(peak_hold_bins_left_, db_bins_left_);
        update_peak_hold_bins(peak_hold_bins_right_, db_bins_right_);
        update_average_bins(average_bins_, db_bins_);
        update_average_bins(average_bins_left_, db_bins_left_);
        update_average_bins(average_bins_right_, db_bins_right_);
        const auto mix_column = gram::render_column(db_bins_, analysis_sample_rate_, settings_);
        const auto left_column = gram::render_column(db_bins_left_, analysis_sample_rate_, settings_);
        const auto right_column = gram::render_column(db_bins_right_, analysis_sample_rate_, settings_);
        scroll_credit_ += effective_scroll_multiplier();
        while (scroll_credit_ >= 1.0) {
            gram::scroll_image_left(image_);
            gram::scroll_image_left(left_image_);
            gram::scroll_image_left(right_image_);
            gram::set_image_column(
                image_,
                image_.width - 1,
                mix_column,
                settings_);
            gram::set_image_column(
                left_image_,
                left_image_.width - 1,
                left_column,
                settings_);
            gram::set_image_column(
                right_image_,
                right_image_.width - 1,
                right_column,
                settings_);
            scroll_credit_ -= 1.0;
        }
        analysis_cursor_ += static_cast<std::size_t>(settings_.hop_size);
        graph_dirty_ = true;
    }
    if (freeze_display_) {
        capture_freeze_snapshot();
    }
}

void LiveApp::trim_analysis_buffer() {
    const std::size_t keep = static_cast<std::size_t>(settings_.fft_size * 4);
    if (analysis_cursor_ > keep && analysis_cursor_ < analysis_buffer_.size()) {
        const std::size_t erase_count = analysis_cursor_ - static_cast<std::size_t>(settings_.fft_size);
        analysis_buffer_.erase(
            analysis_buffer_.begin(),
            analysis_buffer_.begin() + static_cast<std::ptrdiff_t>(erase_count));
        left_analysis_buffer_.erase(
            left_analysis_buffer_.begin(),
            left_analysis_buffer_.begin() + static_cast<std::ptrdiff_t>(erase_count));
        right_analysis_buffer_.erase(
            right_analysis_buffer_.begin(),
            right_analysis_buffer_.begin() + static_cast<std::ptrdiff_t>(erase_count));
        analysis_cursor_ -= erase_count;
    }
}

void LiveApp::start_playback() {
    if (captured_samples_.empty()) {
        MessageBoxW(hwnd_, L"Capture some output audio first.", L"gram_live", MB_ICONINFORMATION);
        return;
    }

    loop_playback_ = false;
    stop_playback();

    WAVEFORMATEX format {};
    format.wFormatTag = WAVE_FORMAT_PCM;
    format.nChannels = 1;
    format.nSamplesPerSec = static_cast<DWORD>(analysis_sample_rate_);
    format.wBitsPerSample = 16;
    format.nBlockAlign = static_cast<WORD>(format.nChannels * format.wBitsPerSample / 8);
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;

    if (waveOutOpen(
            &wave_out_,
            WAVE_MAPPER,
            &format,
            reinterpret_cast<DWORD_PTR>(hwnd_),
            0,
            CALLBACK_WINDOW) != MMSYSERR_NOERROR) {
        MessageBoxW(hwnd_, L"Unable to open the playback device.", L"gram_live", MB_ICONERROR);
        wave_out_ = nullptr;
        return;
    }

    playback_samples_ = captured_samples_;
    output_header_ = {};
    output_header_.lpData = reinterpret_cast<LPSTR>(playback_samples_.data());
    output_header_.dwBufferLength =
        static_cast<DWORD>(playback_samples_.size() * sizeof(std::int16_t));

    waveOutPrepareHeader(wave_out_, &output_header_, sizeof(output_header_));
    if (waveOutWrite(wave_out_, &output_header_, sizeof(output_header_)) != MMSYSERR_NOERROR) {
        stop_playback();
        MessageBoxW(hwnd_, L"Unable to start playback.", L"gram_live", MB_ICONERROR);
        return;
    }

    playing_ = true;
    update_transport_controls();
    update_status();
    next_status_tick_ = GetTickCount64() + 250;
}

void LiveApp::start_loop_playback() {
    if (captured_samples_.empty()) {
        MessageBoxW(hwnd_, L"Capture some output audio first.", L"gram_live", MB_ICONINFORMATION);
        return;
    }
    if (!marker_in_set_ || !marker_out_set_ || marker_in_sample_ == marker_out_sample_) {
        MessageBoxW(hwnd_, L"Set distinct In and Out markers on the spectrogram first.", L"gram_live", MB_ICONINFORMATION);
        return;
    }

    const std::size_t begin = (std::min)(marker_in_sample_, marker_out_sample_);
    const std::size_t end = (std::max)(marker_in_sample_, marker_out_sample_);
    if (begin >= end || end > captured_samples_.size()) {
        MessageBoxW(hwnd_, L"The current marker range is outside the captured audio.", L"gram_live", MB_ICONERROR);
        return;
    }

    stop_playback();
    loop_playback_ = true;

    WAVEFORMATEX format {};
    format.wFormatTag = WAVE_FORMAT_PCM;
    format.nChannels = 1;
    format.nSamplesPerSec = static_cast<DWORD>(analysis_sample_rate_);
    format.wBitsPerSample = 16;
    format.nBlockAlign = static_cast<WORD>(format.nChannels * format.wBitsPerSample / 8);
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign;

    if (waveOutOpen(
            &wave_out_,
            WAVE_MAPPER,
            &format,
            reinterpret_cast<DWORD_PTR>(hwnd_),
            0,
            CALLBACK_WINDOW) != MMSYSERR_NOERROR) {
        MessageBoxW(hwnd_, L"Unable to open the playback device.", L"gram_live", MB_ICONERROR);
        wave_out_ = nullptr;
        loop_playback_ = false;
        return;
    }

    playback_samples_.assign(captured_samples_.begin() + static_cast<std::ptrdiff_t>(begin),
        captured_samples_.begin() + static_cast<std::ptrdiff_t>(end));
    output_header_ = {};
    output_header_.lpData = reinterpret_cast<LPSTR>(playback_samples_.data());
    output_header_.dwBufferLength =
        static_cast<DWORD>(playback_samples_.size() * sizeof(std::int16_t));

    waveOutPrepareHeader(wave_out_, &output_header_, sizeof(output_header_));
    if (waveOutWrite(wave_out_, &output_header_, sizeof(output_header_)) != MMSYSERR_NOERROR) {
        stop_playback();
        MessageBoxW(hwnd_, L"Unable to start loop playback.", L"gram_live", MB_ICONERROR);
        return;
    }

    playing_ = true;
    update_transport_controls();
    update_status();
    next_status_tick_ = GetTickCount64() + 250;
}

void LiveApp::stop_playback() {
    if (wave_out_ == nullptr) {
        playing_ = false;
        loop_playback_ = false;
        return;
    }

    waveOutReset(wave_out_);
    waveOutUnprepareHeader(wave_out_, &output_header_, sizeof(output_header_));
    waveOutClose(wave_out_);
    wave_out_ = nullptr;
    output_header_ = {};
    playback_samples_.clear();
    playing_ = false;
    loop_playback_ = false;
    update_transport_controls();
    update_status();
    next_status_tick_ = GetTickCount64() + 250;
}

void LiveApp::save_capture() {
    if (captured_samples_.empty()) {
        MessageBoxW(hwnd_, L"Nothing has been captured yet.", L"gram_live", MB_ICONINFORMATION);
        return;
    }

    wchar_t path[MAX_PATH] = L"loopback-capture.wav";
    OPENFILENAMEW dialog {};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFilter = L"Wave Files (*.wav)\0*.wav\0All Files (*.*)\0*.*\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrDefExt = L"wav";
    dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&dialog) == FALSE) {
        return;
    }

    try {
        gram::write_wav_mono_16(narrow_system_path(path), analysis_sample_rate_, captured_samples_);
    } catch (const std::exception& error) {
        const auto message = widen(error.what());
        MessageBoxW(hwnd_, message.c_str(), L"gram_live", MB_ICONERROR);
    }
}

void LiveApp::save_png() {
    wchar_t path[MAX_PATH] = L"gram-live.png";
    OPENFILENAMEW dialog {};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFilter = L"PNG Files (*.png)\0*.png\0All Files (*.*)\0*.*\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrDefExt = L"png";
    dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&dialog) == FALSE) {
        return;
    }

    RECT client {};
    GetClientRect(hwnd_, &client);
    const int width = client.right - client.left;
    const int height = client.bottom - client.top;
    if (width <= 0 || height <= 0) {
        MessageBoxW(hwnd_, L"Nothing is available to export yet.", L"gram_live", MB_ICONINFORMATION);
        return;
    }

    BITMAPINFO bitmap_info {};
    bitmap_info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmap_info.bmiHeader.biWidth = width;
    bitmap_info.bmiHeader.biHeight = -height;
    bitmap_info.bmiHeader.biPlanes = 1;
    bitmap_info.bmiHeader.biBitCount = 32;
    bitmap_info.bmiHeader.biCompression = BI_RGB;

    HDC screen_dc = GetDC(hwnd_);
    HDC memory_dc = CreateCompatibleDC(screen_dc);
    void* pixels = nullptr;
    HBITMAP bitmap = CreateDIBSection(screen_dc, &bitmap_info, DIB_RGB_COLORS, &pixels, nullptr, 0);
    if (screen_dc == nullptr || memory_dc == nullptr || bitmap == nullptr || pixels == nullptr) {
        if (bitmap != nullptr) {
            DeleteObject(bitmap);
        }
        if (memory_dc != nullptr) {
            DeleteDC(memory_dc);
        }
        if (screen_dc != nullptr) {
            ReleaseDC(hwnd_, screen_dc);
        }
        MessageBoxW(hwnd_, L"Unable to create an export bitmap.", L"gram_live", MB_ICONERROR);
        return;
    }

    HGDIOBJ old_bitmap = SelectObject(memory_dc, bitmap);
    paint(memory_dc);

    IWICImagingFactory* factory = nullptr;
    IWICStream* stream = nullptr;
    IWICBitmapEncoder* encoder = nullptr;
    IWICBitmapFrameEncode* frame = nullptr;
    IPropertyBag2* property_bag = nullptr;

    HRESULT result = CoCreateInstance(
        CLSID_WICImagingFactory,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&factory));
    if (SUCCEEDED(result)) {
        result = factory->CreateStream(&stream);
    }
    if (SUCCEEDED(result)) {
        result = stream->InitializeFromFilename(path, GENERIC_WRITE);
    }
    if (SUCCEEDED(result)) {
        result = factory->CreateEncoder(GUID_ContainerFormatPng, nullptr, &encoder);
    }
    if (SUCCEEDED(result)) {
        result = encoder->Initialize(stream, WICBitmapEncoderNoCache);
    }
    if (SUCCEEDED(result)) {
        result = encoder->CreateNewFrame(&frame, &property_bag);
    }
    if (SUCCEEDED(result)) {
        result = frame->Initialize(property_bag);
    }
    if (SUCCEEDED(result)) {
        result = frame->SetSize(static_cast<UINT>(width), static_cast<UINT>(height));
    }
    WICPixelFormatGUID format = GUID_WICPixelFormat32bppBGRA;
    if (SUCCEEDED(result)) {
        result = frame->SetPixelFormat(&format);
    }
    if (SUCCEEDED(result)) {
        result = frame->WritePixels(
            static_cast<UINT>(height),
            static_cast<UINT>(width * 4),
            static_cast<UINT>(width * height * 4),
            static_cast<BYTE*>(pixels));
    }
    if (SUCCEEDED(result)) {
        result = frame->Commit();
    }
    if (SUCCEEDED(result)) {
        result = encoder->Commit();
    }

    if (property_bag != nullptr) {
        property_bag->Release();
    }
    if (frame != nullptr) {
        frame->Release();
    }
    if (encoder != nullptr) {
        encoder->Release();
    }
    if (stream != nullptr) {
        stream->Release();
    }
    if (factory != nullptr) {
        factory->Release();
    }

    SelectObject(memory_dc, old_bitmap);
    DeleteObject(bitmap);
    DeleteDC(memory_dc);
    ReleaseDC(hwnd_, screen_dc);

    if (FAILED(result)) {
        MessageBoxW(hwnd_, L"Unable to save the PNG export.", L"gram_live", MB_ICONERROR);
    }
}

void LiveApp::save_note_csv() {
    if (captured_samples_.size() < static_cast<std::size_t>(settings_.fft_size)) {
        MessageBoxW(hwnd_, L"Capture more audio before exporting note data.", L"gram_live", MB_ICONINFORMATION);
        return;
    }

    wchar_t path[MAX_PATH] = L"gram-live-notes.csv";
    OPENFILENAMEW dialog {};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFilter = L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrDefExt = L"csv";
    dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&dialog) == FALSE) {
        return;
    }

    std::ofstream output(narrow_system_path(path), std::ios::binary);
    if (!output) {
        MessageBoxW(hwnd_, L"Unable to open the CSV file for writing.", L"gram_live", MB_ICONERROR);
        return;
    }

    output << "time_seconds,frequency_hz,note,cents,level_db\n";
    gram::FftPlan plan(settings_.fft_size);
    std::vector<double> frame(settings_.fft_size);
    std::vector<double> bins;
    for (std::size_t start = 0;
         start + static_cast<std::size_t>(settings_.fft_size) <= captured_samples_.size();
         start += static_cast<std::size_t>(settings_.hop_size)) {
        for (int sample = 0; sample < settings_.fft_size; ++sample) {
            frame[static_cast<std::size_t>(sample)] =
                static_cast<double>(captured_samples_[start + static_cast<std::size_t>(sample)]) / 32768.0;
        }
        plan.compute_db(frame.data(), bins, kAnalysisFloorDb, nullptr);
        const auto peaks = collect_top_peaks(bins, 1);
        if (peaks.empty()) {
            continue;
        }
        double cents = 0.0;
        const std::wstring note = describe_note_from_frequency(peaks.front().frequency, &cents);
        output << std::fixed << std::setprecision(5)
               << (static_cast<double>(start) / static_cast<double>(analysis_sample_rate_)) << ","
               << std::setprecision(3) << peaks.front().frequency << ","
               << narrow_system_path(note.c_str()) << ","
               << std::setprecision(2) << cents << ","
               << std::setprecision(2) << peaks.front().level_db << "\n";
    }
}

void LiveApp::save_preset() {
    wchar_t path[MAX_PATH] = L"gram-live.preset";
    OPENFILENAMEW dialog {};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFilter = L"Preset Files (*.preset)\0*.preset\0Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrDefExt = L"preset";
    dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&dialog) == FALSE) {
        return;
    }

    auto text_of = [](HWND control) {
        wchar_t buffer[128] {};
        GetWindowTextW(control, buffer, 127);
        return std::wstring(buffer);
    };

    std::ofstream output(narrow_system_path(path), std::ios::binary);
    if (!output) {
        MessageBoxW(hwnd_, L"Unable to open the preset file for writing.", L"gram_live", MB_ICONERROR);
        return;
    }

    output << "sample_rate=" << narrow_system_path(text_of(sample_rate_combo_).c_str()) << "\n";
    output << "fft=" << narrow_system_path(text_of(fft_combo_).c_str()) << "\n";
    output << "palette=" << narrow_system_path(text_of(palette_combo_).c_str()) << "\n";
    output << "scale=" << narrow_system_path(text_of(scale_combo_).c_str()) << "\n";
    output << "max_freq=" << narrow_system_path(text_of(max_freq_combo_).c_str()) << "\n";
    output << "layout=" << narrow_system_path(text_of(channel_combo_).c_str()) << "\n";
    output << "display=" << narrow_system_path(text_of(display_combo_).c_str()) << "\n";
    output << "amplitude=" << narrow_system_path(text_of(amplitude_combo_).c_str()) << "\n";
    output << "mono_channel=" << narrow_system_path(text_of(mono_channel_combo_).c_str()) << "\n";
    output << "history=" << narrow_system_path(text_of(history_combo_).c_str()) << "\n";
    output << "scroll=" << narrow_system_path(text_of(scroll_speed_combo_).c_str()) << "\n";
    output << "sensitivity=" << static_cast<int>(SendMessageW(max_db_slider_, TBM_GETPOS, 0, 0)) << "\n";
    output << "grid=" << (SendMessageW(grid_check_, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    output << "peak_hold=" << (SendMessageW(peak_hold_check_, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    output << "average=" << (SendMessageW(average_check_, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
}

void LiveApp::load_preset() {
    wchar_t path[MAX_PATH] = L"";
    OPENFILENAMEW dialog {};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = hwnd_;
    dialog.lpstrFilter = L"Preset Files (*.preset)\0*.preset\0Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = MAX_PATH;
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

    if (GetOpenFileNameW(&dialog) == FALSE) {
        return;
    }

    std::ifstream input(narrow_system_path(path), std::ios::binary);
    if (!input) {
        MessageBoxW(hwnd_, L"Unable to open the preset file.", L"gram_live", MB_ICONERROR);
        return;
    }

    std::string line;
    while (std::getline(input, line)) {
        const auto separator = line.find('=');
        if (separator == std::string::npos) {
            continue;
        }
        const std::string key = line.substr(0, separator);
        const std::wstring value = widen(line.substr(separator + 1).c_str());
        if (key == "sample_rate") {
            set_combo_to_text(sample_rate_combo_, value.c_str());
        } else if (key == "fft") {
            set_combo_to_text(fft_combo_, value.c_str());
        } else if (key == "palette") {
            set_combo_to_text(palette_combo_, value.c_str());
        } else if (key == "scale") {
            set_combo_to_text(scale_combo_, value.c_str());
        } else if (key == "max_freq") {
            set_combo_to_text(max_freq_combo_, value.c_str());
        } else if (key == "layout") {
            set_combo_to_text(channel_combo_, value.c_str());
        } else if (key == "display") {
            set_combo_to_text(display_combo_, value.c_str());
        } else if (key == "amplitude") {
            set_combo_to_text(amplitude_combo_, value.c_str());
        } else if (key == "mono_channel") {
            set_combo_to_text(mono_channel_combo_, value.c_str());
        } else if (key == "history") {
            set_combo_to_text(history_combo_, value.c_str());
        } else if (key == "scroll") {
            set_combo_to_text(scroll_speed_combo_, value.c_str());
        } else if (key == "sensitivity") {
            SendMessageW(max_db_slider_, TBM_SETPOS, TRUE, static_cast<LPARAM>(_wtoi(value.c_str())));
        } else if (key == "grid") {
            SendMessageW(grid_check_, BM_SETCHECK, _wtoi(value.c_str()) != 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        } else if (key == "peak_hold") {
            SendMessageW(peak_hold_check_, BM_SETCHECK, _wtoi(value.c_str()) != 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        } else if (key == "average") {
            SendMessageW(average_check_, BM_SETCHECK, _wtoi(value.c_str()) != 0 ? BST_CHECKED : BST_UNCHECKED, 0);
        }
    }

    update_level_labels();
    mark_settings_dirty();
}

void LiveApp::update_status() {
    const double captured_seconds =
        analysis_sample_rate_ > 0
            ? static_cast<double>(captured_samples_.size()) / static_cast<double>(analysis_sample_rate_)
            : 0.0;
    const MusicStatusFields music = current_music_status();
    std::wstring corr_text;
    if (use_stereo_display()) {
        std::wostringstream corr_stream;
        corr_stream << std::fixed << std::setprecision(2) << current_stereo_correlation();
        corr_text = corr_stream.str();
    }

    std::wostringstream summary;
    summary << std::fixed << std::setprecision(1)
            << L"Source: default output | Capture: " << (capturing_ ? L"running" : L"stopped")
            << L" | Playback: " << (playing_ ? L"running" : L"idle");
    const std::wstring next_status = summary.str();
    if (next_status != status_text_) {
        status_text_ = next_status;
    }
    if (status_bar_ != nullptr) {
        std::vector<std::wstring> parts_text;
        std::vector<std::wstring> parts_measure;
        auto push_part = [&](std::wstring text, std::wstring measure = {}) {
            parts_text.push_back(std::move(text));
            parts_measure.push_back(measure.empty() ? parts_text.back() : std::move(measure));
        };

        push_part(L"Source: default output loopback");
        push_part(last_stream_sample_rate_ > 0 ? (L"Stream: " + std::to_wstring(last_stream_sample_rate_) + L" Hz") : L"",
            L"Stream: 384000 Hz");
        if (use_stereo_display()) {
            push_part(L"Channel: Stereo", L"Channel: Mono Right");
        } else if (mono_channel_mode_ == MonoChannelMode::Left) {
            push_part(L"Channel: Mono Left", L"Channel: Mono Right");
        } else if (mono_channel_mode_ == MonoChannelMode::Right) {
            push_part(L"Channel: Mono Right", L"Channel: Mono Right");
        } else {
            push_part(L"Channel: Mono Mix", L"Channel: Mono Right");
        }
        push_part(freeze_display_ ? L"Display: frozen" : L"Display: live", L"Display: frozen");
        push_part(playing_ ? (loop_playback_ ? L"Playback: looping" : L"Playback: running") : L"Playback: idle",
            L"Playback: looping");
        {
            std::wostringstream captured;
            captured << std::fixed << std::setprecision(1) << captured_seconds << L"s";
            push_part(L"Captured: " + captured.str(), L"Captured: 9999.9s");
        }
        {
            std::wostringstream history;
            history << std::fixed << std::setprecision(0) << history_seconds_ << L"s";
            push_part(L"Hist: " + history.str(), L"Hist: 999s");
        }
        {
            std::wostringstream fft_avg;
            fft_avg << std::fixed << std::setprecision(3) << profiling_.average_fft_ms << L" ms";
            push_part(L"FFT: " + fft_avg.str(), L"FFT: 999.999 ms");
        }
        push_part(music.pitch.empty() ? L"" : L"Pitch: " + music.pitch, L"Pitch: C#10");
        push_part(music.cents.empty() ? L"" : L"Cents: " + music.cents, L"Cents: +100.0 cents");
        push_part(music.chord.empty() ? L"" : L"Chord: " + music.chord, L"Chord: C# major");
        push_part(corr_text.empty() ? L"" : L"Corr: " + corr_text, L"Corr: -1.00");
        push_part(trim_for_status(now_playing_text_, 72), fixed_status_text(L"Now playing: " + std::wstring(56, L'W'), 72));
        push_part(cursor_text_, L"Cursor: 30000 Hz, -180.0 dB");
        if (marker_in_set_ || marker_out_set_) {
            std::wostringstream marks;
            marks << L"Marks:";
            if (marker_in_set_ && analysis_sample_rate_ > 0) {
                marks << L" In " << std::fixed << std::setprecision(2)
                      << static_cast<double>(marker_in_sample_) / static_cast<double>(analysis_sample_rate_) << L"s";
            }
            if (marker_out_set_ && analysis_sample_rate_ > 0) {
                marks << L" Out " << std::fixed << std::setprecision(2)
                      << static_cast<double>(marker_out_sample_) / static_cast<double>(analysis_sample_rate_) << L"s";
            }
            push_part(marks.str(), L"Marks: In 9999.99s Out 9999.99s");
        } else {
            push_part(L"", L"Marks: In 9999.99s Out 9999.99s");
        }
        push_part(settings_dirty_ ? L"Apply pending" : L"", L"Apply pending");

        HDC status_dc = GetDC(status_bar_);
        if (status_dc != nullptr) {
            HFONT font = reinterpret_cast<HFONT>(SendMessageW(status_bar_, WM_GETFONT, 0, 0));
            HGDIOBJ old_font = nullptr;
            if (font != nullptr) {
                old_font = SelectObject(status_dc, font);
            }
            std::vector<int> parts;
            parts.reserve(parts_text.size());
            int current = 0;
            for (const auto& text : parts_measure) {
                SIZE size {};
                GetTextExtentPoint32W(status_dc, text.c_str(), static_cast<int>(text.size()), &size);
                current += size.cx + 18;
                parts.push_back(current);
            }
            if (!parts.empty()) {
                parts.back() = -1;
                SendMessageW(status_bar_, SB_SETPARTS, static_cast<WPARAM>(parts.size()), reinterpret_cast<LPARAM>(parts.data()));
            }
            if (old_font != nullptr) {
                SelectObject(status_dc, old_font);
            }
            ReleaseDC(status_bar_, status_dc);
        }
        for (std::size_t index = 0; index < parts_text.size(); ++index) {
            SendMessageW(status_bar_, SB_SETTEXTW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(parts_text[index].c_str()));
        }
    }
}

void LiveApp::update_level_labels() {
    const int sensitivity = static_cast<int>(SendMessageW(max_db_slider_, TBM_GETPOS, 0, 0));
    std::wostringstream max_label;
    max_label << L"Sensitivity: " << sensitivity << L" dB";
    if (max_db_label_ != nullptr) {
        SetWindowTextW(max_db_label_, max_label.str().c_str());
    }
}

void LiveApp::mark_settings_dirty() {
    update_level_labels();
    settings_dirty_ = true;
    if (apply_button_ != nullptr) {
        EnableWindow(apply_button_, TRUE);
    }
    update_channel_controls();
    update_view_controls();
    update_status();
}

void LiveApp::apply_live_settings_change(int control_id) {
    if (control_requires_apply(control_id) || settings_dirty_) {
        mark_settings_dirty();
        return;
    }
    update_level_labels();
    apply_settings_with_diagnostics(this);
}

void LiveApp::update_channel_controls() {
    if (mono_channel_combo_ == nullptr || mono_channel_label_ == nullptr) {
        return;
    }

    wchar_t pending_layout[32] {};
    GetWindowTextW(channel_combo_, pending_layout, 31);
    bool stereo_mode = false;
    if (last_stream_channels_ >= 2) {
        stereo_mode = wcscmp(pending_layout, L"mono") != 0;
    }

    if (stereo_mode) {
        set_combo_to_text(mono_channel_combo_, L"Stereo");
        EnableWindow(mono_channel_combo_, FALSE);
    } else {
        if (mono_channel_mode_ == MonoChannelMode::Left) {
            set_combo_to_text(mono_channel_combo_, L"left");
        } else if (mono_channel_mode_ == MonoChannelMode::Right) {
            set_combo_to_text(mono_channel_combo_, L"right");
        } else {
            set_combo_to_text(mono_channel_combo_, L"mix");
        }
        EnableWindow(mono_channel_combo_, TRUE);
    }
}

void LiveApp::update_transport_controls() {
    if (start_button_ != nullptr) {
        EnableWindow(start_button_, capturing_ ? FALSE : TRUE);
    }
    if (stop_button_ != nullptr) {
        EnableWindow(stop_button_, (capturing_ || playing_) ? TRUE : FALSE);
    }
    if (!playing_ && freeze_display_) {
        freeze_display_ = false;
        if (freeze_button_ != nullptr) {
            SetWindowTextW(freeze_button_, L"Freeze");
        }
        graph_dirty_ = true;
    }
    if (freeze_button_ != nullptr) {
        EnableWindow(freeze_button_, playing_ ? TRUE : FALSE);
    }
}

void LiveApp::update_view_controls() {
    wchar_t display_text[32] {};
    if (display_combo_ != nullptr) {
        GetWindowTextW(display_combo_, display_text, 31);
    }
    const bool spectrum_mode = wcscmp(display_text, L"spectrum") == 0;
    ShowWindow(peak_hold_check_, spectrum_mode ? SW_SHOW : SW_HIDE);
    ShowWindow(average_check_, spectrum_mode ? SW_SHOW : SW_HIDE);
    update_tuner_controls();
}

void LiveApp::update_tuner_controls() {
    wchar_t display_text[32] {};
    if (display_combo_ != nullptr) {
        GetWindowTextW(display_combo_, display_text, 31);
    }
    const bool tuner_mode = wcscmp(display_text, L"tuner") == 0;
    const bool tuner_idle = !capturing_;
    ShowWindow(instrument_label_, tuner_mode ? SW_SHOW : SW_HIDE);
    ShowWindow(instrument_combo_, tuner_mode ? SW_SHOW : SW_HIDE);
    ShowWindow(tuner_start_button_, SW_HIDE);
    ShowWindow(tuner_stop_button_, SW_HIDE);
    if (instrument_combo_ != nullptr) {
        EnableWindow(instrument_combo_, tuner_idle ? TRUE : FALSE);
    }
}

bool LiveApp::control_requires_apply(int control_id) const {
    switch (control_id) {
    case SampleRateCombo:
    case FftCombo:
    case ScaleCombo:
    case MinFreqCombo:
    case MaxFreqCombo:
        return true;
    default:
        return false;
    }
}

void LiveApp::update_now_playing() {
    try {
        if (media_session_manager_ == nullptr) {
            media_session_manager_ =
                winrt::Windows::Media::Control::GlobalSystemMediaTransportControlsSessionManager::RequestAsync().get();
        }

        auto session = media_session_manager_.GetCurrentSession();
        if (session == nullptr) {
            now_playing_text_.clear();
            return;
        }

        auto properties = session.TryGetMediaPropertiesAsync().get();
        std::wstring title = properties.Title().c_str();
        std::wstring artist = properties.Artist().c_str();
        if (!title.empty() && !artist.empty()) {
            now_playing_text_ = L"Now playing: " + artist + L" - " + title;
        } else if (!title.empty()) {
            now_playing_text_ = L"Now playing: " + title;
        } else {
            now_playing_text_.clear();
        }
    } catch (...) {
        now_playing_text_.clear();
    }
}

void LiveApp::update_cursor_readout(POINT point, bool force_clear) {
    std::wstring next_text = L"Cursor: off";
    cursor_sample_valid_ = false;
    if (!force_clear) {
        const bool spectrum_like = display_mode_ == DisplayMode::Spectrum || display_mode_ == DisplayMode::Room;
        if (display_mode_ == DisplayMode::Tuner || display_mode_ == DisplayMode::Vectorscope) {
            if (display_mode_ != DisplayMode::Room) {
                if (cursor_text_ != next_text) {
                    cursor_text_ = next_text;
                    update_status();
                }
                return;
            }
        }
        RECT plot = plot_rect();
        bool inside_plot = false;

        if (display_mode_ == DisplayMode::Spectrogram && use_stereo_display()) {
            const RECT left_plot = spectrogram_pane_rect(plot, 0);
            const RECT right_plot = spectrogram_pane_rect(plot, 1);
            if (PtInRect(&left_plot, point)) {
                plot = left_plot;
                inside_plot = true;
            } else if (PtInRect(&right_plot, point)) {
                plot = right_plot;
                inside_plot = true;
            }
        } else {
            inside_plot = PtInRect(&plot, point) != FALSE;
        }

        const int plot_width = plot.right - plot.left;
        const int plot_height = plot.bottom - plot.top;
        if (inside_plot && plot_width > 1 && plot_height > 1) {
            const double normalized_x = std::clamp(
                static_cast<double>(point.x - plot.left) /
                    static_cast<double>((std::max)(plot_width - 1, 1)),
                0.0,
                1.0);
            const double normalized_y = std::clamp(
                static_cast<double>(plot.bottom - 1 - point.y) /
                    static_cast<double>((std::max)(plot_height - 1, 1)),
                0.0,
                1.0);
            const double frequency =
                spectrum_like ? frequency_for_position(normalized_x) : frequency_for_position(normalized_y);

            if (spectrum_like) {
                std::wostringstream stream;
                stream << L"Cursor: " << format_frequency_label(frequency) << L" | ";
                if (use_stereo_display()) {
                    const double left_db = spectrum_level_at_position(
                        freeze_display_ ? frozen_db_bins_left_ : db_bins_left_,
                        normalized_x);
                    const double right_db = spectrum_level_at_position(
                        freeze_display_ ? frozen_db_bins_right_ : db_bins_right_,
                        normalized_x);
                    stream << L"L " << format_db_text(left_db)
                           << L" | R " << format_db_text(right_db);
                } else {
                    const std::vector<double>* bins = freeze_display_ ? &frozen_db_bins_ : &db_bins_;
                    if (mono_channel_mode_ == MonoChannelMode::Left) {
                        bins = freeze_display_ ? &frozen_db_bins_left_ : &db_bins_left_;
                    } else if (mono_channel_mode_ == MonoChannelMode::Right) {
                        bins = freeze_display_ ? &frozen_db_bins_right_ : &db_bins_right_;
                    }
                    stream << format_db_text(spectrum_level_at_position(*bins, normalized_x));
                }
                next_text = stream.str();
            } else {
                const double seconds_visible = visible_history_seconds();
                const double seconds_ago = (1.0 - normalized_x) * seconds_visible;
                if (analysis_sample_rate_ > 0 && captured_samples_.size() > 0) {
                    const double sample_offset = seconds_ago * static_cast<double>(analysis_sample_rate_);
                    const auto rounded = static_cast<long long>(std::llround(sample_offset));
                    const auto clamped = std::clamp<long long>(
                        static_cast<long long>(captured_samples_.size()) - rounded,
                        0,
                        static_cast<long long>(captured_samples_.size()));
                    cursor_sample_index_ = static_cast<std::size_t>(clamped);
                    cursor_sample_valid_ = true;
                }
                std::wostringstream stream;
                stream << L"Cursor: " << format_frequency_label(frequency)
                       << L" | " << std::fixed << std::setprecision(2)
                       << seconds_ago << L"s ago";
                next_text = stream.str();
            }
        }
    }

    if (next_text != cursor_text_) {
        cursor_text_ = next_text;
        update_status();
    }
}

void LiveApp::capture_freeze_snapshot() {
    frozen_image_ = image_;
    frozen_left_image_ = left_image_;
    frozen_right_image_ = right_image_;
    frozen_db_bins_ = db_bins_;
    frozen_db_bins_left_ = db_bins_left_;
    frozen_db_bins_right_ = db_bins_right_;
    frozen_peak_hold_bins_ = peak_hold_bins_;
    frozen_peak_hold_bins_left_ = peak_hold_bins_left_;
    frozen_peak_hold_bins_right_ = peak_hold_bins_right_;
    frozen_average_bins_ = average_bins_;
    frozen_average_bins_left_ = average_bins_left_;
    frozen_average_bins_right_ = average_bins_right_;
    reference_available_ = true;
}

double LiveApp::base_history_seconds() const {
    if (analysis_sample_rate_ <= 0 || settings_.hop_size <= 0 || image_.width <= 1) {
        return 0.0;
    }
    return (static_cast<double>(image_.width) * static_cast<double>(settings_.hop_size)) /
        static_cast<double>(analysis_sample_rate_);
}

double LiveApp::effective_scroll_multiplier() const {
    const double base_seconds = base_history_seconds();
    if (base_seconds <= 0.0) {
        return 1.0;
    }
    const double target = std::clamp(history_seconds_, 5.0, 60.0);
    const double speed = std::clamp(scroll_speed_, 0.5, 4.0);
    return std::clamp((base_seconds / target) * speed, 0.0625, 16.0);
}

double LiveApp::visible_history_seconds() const {
    const double base_seconds = base_history_seconds();
    const double multiplier = effective_scroll_multiplier();
    if (base_seconds <= 0.0 || multiplier <= 0.0) {
        return base_seconds;
    }
    return base_seconds / multiplier;
}

double LiveApp::normalized_for_frequency(double frequency) const {
    const double max_frequency = max_display_frequency();
    const double min_frequency = std::clamp(settings_.min_frequency, 0.0, max_frequency);
    if (frequency <= 0.0 || max_frequency <= 0.0) {
        return 0.0;
    }
    if (settings_.frequency_scale == gram::FrequencyScale::Linear) {
        const double span = (std::max)(max_frequency - min_frequency, 1.0);
        return std::clamp((frequency - min_frequency) / span, 0.0, 1.0);
    }
    if (settings_.frequency_scale == gram::FrequencyScale::Logarithmic) {
        const double min_log_freq = std::max(min_frequency, 1.0);
        const double low = std::log10(min_log_freq);
        const double high = std::log10((std::max)(max_frequency, min_log_freq));
        return std::clamp((std::log10((std::max)(frequency, min_log_freq)) - low) / (high - low), 0.0, 1.0);
    }
    const double min_octave_freq = std::max(min_frequency, 1.0);
    const double top = (std::max)(max_frequency, min_octave_freq);
    return std::clamp(std::log2((std::max)(frequency, min_octave_freq) / min_octave_freq) / std::log2(top / min_octave_freq), 0.0, 1.0);
}

const std::vector<double>& LiveApp::selected_bins(bool frozen) const {
    if (use_stereo_display()) {
        return frozen ? frozen_db_bins_ : db_bins_;
    }
    if (mono_channel_mode_ == MonoChannelMode::Left) {
        return frozen ? frozen_db_bins_left_ : db_bins_left_;
    }
    if (mono_channel_mode_ == MonoChannelMode::Right) {
        return frozen ? frozen_db_bins_right_ : db_bins_right_;
    }
    return frozen ? frozen_db_bins_ : db_bins_;
}

std::vector<SpectrumPeak> LiveApp::collect_top_peaks(const std::vector<double>& bins, int count) const {
    std::vector<SpectrumPeak> peaks;
    if (bins.size() < 3 || analysis_sample_rate_ <= 0) {
        return peaks;
    }

    const double nyquist = analysis_sample_rate_ / 2.0;
    const double min_peak_frequency = std::max(settings_.min_frequency, 1.0);
    for (std::size_t index = 1; index + 1 < bins.size(); ++index) {
        const double level = bins[index];
        if (level < settings_.min_db + 12.0) {
            continue;
        }
        if (level < bins[index - 1] || level < bins[index + 1]) {
            continue;
        }
        const double frequency = (static_cast<double>(index) / static_cast<double>(bins.size() - 1)) * nyquist;
        if (frequency < min_peak_frequency || frequency > max_display_frequency()) {
            continue;
        }
        SpectrumPeak peak;
        peak.frequency = frequency;
        peak.level_db = level;
        double cents = 0.0;
        peak.midi_note = midi_from_frequency(frequency);
        describe_note_from_frequency(frequency, &cents);
        peak.cents = cents;
        peaks.push_back(peak);
    }

    std::sort(peaks.begin(), peaks.end(), [](const SpectrumPeak& left, const SpectrumPeak& right) {
        return left.level_db > right.level_db;
    });
    if (static_cast<int>(peaks.size()) > count) {
        peaks.resize(static_cast<std::size_t>(count));
    }
    return peaks;
}

double LiveApp::current_stereo_correlation() const {
    return stereo_correlation_;
}

MusicStatusFields LiveApp::current_music_status() const {
    MusicStatusFields fields;
    const auto peaks = collect_top_peaks(selected_bins(freeze_display_), use_stereo_display() ? 4 : 5);
    if (peaks.empty()) {
        return fields;
    }

    std::vector<int> midi_notes;
    midi_notes.reserve(peaks.size());
    for (const auto& peak : peaks) {
        midi_notes.push_back(peak.midi_note);
    }

    double cents = 0.0;
    fields.pitch = describe_note_from_frequency(peaks.front().frequency, &cents);
    fields.cents = format_signed_cents(cents);
    fields.chord = chord_hint_from_midis(midi_notes);
    return fields;
}

std::wstring LiveApp::current_music_summary() const {
    const MusicStatusFields fields = current_music_status();
    if (fields.pitch.empty()) {
        return L"";
    }
    std::wostringstream stream;
    stream << L"Music: " << fields.pitch;
    if (!fields.cents.empty()) {
        stream << L" " << fields.cents;
    }
    if (!fields.chord.empty()) {
        stream << L" | " << fields.chord;
    }
    if (use_stereo_display()) {
        stream << L" | Corr " << std::fixed << std::setprecision(2) << current_stereo_correlation();
        if (current_stereo_correlation() < -0.2) {
            stream << L" mono risk";
        }
    }
    if (!transient_markers_.empty() && analysis_sample_rate_ > 0 && transient_markers_.size() >= 4) {
        const std::size_t last = transient_markers_.back();
        const std::size_t prev = transient_markers_[transient_markers_.size() - 4];
        const double seconds = static_cast<double>(last - prev) / static_cast<double>(analysis_sample_rate_);
        if (seconds > 0.1) {
            const double bpm = 60.0 * 3.0 / seconds;
            stream << L" | BPM ~" << std::fixed << std::setprecision(0) << bpm;
        }
    }
    return stream.str();
}

std::vector<std::wstring> LiveApp::active_instrument_notes() const {
    if (instrument_mode_ == L"guitar") {
        return { L"E2", L"A2", L"D3", L"G3", L"B3", L"E4" };
    }
    if (instrument_mode_ == L"bass") {
        return { L"E1", L"A1", L"D2", L"G2" };
    }
    if (instrument_mode_ == L"ukulele") {
        return { L"G4", L"C4", L"E4", L"A4" };
    }
    if (instrument_mode_ == L"violin") {
        return { L"G3", L"D4", L"A4", L"E5" };
    }
    if (instrument_mode_ == L"viola") {
        return { L"C3", L"G3", L"D4", L"A4" };
    }
    if (instrument_mode_ == L"cello") {
        return { L"C2", L"G2", L"D3", L"A3" };
    }
    if (instrument_mode_ == L"mandolin") {
        return { L"G3", L"D4", L"A4", L"E5" };
    }
    return {};
}

std::size_t LiveApp::marker_sample_from_cursor() const {
    if (cursor_sample_valid_) {
        return cursor_sample_index_;
    }
    return captured_samples_.size();
}

void LiveApp::draw_spectrogram_markers(HDC dc, const RECT& plot) {
    if (!show_beat_marks_ && !marker_in_set_ && !marker_out_set_) {
        return;
    }
    if (captured_samples_.empty() || analysis_sample_rate_ <= 0) {
        return;
    }

    const double visible_seconds = visible_history_seconds();
    if (visible_seconds <= 0.0) {
        return;
    }

    HPEN in_pen = CreatePen(PS_SOLID, 1, RGB(80, 255, 150));
    HPEN out_pen = CreatePen(PS_SOLID, 1, RGB(255, 120, 120));
    SetBkMode(dc, TRANSPARENT);

    auto draw_marker = [&](std::size_t sample_index, const wchar_t* label, HPEN pen, COLORREF text_color) {
        if (sample_index > captured_samples_.size()) {
            return;
        }
        const double samples_ago = static_cast<double>(captured_samples_.size() - sample_index);
        const double seconds_ago = samples_ago / static_cast<double>(analysis_sample_rate_);
        if (seconds_ago < 0.0 || seconds_ago > visible_seconds) {
            return;
        }
        const double normalized = 1.0 - (seconds_ago / visible_seconds);
        const int x = plot.left + static_cast<int>(std::lround(normalized * static_cast<double>(plot.right - plot.left - 1)));
        HGDIOBJ old_pen = SelectObject(dc, pen);
        MoveToEx(dc, x, plot.top, nullptr);
        LineTo(dc, x, plot.bottom);
        SelectObject(dc, old_pen);
        SetTextColor(dc, text_color);
        TextOutW(dc, x + 3, plot.top + 6, label, static_cast<int>(wcslen(label)));
    };

    if (marker_in_set_) {
        draw_marker(marker_in_sample_, L"In", in_pen, RGB(80, 255, 150));
    }
    if (marker_out_set_) {
        draw_marker(marker_out_sample_, L"Out", out_pen, RGB(255, 120, 120));
    }

    HPEN transient_pen = CreatePen(PS_DOT, 1, RGB(255, 220, 80));
    if (show_beat_marks_) {
        for (const std::size_t sample_index : transient_markers_) {
            if (sample_index > captured_samples_.size()) {
                continue;
            }
            const double samples_ago = static_cast<double>(captured_samples_.size() - sample_index);
            const double seconds_ago = samples_ago / static_cast<double>(analysis_sample_rate_);
            if (seconds_ago < 0.0 || seconds_ago > visible_seconds) {
                continue;
            }
            const double normalized = 1.0 - (seconds_ago / visible_seconds);
            const int x = plot.left + static_cast<int>(std::lround(normalized * static_cast<double>(plot.right - plot.left - 1)));
            HGDIOBJ old_pen = SelectObject(dc, transient_pen);
            MoveToEx(dc, x, plot.top, nullptr);
            LineTo(dc, x, plot.bottom);
            SelectObject(dc, old_pen);
        }
    }

    DeleteObject(in_pen);
    DeleteObject(out_pen);
    DeleteObject(transient_pen);
}

void LiveApp::draw_note_guides(HDC dc, const RECT& plot, bool horizontal_axis) {
    if (!show_note_guides_) {
        return;
    }
    const double max_frequency = max_display_frequency();
    if (max_frequency <= 40.0) {
        return;
    }

    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(110, 110, 110));
    HPEN note_pen = CreatePen(PS_DOT, 1, RGB(70, 70, 70));
    HGDIOBJ old_pen = SelectObject(dc, note_pen);

    for (int midi = 24; midi <= 120; ++midi) {
        const double frequency = frequency_from_midi(midi);
        if (frequency < 20.0 || frequency > max_frequency) {
            continue;
        }
        const double normalized = normalized_for_frequency(frequency);
        if (horizontal_axis) {
            const int x = plot.left + static_cast<int>(std::lround(normalized * static_cast<double>(plot.right - plot.left - 1)));
            MoveToEx(dc, x, plot.top, nullptr);
            LineTo(dc, x, plot.bottom);
            if ((midi % 12) == 0) {
                const std::wstring name = note_name_from_midi(midi);
                TextOutW(dc, x + 2, plot.bottom + 22, name.c_str(), static_cast<int>(name.size()));
            }
        } else {
            const int y = plot.bottom - 1 - static_cast<int>(std::lround(normalized * static_cast<double>(plot.bottom - plot.top - 1)));
            MoveToEx(dc, plot.left, y, nullptr);
            LineTo(dc, plot.right, y);
            if ((midi % 12) == 0) {
                const std::wstring name = note_name_from_midi(midi);
                TextOutW(dc, plot.left + 4, y - 8, name.c_str(), static_cast<int>(name.size()));
            }
        }
    }

    SelectObject(dc, old_pen);
    DeleteObject(note_pen);
}

void LiveApp::draw_peak_labels(HDC dc, const RECT& plot, const std::vector<SpectrumPeak>& peaks) {
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(240, 240, 200));
    for (const auto& peak : peaks) {
        const double normalized_x = normalized_for_frequency(peak.frequency);
        const int x = plot.left + static_cast<int>(std::lround(normalized_x * static_cast<double>(plot.right - plot.left - 1)));
        const int y = plot.top + 10 + static_cast<int>(&peak - peaks.data()) * 16;
        const std::wstring name = note_name_from_midi(peak.midi_note);
        TextOutW(dc, x + 2, y, name.c_str(), static_cast<int>(name.size()));
    }
}

void LiveApp::draw_tuner(HDC dc, const RECT& outer, const RECT& plot) {
    draw_spectrum_axes(dc, outer, plot);
    const auto peaks = collect_top_peaks(selected_bins(freeze_display_), 5);
    if (peaks.empty()) {
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, RGB(210, 210, 210));
        const wchar_t* text = L"No stable pitch";
        TextOutW(dc, plot.left + 24, plot.top + 24, text, static_cast<int>(wcslen(text)));
        return;
    }

    double cents = 0.0;
    const std::wstring note = describe_note_from_frequency(peaks.front().frequency, &cents);
    const std::wstring cents_text = format_signed_cents(cents);
    const std::wstring freq_text = format_frequency_label(peaks.front().frequency);

    RECT note_rect = plot;
    note_rect.top += 20;
    HFONT large_font = CreateFontW(84, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    HFONT small_font = CreateFontW(26, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    HGDIOBJ old_font = SelectObject(dc, large_font);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(255, 240, 170));
    DrawTextW(dc, note.c_str(), -1, &note_rect, DT_CENTER | DT_TOP | DT_SINGLELINE);

    SelectObject(dc, small_font);
    RECT info_rect = plot;
    info_rect.top += 118;
    std::wstring info = freq_text + L" | " + cents_text;
    DrawTextW(dc, info.c_str(), -1, &info_rect, DT_CENTER | DT_TOP | DT_SINGLELINE);

    const int bar_y = plot.top + 170;
    const int bar_left = plot.left + 40;
    const int bar_right = plot.right - 40;
    HPEN center_pen = CreatePen(PS_SOLID, 2, RGB(200, 200, 200));
    HPEN needle_pen = CreatePen(PS_SOLID, 3, std::abs(cents) < 5.0 ? RGB(80, 255, 120) : RGB(255, 120, 120));
    HGDIOBJ old_pen = SelectObject(dc, center_pen);
    MoveToEx(dc, bar_left, bar_y, nullptr);
    LineTo(dc, bar_right, bar_y);
    const int center_x = (bar_left + bar_right) / 2;
    MoveToEx(dc, center_x, bar_y - 16, nullptr);
    LineTo(dc, center_x, bar_y + 16);
    SelectObject(dc, needle_pen);
    const double clamped_cents = std::clamp(cents, -50.0, 50.0);
    const int needle_x = center_x + static_cast<int>(std::lround((clamped_cents / 50.0) * static_cast<double>((bar_right - bar_left) / 2)));
    MoveToEx(dc, needle_x, bar_y - 24, nullptr);
    LineTo(dc, needle_x, bar_y + 24);
    SelectObject(dc, old_pen);
    DeleteObject(center_pen);
    DeleteObject(needle_pen);

    RECT harm_rect = plot;
    harm_rect.top += 210;
    std::wostringstream harmonics;
    harmonics << L"Harmonics:";
    for (int harmonic = 2; harmonic <= 6; ++harmonic) {
        harmonics << L" " << harmonic << L"x "
                  << format_frequency_label(peaks.front().frequency * harmonic);
    }
    DrawTextW(dc, harmonics.str().c_str(), -1, &harm_rect, DT_CENTER | DT_TOP | DT_SINGLELINE);

    const auto strings = active_instrument_notes();
    if (!strings.empty()) {
        RECT strings_rect = plot;
        strings_rect.top += 242;
        std::wostringstream stream;
        stream << L"Targets:";
        for (const auto& string_note : strings) {
            stream << L" " << string_note;
        }
        DrawTextW(dc, stream.str().c_str(), -1, &strings_rect, DT_CENTER | DT_TOP | DT_SINGLELINE);
    }

    SelectObject(dc, old_font);
    DeleteObject(large_font);
    DeleteObject(small_font);
}

void LiveApp::draw_vectorscope(HDC dc, const RECT& outer, const RECT& plot) {
    (void)outer;
    HPEN axis_pen = CreatePen(PS_SOLID, 1, RGB(96, 96, 96));
    HGDIOBJ old_pen = SelectObject(dc, axis_pen);
    HGDIOBJ old_brush = SelectObject(dc, GetStockObject(NULL_BRUSH));
    Rectangle(dc, plot.left, plot.top, plot.right, plot.bottom);
    const int center_x = (plot.left + plot.right) / 2;
    const int center_y = (plot.top + plot.bottom) / 2;
    MoveToEx(dc, center_x, plot.top, nullptr);
    LineTo(dc, center_x, plot.bottom);
    MoveToEx(dc, plot.left, center_y, nullptr);
    LineTo(dc, plot.right, center_y);

    if (last_stream_channels_ < 2 || captured_left_samples_.empty()) {
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, RGB(210, 210, 210));
        const wchar_t* text = L"Vectorscope requires stereo audio";
        TextOutW(dc, plot.left + 20, plot.top + 20, text, static_cast<int>(wcslen(text)));
    } else {
        HPEN trace_pen = CreatePen(PS_SOLID, 1, RGB(120, 240, 255));
        SelectObject(dc, trace_pen);
        const std::size_t take = (std::min)(captured_left_samples_.size(), static_cast<std::size_t>(4096));
        const std::size_t start = captured_left_samples_.size() - take;
        bool first = true;
        for (std::size_t index = start; index < captured_left_samples_.size(); index += 8) {
            const double left = static_cast<double>(captured_left_samples_[index]) / 32768.0;
            const double right = static_cast<double>(captured_right_samples_[index]) / 32768.0;
            const int x = center_x + static_cast<int>(std::lround(left * static_cast<double>((plot.right - plot.left) / 2 - 12)));
            const int y = center_y - static_cast<int>(std::lround(right * static_cast<double>((plot.bottom - plot.top) / 2 - 12)));
            if (first) {
                MoveToEx(dc, x, y, nullptr);
                first = false;
            } else {
                LineTo(dc, x, y);
            }
        }
        DeleteObject(trace_pen);

        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, current_stereo_correlation() < -0.2 ? RGB(255, 120, 120) : RGB(210, 210, 210));
        std::wostringstream corr;
        corr << L"Correlation: " << std::fixed << std::setprecision(2) << current_stereo_correlation();
        const std::wstring corr_text = corr.str();
        TextOutW(dc, plot.left + 12, plot.top + 12, corr_text.c_str(), static_cast<int>(corr_text.size()));
    }

    SelectObject(dc, old_pen);
    SelectObject(dc, old_brush);
    DeleteObject(axis_pen);
}

void LiveApp::draw_room(HDC dc, const RECT& outer, const RECT& plot) {
    draw_spectrum_axes(dc, outer, plot);
    draw_note_guides(dc, plot, true);
    const auto& room_bins = freeze_display_ ? frozen_average_bins_ : average_bins_;
    draw_spectrum_trace(dc, plot, room_bins, RGB(255, 210, 110), 2, PS_SOLID);
    if (reference_available_ && !freeze_display_) {
        draw_spectrum_trace(dc, plot, frozen_average_bins_, RGB(140, 140, 140), 1, PS_DOT);
    }
    draw_peak_labels(dc, plot, collect_top_peaks(room_bins, 5));
}

double LiveApp::max_display_frequency() const {
    const int effective_rate =
        capture_sample_rate_ > 0
            ? (std::min)(analysis_sample_rate_, capture_sample_rate_)
            : (std::max)(analysis_sample_rate_, last_stream_sample_rate_);
    const double nyquist = effective_rate > 0 ? effective_rate / 2.0 : 22050.0;
    if (settings_.max_frequency > 0.0) {
        return std::min(settings_.max_frequency, nyquist);
    }
    return nyquist;
}

bool LiveApp::use_stereo_display() const {
    if (last_stream_channels_ < 2) {
        return false;
    }
    if (channel_display_mode_ == ChannelDisplayMode::Mono) {
        return false;
    }
    return true;
}

int LiveApp::spectrogram_pane_count() const {
    return use_stereo_display() ? 2 : 1;
}

RECT LiveApp::spectrogram_pane_rect(const RECT& plot, int index) const {
    if (!use_stereo_display()) {
        return plot;
    }

    const int gap = 12;
    const int pane_height = (std::max)(static_cast<int>((plot.bottom - plot.top - gap) / 2), 1);
    RECT pane = plot;
    if (index <= 0) {
        pane.bottom = pane.top + pane_height;
    } else {
        pane.top = plot.top + pane_height + gap;
    }
    return pane;
}

double LiveApp::frequency_for_position(double normalized) const {
    const double max_frequency = max_display_frequency();
    const double min_frequency = std::clamp(settings_.min_frequency, 0.0, max_frequency);

    if (settings_.frequency_scale == gram::FrequencyScale::Linear) {
        return min_frequency + normalized * (max_frequency - min_frequency);
    }
    if (settings_.frequency_scale == gram::FrequencyScale::Logarithmic) {
        const double min_log_freq = std::max(min_frequency, 1.0);
        const double span =
            std::log10(std::max(max_frequency, min_log_freq)) - std::log10(min_log_freq);
        return std::pow(10.0, std::log10(min_log_freq) + normalized * span);
    }

    const double min_octave_freq = std::max(min_frequency, 1.0);
    const double top = std::max(max_frequency, min_octave_freq);
    const double octave_span = std::log2(top / min_octave_freq);
    return min_octave_freq * std::pow(2.0, normalized * octave_span);
}

double LiveApp::spectrum_level_at_position(const std::vector<double>& db_bins, double normalized) const {
    if (db_bins.empty()) {
        return settings_.min_db;
    }

    const double nyquist = analysis_sample_rate_ > 0 ? analysis_sample_rate_ / 2.0 : 22050.0;
    const double target_frequency = frequency_for_position(normalized);
    const auto rounded = static_cast<long long>(
        std::lround((target_frequency / nyquist) * static_cast<double>(db_bins.size() - 1)));
    const auto clamped = std::clamp<long long>(rounded, 0, static_cast<long long>(db_bins.size() - 1));
    return db_bins[static_cast<std::size_t>(clamped)];
}

double LiveApp::normalized_amplitude(double db_value) const {
    const double clamped_db = std::clamp(db_value, settings_.min_db, settings_.max_db);
    if (spectrum_amplitude_mode_ == SpectrumAmplitudeMode::Decibels) {
        return (clamped_db - settings_.min_db) / (settings_.max_db - settings_.min_db);
    }

    const double min_linear = std::pow(10.0, settings_.min_db / 20.0);
    const double max_linear = std::pow(10.0, settings_.max_db / 20.0);
    const double value = std::pow(10.0, clamped_db / 20.0);
    return (value - min_linear) / (max_linear - min_linear);
}

std::wstring LiveApp::format_amplitude_label(double normalized) const {
    std::wostringstream stream;
    if (spectrum_amplitude_mode_ == SpectrumAmplitudeMode::Decibels) {
        const double value = settings_.min_db + normalized * (settings_.max_db - settings_.min_db);
        stream << std::fixed << std::setprecision(0) << value << L" dB";
    } else {
        const double min_linear = std::pow(10.0, settings_.min_db / 20.0);
        const double max_linear = std::pow(10.0, settings_.max_db / 20.0);
        const double value = min_linear + normalized * (max_linear - min_linear);
        stream << std::fixed << std::setprecision(2) << value;
    }
    return stream.str();
}

std::wstring LiveApp::format_frequency_label(double frequency) const {
    std::wostringstream stream;
    if (frequency >= 1000.0) {
        stream << std::fixed << std::setprecision(frequency >= 10000.0 ? 0 : 1)
               << (frequency / 1000.0) << L" kHz";
    } else {
        stream << std::fixed << std::setprecision(0) << frequency << L" Hz";
    }
    return stream.str();
}

void LiveApp::draw_spectrogram_axes(HDC dc, const RECT& outer, const RECT& plot, const wchar_t* title) {
    HPEN grid_pen = CreatePen(PS_SOLID, 1, RGB(48, 48, 48));
    HPEN axis_pen = CreatePen(PS_SOLID, 1, RGB(96, 96, 96));
    HGDIOBJ old_pen = SelectObject(dc, grid_pen);
    HGDIOBJ old_brush = SelectObject(dc, GetStockObject(NULL_BRUSH));
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(180, 180, 180));

    if (show_grid_) {
        for (int index = 1; index < 10; ++index) {
            const int x = plot.left + ((plot.right - plot.left) * index) / 10;
            MoveToEx(dc, x, plot.top, nullptr);
            LineTo(dc, x, plot.bottom);
        }
        for (int index = 1; index < 6; ++index) {
            const int y = plot.top + ((plot.bottom - plot.top) * index) / 6;
            MoveToEx(dc, plot.left, y, nullptr);
            LineTo(dc, plot.right, y);
        }
    }

    SelectObject(dc, axis_pen);
    Rectangle(dc, plot.left, plot.top, plot.right, plot.bottom);
    draw_note_guides(dc, plot, false);

    for (int index = 0; index <= 6; ++index) {
        const double normalized = static_cast<double>(index) / 6.0;
        const int y = plot.bottom - static_cast<int>((plot.bottom - plot.top) * normalized);
        const std::wstring text = format_frequency_label(frequency_for_position(normalized));
        TextOutW(dc, outer.left + 4, y - 8, text.c_str(), static_cast<int>(text.size()));
        TextOutW(dc, plot.right + 8, y - 8, text.c_str(), static_cast<int>(text.size()));
    }

    if (title != nullptr && *title != L'\0') {
        SetTextColor(dc, RGB(220, 220, 220));
        TextOutW(dc, plot.left + 8, plot.top + 6, title, static_cast<int>(wcslen(title)));
    }

    SelectObject(dc, old_pen);
    SelectObject(dc, old_brush);
    DeleteObject(grid_pen);
    DeleteObject(axis_pen);
}

void LiveApp::draw_spectrum_axes(HDC dc, const RECT& outer, const RECT& plot) {
    HPEN grid_pen = CreatePen(PS_SOLID, 1, RGB(48, 48, 48));
    HPEN axis_pen = CreatePen(PS_SOLID, 1, RGB(96, 96, 96));
    HGDIOBJ old_pen = SelectObject(dc, grid_pen);
    HGDIOBJ old_brush = SelectObject(dc, GetStockObject(NULL_BRUSH));
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(180, 180, 180));

    if (show_grid_) {
        for (int index = 1; index < 6; ++index) {
            const int x = plot.left + ((plot.right - plot.left) * index) / 6;
            MoveToEx(dc, x, plot.top, nullptr);
            LineTo(dc, x, plot.bottom);
        }
        for (int index = 1; index < 6; ++index) {
            const int y = plot.top + ((plot.bottom - plot.top) * index) / 6;
            MoveToEx(dc, plot.left, y, nullptr);
            LineTo(dc, plot.right, y);
        }
    }

    SelectObject(dc, axis_pen);
    Rectangle(dc, plot.left, plot.top, plot.right, plot.bottom);
    draw_note_guides(dc, plot, true);

    for (int index = 0; index <= 6; ++index) {
        const double normalized = 1.0 - static_cast<double>(index) / 6.0;
        const int y = plot.top + ((plot.bottom - plot.top) * index) / 6;
        const std::wstring text = format_amplitude_label(normalized);
        TextOutW(dc, outer.left + 4, y - 8, text.c_str(), static_cast<int>(text.size()));
        TextOutW(dc, plot.right + 8, y - 8, text.c_str(), static_cast<int>(text.size()));
    }

    for (int index = 0; index <= 6; ++index) {
        const double normalized = static_cast<double>(index) / 6.0;
        const int x = plot.left + ((plot.right - plot.left) * index) / 6;
        const std::wstring text = format_frequency_label(frequency_for_position(normalized));
        TextOutW(dc, x - 18, plot.bottom + 8, text.c_str(), static_cast<int>(text.size()));
    }

    SelectObject(dc, old_pen);
    SelectObject(dc, old_brush);
    DeleteObject(grid_pen);
    DeleteObject(axis_pen);
}

void LiveApp::draw_spectrum_trace(HDC dc, const RECT& plot, const std::vector<double>& db_bins, COLORREF color, int line_width, int style) {
    if (db_bins.empty()) {
        return;
    }

    HPEN line_pen = CreatePen(style, line_width, color);
    HGDIOBJ old_pen = SelectObject(dc, line_pen);

    const int width = (std::max)(static_cast<int>(plot.right - plot.left), 1);
    const int height = (std::max)(static_cast<int>(plot.bottom - plot.top), 1);
    for (int x = 0; x < width; ++x) {
        const double normalized = width > 1
            ? static_cast<double>(x) / static_cast<double>(width - 1)
            : 0.0;
        const double db_value = spectrum_level_at_position(db_bins, normalized);
        const double amplitude = normalized_amplitude(db_value);
        const int y = plot.bottom - 1 -
            static_cast<int>(std::lround(amplitude * static_cast<double>(height - 1)));
        const int px = plot.left + x;
        if (x == 0) {
            MoveToEx(dc, px, y, nullptr);
        } else {
            LineTo(dc, px, y);
        }
    }

    SelectObject(dc, old_pen);
    DeleteObject(line_pen);
}

void LiveApp::draw_spectrum(HDC dc, const RECT& outer, const RECT& plot) {
    draw_spectrum_axes(dc, outer, plot);

    const auto& mono_bins = freeze_display_ ? frozen_db_bins_ : db_bins_;
    const auto& left_bins = freeze_display_ ? frozen_db_bins_left_ : db_bins_left_;
    const auto& right_bins = freeze_display_ ? frozen_db_bins_right_ : db_bins_right_;
    const auto& mono_peak = freeze_display_ ? frozen_peak_hold_bins_ : peak_hold_bins_;
    const auto& left_peak = freeze_display_ ? frozen_peak_hold_bins_left_ : peak_hold_bins_left_;
    const auto& right_peak = freeze_display_ ? frozen_peak_hold_bins_right_ : peak_hold_bins_right_;
    const auto& mono_average = freeze_display_ ? frozen_average_bins_ : average_bins_;
    const auto& left_average = freeze_display_ ? frozen_average_bins_left_ : average_bins_left_;
    const auto& right_average = freeze_display_ ? frozen_average_bins_right_ : average_bins_right_;
    const auto& reference_bins = selected_bins(true);
    if (reference_available_ && !freeze_display_ && !reference_bins.empty()) {
        draw_spectrum_trace(dc, plot, reference_bins, RGB(130, 130, 130), 1, PS_DOT);
    }

    if (use_stereo_display()) {
        if (show_average_) {
            draw_spectrum_trace(dc, plot, left_average, RGB(110, 255, 255), 1, PS_DOT);
            draw_spectrum_trace(dc, plot, right_average, RGB(255, 210, 120), 1, PS_DOT);
        }
        if (show_peak_hold_) {
            draw_spectrum_trace(dc, plot, left_peak, RGB(180, 245, 255), 1, PS_SOLID);
            draw_spectrum_trace(dc, plot, right_peak, RGB(255, 225, 180), 1, PS_SOLID);
        }
        draw_spectrum_trace(dc, plot, left_bins, RGB(80, 220, 255), 2, PS_SOLID);
        draw_spectrum_trace(dc, plot, right_bins, RGB(255, 170, 70), 2, PS_SOLID);
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, RGB(80, 220, 255));
        const wchar_t* left_text = L"L";
        TextOutW(dc, plot.left + 8, plot.top + 6, left_text, 1);
        SetTextColor(dc, RGB(255, 170, 70));
        const wchar_t* right_text = L"R";
        TextOutW(dc, plot.left + 28, plot.top + 6, right_text, 1);
    } else {
        const std::vector<double>* bins = &mono_bins;
        const std::vector<double>* peak_bins = &mono_peak;
        const std::vector<double>* average_bins = &mono_average;
        if (mono_channel_mode_ == MonoChannelMode::Left) {
            bins = &left_bins;
            peak_bins = &left_peak;
            average_bins = &left_average;
        } else if (mono_channel_mode_ == MonoChannelMode::Right) {
            bins = &right_bins;
            peak_bins = &right_peak;
            average_bins = &right_average;
        }
        if (show_average_) {
            draw_spectrum_trace(dc, plot, *average_bins, RGB(120, 255, 140), 1, PS_DOT);
        }
        if (show_peak_hold_) {
            draw_spectrum_trace(dc, plot, *peak_bins, RGB(255, 255, 180), 1, PS_SOLID);
        }
        draw_spectrum_trace(dc, plot, *bins, RGB(255, 180, 60), 2, PS_SOLID);
        const auto peaks = collect_top_peaks(*bins, 5);
        draw_peak_labels(dc, plot, peaks);
        if (!peaks.empty()) {
            HPEN harmonic_pen = CreatePen(PS_DOT, 1, RGB(160, 160, 220));
            HGDIOBJ old_pen = SelectObject(dc, harmonic_pen);
            for (int harmonic = 2; harmonic <= 6; ++harmonic) {
                const double frequency = peaks.front().frequency * harmonic;
                if (frequency > max_display_frequency()) {
                    break;
                }
                const double normalized = normalized_for_frequency(frequency);
                const int x = plot.left + static_cast<int>(std::lround(normalized * static_cast<double>(plot.right - plot.left - 1)));
                MoveToEx(dc, x, plot.top, nullptr);
                LineTo(dc, x, plot.bottom);
            }
            SelectObject(dc, old_pen);
            DeleteObject(harmonic_pen);
        }
    }
}

void LiveApp::draw_spectrogram_image(HDC dc, const RECT& plot, const gram::Image& image) {
    BITMAPINFO bitmap_info {};
    bitmap_info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmap_info.bmiHeader.biWidth = image.width;
    bitmap_info.bmiHeader.biHeight = -image.height;
    bitmap_info.bmiHeader.biPlanes = 1;
    bitmap_info.bmiHeader.biBitCount = 32;
    bitmap_info.bmiHeader.biCompression = BI_RGB;

    StretchDIBits(
        dc,
        plot.left,
        plot.top,
        plot.right - plot.left,
        plot.bottom - plot.top,
        0,
        0,
        image.width,
        image.height,
        image.pixels.data(),
        &bitmap_info,
        DIB_RGB_COLORS,
        SRCCOPY);
}

void LiveApp::paint(HDC dc) {
    RECT client {};
    GetClientRect(hwnd_, &client);
    const int client_width = client.right - client.left;
    const int client_height = client.bottom - client.top;
    if (client_width <= 0 || client_height <= 0) {
        return;
    }

    HDC memory_dc = CreateCompatibleDC(dc);
    HBITMAP bitmap = CreateCompatibleBitmap(dc, client_width, client_height);
    if (memory_dc == nullptr || bitmap == nullptr) {
        if (bitmap != nullptr) {
            DeleteObject(bitmap);
        }
        if (memory_dc != nullptr) {
            DeleteDC(memory_dc);
        }
        return;
    }

    HGDIOBJ old_bitmap = SelectObject(memory_dc, bitmap);
    FillRect(memory_dc, &client, reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1));

    const RECT rect = graph_rect();
    const RECT plot = plot_rect();
    const int graph_width = rect.right - rect.left;
    const int graph_height = rect.bottom - rect.top;
    if (graph_width > 0 && graph_height > 0) {
        FillRect(memory_dc, &rect, reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
        if (display_mode_ == DisplayMode::Spectrogram) {
            if (use_stereo_display()) {
                const RECT left_plot = spectrogram_pane_rect(plot, 0);
                const RECT right_plot = spectrogram_pane_rect(plot, 1);
                draw_spectrogram_image(memory_dc, left_plot, freeze_display_ ? frozen_left_image_ : left_image_);
                draw_spectrogram_axes(memory_dc, rect, left_plot, L"Left");
                draw_spectrogram_markers(memory_dc, left_plot);
                draw_spectrogram_image(memory_dc, right_plot, freeze_display_ ? frozen_right_image_ : right_image_);
                draw_spectrogram_axes(memory_dc, rect, right_plot, L"Right");
                draw_spectrogram_markers(memory_dc, right_plot);
            } else {
                const gram::Image* image = freeze_display_ ? &frozen_image_ : &image_;
                if (mono_channel_mode_ == MonoChannelMode::Left) {
                    image = freeze_display_ ? &frozen_left_image_ : &left_image_;
                } else if (mono_channel_mode_ == MonoChannelMode::Right) {
                    image = freeze_display_ ? &frozen_right_image_ : &right_image_;
                }
                draw_spectrogram_image(memory_dc, plot, *image);
                draw_spectrogram_axes(memory_dc, rect, plot, nullptr);
                draw_spectrogram_markers(memory_dc, plot);
            }
        } else if (display_mode_ == DisplayMode::Spectrum) {
            draw_spectrum(memory_dc, rect, plot);
        } else if (display_mode_ == DisplayMode::Tuner) {
            draw_tuner(memory_dc, rect, plot);
        } else if (display_mode_ == DisplayMode::Vectorscope) {
            draw_vectorscope(memory_dc, rect, plot);
        } else {
            draw_room(memory_dc, rect, plot);
        }
    }

    BitBlt(dc, 0, 0, client_width, client_height, memory_dc, 0, 0, SRCCOPY);
    SelectObject(memory_dc, old_bitmap);
    DeleteObject(bitmap);
    DeleteDC(memory_dc);
}

} // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int show) {
    try {
        LiveApp app;
        if (!app.create(instance, show)) {
            return 1;
        }

        MSG message {};
        while (GetMessageW(&message, nullptr, 0, 0) > 0) {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
        return static_cast<int>(message.wParam);
    } catch (const std::exception& error) {
        const auto message = widen(error.what());
        show_logged_error(nullptr, L"wWinMain.exception", message);
        return 1;
    } catch (...) {
        show_logged_error(nullptr, L"wWinMain.unknown", L"An unexpected error occurred.");
        return 1;
    }
}
