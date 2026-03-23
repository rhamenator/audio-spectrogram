#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#include <windows.h>
#include <commctrl.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <functiondiscoverykeys_devpkey.h>

#include <algorithm>
#include <iomanip>
#include <sstream>
#include <string>
#include <vector>

namespace {

constexpr int kEditId = 2001;
constexpr int kOkId = 2002;

template <typename T>
void safe_release(T*& value) {
    if (value != nullptr) {
        value->Release();
        value = nullptr;
    }
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

std::wstring state_text(DWORD state) {
    std::vector<std::wstring> names;
    if ((state & DEVICE_STATE_ACTIVE) != 0) {
        names.push_back(L"ACTIVE");
    }
    if ((state & DEVICE_STATE_DISABLED) != 0) {
        names.push_back(L"DISABLED");
    }
    if ((state & DEVICE_STATE_NOTPRESENT) != 0) {
        names.push_back(L"NOTPRESENT");
    }
    if ((state & DEVICE_STATE_UNPLUGGED) != 0) {
        names.push_back(L"UNPLUGGED");
    }
    if (names.empty()) {
        return L"UNKNOWN";
    }
    std::wstring text;
    for (std::size_t index = 0; index < names.size(); ++index) {
        if (index > 0) {
            text += L"|";
        }
        text += names[index];
    }
    return text;
}

std::wstring read_friendly_name(IMMDevice* device) {
    IPropertyStore* store = nullptr;
    PROPVARIANT value {};
    PropVariantInit(&value);
    std::wstring text = L"(unknown)";
    if (SUCCEEDED(device->OpenPropertyStore(STGM_READ, &store))) {
        if (SUCCEEDED(store->GetValue(PKEY_Device_FriendlyName, &value)) && value.vt == VT_LPWSTR && value.pwszVal != nullptr) {
            text = value.pwszVal;
        }
    }
    PropVariantClear(&value);
    safe_release(store);
    return text;
}

std::wstring wave_format_text(const WAVEFORMATEX* format) {
    if (format == nullptr) {
        return L"(none)";
    }
    std::wostringstream stream;
    stream << format->nSamplesPerSec << L" Hz, "
           << format->nChannels << L" ch, "
           << format->wBitsPerSample << L" bit, tag 0x"
           << std::hex << std::uppercase << format->wFormatTag;
    return stream.str();
}

std::wstring sample_rate_support_report(IAudioClient* client, const WAVEFORMATEX* base_format) {
    if (client == nullptr || base_format == nullptr) {
        return L"  Shared-mode rates: unavailable\n";
    }

    const int rates[] = { 8000, 11025, 16000, 22050, 24000, 32000, 44100, 48000, 88200, 96000, 192000 };
    std::wostringstream stream;
    stream << L"  Shared-mode exact support:";
    for (int rate : rates) {
        WAVEFORMATEX probe = *base_format;
        probe.nSamplesPerSec = static_cast<DWORD>(rate);
        probe.nAvgBytesPerSec = probe.nSamplesPerSec * probe.nBlockAlign;
        WAVEFORMATEX* closest = nullptr;
        const HRESULT hr = client->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED, &probe, &closest);
        if (hr == S_OK) {
            stream << L"\n    " << rate << L" Hz: exact";
        } else if (hr == S_FALSE) {
            stream << L"\n    " << rate << L" Hz: closest available";
        } else {
            stream << L"\n    " << rate << L" Hz: no (" << format_hresult(hr) << L")";
        }
        if (closest != nullptr) {
            CoTaskMemFree(closest);
        }
    }
    stream << L"\n";
    return stream.str();
}

void append_endpoint_report(std::wostringstream& output, IMMDevice* device, const wchar_t* role_label, bool is_default) {
    wchar_t* id = nullptr;
    DWORD state = 0;
    IAudioClient* client = nullptr;
    WAVEFORMATEX* mix = nullptr;

    output << role_label << (is_default ? L" (default)" : L"") << L"\n";
    output << L"  Name: " << read_friendly_name(device) << L"\n";

    if (SUCCEEDED(device->GetId(&id)) && id != nullptr) {
        output << L"  Id: " << id << L"\n";
    }
    if (SUCCEEDED(device->GetState(&state))) {
        output << L"  State: " << state_text(state) << L"\n";
    }

    const HRESULT activate = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&client));
    if (FAILED(activate)) {
        output << L"  Activate: " << format_hresult(activate) << L"\n\n";
        if (id != nullptr) {
            CoTaskMemFree(id);
        }
        safe_release(client);
        return;
    }

    const HRESULT mix_result = client->GetMixFormat(&mix);
    if (SUCCEEDED(mix_result)) {
        output << L"  Mix format: " << wave_format_text(mix) << L"\n";
    } else {
        output << L"  Mix format: " << format_hresult(mix_result) << L"\n";
    }

    REFERENCE_TIME default_period = 0;
    REFERENCE_TIME minimum_period = 0;
    const HRESULT period_result = client->GetDevicePeriod(&default_period, &minimum_period);
    if (SUCCEEDED(period_result)) {
        output << L"  Device period: default " << (default_period / 10000.0) << L" ms, minimum "
               << (minimum_period / 10000.0) << L" ms\n";
    } else {
        output << L"  Device period: " << format_hresult(period_result) << L"\n";
    }

    output << sample_rate_support_report(client, mix);
    output << L"\n";

    if (mix != nullptr) {
        CoTaskMemFree(mix);
    }
    if (id != nullptr) {
        CoTaskMemFree(id);
    }
    safe_release(client);
}

std::wstring build_report() {
    std::wostringstream output;
    output << L"gram_audio_info\n";
    output << L"Windows audio endpoint report\n\n";

    const HRESULT init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(init) && init != RPC_E_CHANGED_MODE) {
        output << L"COM init failed: " << format_hresult(init) << L"\n";
        return output.str();
    }
    const bool com_initialized = SUCCEEDED(init);

    IMMDeviceEnumerator* enumerator = nullptr;
    const HRESULT create = CoCreateInstance(
        __uuidof(MMDeviceEnumerator),
        nullptr,
        CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        reinterpret_cast<void**>(&enumerator));
    if (FAILED(create)) {
        output << L"Unable to create MMDeviceEnumerator: " << format_hresult(create) << L"\n";
        if (com_initialized) {
            CoUninitialize();
        }
        return output.str();
    }

    const ERole roles[] = { eConsole, eMultimedia, eCommunications };
    const wchar_t* role_names[] = { L"Default console render", L"Default multimedia render", L"Default communications render" };
    for (int index = 0; index < 3; ++index) {
        IMMDevice* device = nullptr;
        const HRESULT hr = enumerator->GetDefaultAudioEndpoint(eRender, roles[index], &device);
        if (SUCCEEDED(hr) && device != nullptr) {
            append_endpoint_report(output, device, role_names[index], true);
        } else {
            output << role_names[index] << L"\n  Error: " << format_hresult(hr) << L"\n\n";
        }
        safe_release(device);
    }

    IMMDeviceCollection* collection = nullptr;
    const HRESULT list = enumerator->EnumAudioEndpoints(
        eRender,
        DEVICE_STATE_ACTIVE | DEVICE_STATE_DISABLED | DEVICE_STATE_NOTPRESENT | DEVICE_STATE_UNPLUGGED,
        &collection);
    if (FAILED(list)) {
        output << L"Unable to enumerate render devices: " << format_hresult(list) << L"\n";
    } else {
        UINT count = 0;
        collection->GetCount(&count);
        output << L"Render endpoints (" << count << L")\n\n";
        for (UINT index = 0; index < count; ++index) {
            IMMDevice* device = nullptr;
            if (SUCCEEDED(collection->Item(index, &device)) && device != nullptr) {
                std::wostringstream label;
                label << L"Render endpoint #" << (index + 1);
                append_endpoint_report(output, device, label.str().c_str(), false);
            }
            safe_release(device);
        }
    }

    safe_release(collection);
    safe_release(enumerator);
    if (com_initialized) {
        CoUninitialize();
    }
    return output.str();
}

class AudioInfoWindow {
public:
    bool create(HINSTANCE instance, int show) {
        INITCOMMONCONTROLSEX controls {};
        controls.dwSize = sizeof(controls);
        controls.dwICC = ICC_WIN95_CLASSES;
        InitCommonControlsEx(&controls);

        WNDCLASSW window_class {};
        window_class.lpfnWndProc = &AudioInfoWindow::window_proc;
        window_class.hInstance = instance;
        window_class.lpszClassName = L"GramAudioInfoWindow";
        window_class.hCursor = LoadCursorW(nullptr, MAKEINTRESOURCEW(32512));
        window_class.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        RegisterClassW(&window_class);

        hwnd_ = CreateWindowExW(
            0,
            window_class.lpszClassName,
            L"gram_audio_info",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            920,
            700,
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

private:
    static LRESULT CALLBACK window_proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
        AudioInfoWindow* self = nullptr;
        if (message == WM_NCCREATE) {
            auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
            self = static_cast<AudioInfoWindow*>(create->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            self->hwnd_ = hwnd;
        } else {
            self = reinterpret_cast<AudioInfoWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }
        if (self != nullptr) {
            return self->handle_message(message, wparam, lparam);
        }
        return DefWindowProcW(hwnd, message, wparam, lparam);
    }

    LRESULT handle_message(UINT message, WPARAM wparam, LPARAM lparam) {
        switch (message) {
        case WM_CREATE:
            edit_ = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                L"EDIT",
                nullptr,
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                0,
                0,
                0,
                0,
                hwnd_,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(kEditId)),
                nullptr,
                nullptr);
            ok_button_ = CreateWindowExW(
                0,
                L"BUTTON",
                L"OK",
                WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                0,
                0,
                0,
                0,
                hwnd_,
                reinterpret_cast<HMENU>(static_cast<INT_PTR>(kOkId)),
                nullptr,
                nullptr);
            SetWindowTextW(edit_, build_report().c_str());
            return 0;

        case WM_SIZE: {
            RECT client {};
            GetClientRect(hwnd_, &client);
            MoveWindow(edit_, 12, 12, client.right - 24, client.bottom - 68, TRUE);
            MoveWindow(ok_button_, client.right - 104, client.bottom - 40, 88, 28, TRUE);
            return 0;
        }

        case WM_COMMAND:
            if (LOWORD(wparam) == kOkId) {
                DestroyWindow(hwnd_);
                return 0;
            }
            break;

        case WM_CLOSE:
            DestroyWindow(hwnd_);
            return 0;

        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;

        default:
            break;
        }
        return DefWindowProcW(hwnd_, message, wparam, lparam);
    }

    HWND hwnd_ = nullptr;
    HWND edit_ = nullptr;
    HWND ok_button_ = nullptr;
};

} // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int show) {
    AudioInfoWindow app;
    if (!app.create(instance, show)) {
        MessageBoxW(nullptr, L"Unable to create the audio info window.", L"gram_audio_info", MB_ICONERROR);
        return 1;
    }

    MSG message {};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    return static_cast<int>(message.wParam);
}
