[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$typeName = 'AudioSpectrogram.PowerShell.AudioEndpointProbe'
if (-not ($typeName -as [type])) {
Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace AudioSpectrogram.PowerShell {
    internal enum EDataFlow {
        eRender = 0,
        eCapture = 1,
        eAll = 2
    }

    internal enum ERole {
        eConsole = 0,
        eMultimedia = 1,
        eCommunications = 2
    }

    [Flags]
    internal enum DeviceState : uint {
        ACTIVE = 0x00000001,
        DISABLED = 0x00000002,
        NOTPRESENT = 0x00000004,
        UNPLUGGED = 0x00000008
    }

    internal enum AUDCLNT_SHAREMODE {
        SHARED = 0,
        EXCLUSIVE = 1
    }

    [Flags]
    internal enum CLSCTX : uint {
        INPROC_SERVER = 0x1,
        INPROC_HANDLER = 0x2,
        LOCAL_SERVER = 0x4,
        REMOTE_SERVER = 0x10,
        ALL = INPROC_SERVER | INPROC_HANDLER | LOCAL_SERVER | REMOTE_SERVER
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PROPERTYKEY {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct PROPVARIANT {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pointerValue;
    }

    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    internal struct WAVEFORMATEX {
        public ushort wFormatTag;
        public ushort nChannels;
        public uint nSamplesPerSec;
        public uint nAvgBytesPerSec;
        public ushort nBlockAlign;
        public ushort wBitsPerSample;
        public ushort cbSize;
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumeratorComObject {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator {
        [PreserveSig]
        int EnumAudioEndpoints(EDataFlow dataFlow, DeviceState stateMask, out IMMDeviceCollection devices);
        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice endpoint);
        [PreserveSig]
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string id, out IMMDevice device);
        [PreserveSig]
        int RegisterEndpointNotificationCallback(IntPtr client);
        [PreserveSig]
        int UnregisterEndpointNotificationCallback(IntPtr client);
    }

    [ComImport]
    [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection {
        [PreserveSig]
        int GetCount(out uint count);
        [PreserveSig]
        int Item(uint index, out IMMDevice device);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice {
        [PreserveSig]
        int Activate(ref Guid iid, CLSCTX clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object interfacePointer);
        [PreserveSig]
        int OpenPropertyStore(int stgmAccess, out IPropertyStore properties);
        [PreserveSig]
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
        [PreserveSig]
        int GetState(out uint state);
    }

    [ComImport]
    [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore {
        [PreserveSig]
        int GetCount(out uint propertyCount);
        [PreserveSig]
        int GetAt(uint propertyIndex, out PROPERTYKEY key);
        [PreserveSig]
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT value);
        [PreserveSig]
        int SetValue(ref PROPERTYKEY key, ref PROPVARIANT value);
        [PreserveSig]
        int Commit();
    }

    [ComImport]
    [Guid("1CB9AD4C-DBFA-4c32-B178-C2F568A703B2")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioClient {
        [PreserveSig]
        int Initialize(AUDCLNT_SHAREMODE shareMode, uint streamFlags, long hnsBufferDuration, long hnsPeriodicity, IntPtr format, IntPtr audioSessionGuid);
        [PreserveSig]
        int GetBufferSize(out uint bufferSize);
        [PreserveSig]
        int GetStreamLatency(out long latency);
        [PreserveSig]
        int GetCurrentPadding(out uint currentPadding);
        [PreserveSig]
        int IsFormatSupported(AUDCLNT_SHAREMODE shareMode, IntPtr format, out IntPtr closestMatch);
        [PreserveSig]
        int GetMixFormat(out IntPtr deviceFormat);
        [PreserveSig]
        int GetDevicePeriod(out long defaultDevicePeriod, out long minimumDevicePeriod);
        [PreserveSig]
        int Start();
        [PreserveSig]
        int Stop();
        [PreserveSig]
        int Reset();
        [PreserveSig]
        int SetEventHandle(IntPtr eventHandle);
        [PreserveSig]
        int GetService(ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object service);
    }

    internal static class NativeMethods {
        [DllImport("ole32.dll")]
        internal static extern int PropVariantClear(ref PROPVARIANT value);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        internal static extern int FormatMessageW(
            int flags,
            IntPtr source,
            int messageId,
            int languageId,
            StringBuilder buffer,
            int size,
            IntPtr arguments);
    }

    public static class AudioEndpointProbe {
        private static readonly PROPERTYKEY PKEY_Device_FriendlyName = new PROPERTYKEY {
            fmtid = new Guid("a45c254e-df1c-4efd-8020-67d146a850e0"),
            pid = 14
        };

        private static string HResultText(int hr) {
            return string.Format("0x{0:X8}", hr);
        }

        private static string DeviceStateText(uint state) {
            var names = new List<string>();
            if ((state & (uint)DeviceState.ACTIVE) != 0) names.Add("ACTIVE");
            if ((state & (uint)DeviceState.DISABLED) != 0) names.Add("DISABLED");
            if ((state & (uint)DeviceState.NOTPRESENT) != 0) names.Add("NOTPRESENT");
            if ((state & (uint)DeviceState.UNPLUGGED) != 0) names.Add("UNPLUGGED");
            return names.Count == 0 ? "UNKNOWN" : string.Join("|", names);
        }

        private static string EndpointKind(string name) {
            var lower = (name ?? string.Empty).ToLowerInvariant();
            if (lower.Contains("fxsound") || lower.Contains("enhancer") || lower.Contains("virtual") ||
                lower.Contains("voicemeeter") || lower.Contains("vb-audio") || lower.Contains("sonar") ||
                lower.Contains("apo")) {
                return "virtual/processed";
            }
            if (lower.Contains("realtek") || lower.Contains("usb") || lower.Contains("display audio") ||
                lower.Contains("headphones") || lower.Contains("speakers") || lower.Contains("hdmi")) {
                return "hardware-backed";
            }
            return "unknown";
        }

        private static string FriendlyName(IMMDevice device) {
            IPropertyStore store = null;
            PROPVARIANT value = new PROPVARIANT();
            PROPERTYKEY key = PKEY_Device_FriendlyName;
            try {
                Marshal.ThrowExceptionForHR(device.OpenPropertyStore(0, out store));
                Marshal.ThrowExceptionForHR(store.GetValue(ref key, out value));
                if (value.vt == 31 && value.pointerValue != IntPtr.Zero) {
                    return Marshal.PtrToStringUni(value.pointerValue) ?? string.Empty;
                }
                return string.Empty;
            } finally {
                NativeMethods.PropVariantClear(ref value);
                if (store != null) Marshal.ReleaseComObject(store);
            }
        }

        private static string WaveFormatText(IntPtr formatPtr) {
            if (formatPtr == IntPtr.Zero) return "(none)";
            var format = Marshal.PtrToStructure<WAVEFORMATEX>(formatPtr);
            return string.Format("{0} Hz, {1} ch, {2} bit, tag 0x{3:X4}", format.nSamplesPerSec, format.nChannels, format.wBitsPerSample, format.wFormatTag);
        }

        private static byte[] CloneFormatBlob(IntPtr formatPtr) {
            var format = Marshal.PtrToStructure<WAVEFORMATEX>(formatPtr);
            int size = Marshal.SizeOf<WAVEFORMATEX>() + format.cbSize;
            var blob = new byte[size];
            Marshal.Copy(formatPtr, blob, 0, size);
            return blob;
        }

        private static string InitializeProbe(IMMDevice device, byte[] probeBlob) {
            object raw = null;
            IAudioClient client = null;
            IntPtr formatPtr = IntPtr.Zero;
            try {
                var iid = typeof(IAudioClient).GUID;
                Marshal.ThrowExceptionForHR(device.Activate(ref iid, CLSCTX.ALL, IntPtr.Zero, out raw));
                client = (IAudioClient)raw;
                formatPtr = Marshal.AllocCoTaskMem(probeBlob.Length);
                Marshal.Copy(probeBlob, 0, formatPtr, probeBlob.Length);
                int hr = client.Initialize(AUDCLNT_SHAREMODE.SHARED, 0, 10000000, 0, formatPtr, IntPtr.Zero);
                return hr == 0 ? "init ok" : "init " + HResultText(hr);
            } catch (Exception ex) {
                return "init " + ex.GetType().Name;
            } finally {
                if (formatPtr != IntPtr.Zero) Marshal.FreeCoTaskMem(formatPtr);
                if (client != null) Marshal.ReleaseComObject(client);
            }
        }

        private static string SharedSupport(IMMDevice device, IAudioClient client, IntPtr mixPtr) {
            var rates = new[] { 8000, 11025, 16000, 22050, 24000, 32000, 44100, 48000, 88200, 96000, 192000 };
            var blob = CloneFormatBlob(mixPtr);
            int sampleRateOffset = (int)Marshal.OffsetOf(typeof(WAVEFORMATEX), "nSamplesPerSec");
            int avgBytesOffset = (int)Marshal.OffsetOf(typeof(WAVEFORMATEX), "nAvgBytesPerSec");
            int blockAlignOffset = (int)Marshal.OffsetOf(typeof(WAVEFORMATEX), "nBlockAlign");
            ushort blockAlign = BitConverter.ToUInt16(blob, blockAlignOffset);
            var sb = new StringBuilder();
            sb.AppendLine("  Shared-mode support:");
            foreach (int rate in rates) {
                Array.Copy(BitConverter.GetBytes((uint)rate), 0, blob, sampleRateOffset, 4);
                Array.Copy(BitConverter.GetBytes((uint)(rate * blockAlign)), 0, blob, avgBytesOffset, 4);
                IntPtr probePtr = Marshal.AllocCoTaskMem(blob.Length);
                IntPtr closest = IntPtr.Zero;
                try {
                    Marshal.Copy(blob, 0, probePtr, blob.Length);
                    int hr = client.IsFormatSupported(AUDCLNT_SHAREMODE.SHARED, probePtr, out closest);
                    if (hr == 0) {
                        sb.Append("    ").Append(rate).Append(" Hz: exact, ").Append(InitializeProbe(device, blob)).AppendLine();
                    } else if (hr == 1) {
                        sb.Append("    ").Append(rate).Append(" Hz: closest available");
                        if (closest != IntPtr.Zero) {
                            sb.Append(" -> ").Append(WaveFormatText(closest));
                        }
                        sb.AppendLine();
                    } else {
                        sb.Append("    ").Append(rate).Append(" Hz: no (").Append(HResultText(hr)).AppendLine(")");
                    }
                } finally {
                    if (closest != IntPtr.Zero) Marshal.FreeCoTaskMem(closest);
                    Marshal.FreeCoTaskMem(probePtr);
                }
            }
            return sb.ToString();
        }

        private static string EndpointReport(IMMDevice device, string label, bool isDefault) {
            var sb = new StringBuilder();
            try {
                string name = FriendlyName(device);
                string id = string.Empty;
                uint state = 0;
                Marshal.ThrowExceptionForHR(device.GetId(out id));
                Marshal.ThrowExceptionForHR(device.GetState(out state));

                sb.Append(label);
                if (isDefault) sb.Append(" (default)");
                sb.AppendLine();
                sb.Append("  Name: ").AppendLine(name);
                sb.Append("  Path kind: ").AppendLine(EndpointKind(name));
                sb.Append("  Id: ").AppendLine(id);
                sb.Append("  State: ").AppendLine(DeviceStateText(state));

                object raw = null;
                IAudioClient client = null;
                IntPtr mixPtr = IntPtr.Zero;
                var iid = typeof(IAudioClient).GUID;
                int activate = device.Activate(ref iid, CLSCTX.ALL, IntPtr.Zero, out raw);
                if (activate != 0) {
                    sb.Append("  Activate: ").AppendLine(HResultText(activate));
                    sb.AppendLine();
                    return sb.ToString();
                }
                client = (IAudioClient)raw;
                int mixHr = client.GetMixFormat(out mixPtr);
                if (mixHr == 0) {
                    sb.Append("  Mix format: ").AppendLine(WaveFormatText(mixPtr));
                } else {
                    sb.Append("  Mix format: ").AppendLine(HResultText(mixHr));
                }
                long defaultPeriod;
                long minimumPeriod;
                int periodHr = client.GetDevicePeriod(out defaultPeriod, out minimumPeriod);
                if (periodHr == 0) {
                    sb.Append("  Device period: default ").Append(defaultPeriod / 10000.0).Append(" ms, minimum ")
                      .Append(minimumPeriod / 10000.0).AppendLine(" ms");
                } else {
                    sb.Append("  Device period: ").AppendLine(HResultText(periodHr));
                }
                if (mixPtr != IntPtr.Zero) {
                    sb.Append(SharedSupport(device, client, mixPtr));
                }
                if (mixPtr != IntPtr.Zero) Marshal.FreeCoTaskMem(mixPtr);
                if (client != null) Marshal.ReleaseComObject(client);
            } catch (Exception ex) {
                sb.Append(label);
                if (isDefault) sb.Append(" (default)");
                sb.AppendLine();
                sb.Append("  Error: ").AppendLine(ex.Message);
                sb.AppendLine();
            }
            sb.AppendLine();
            return sb.ToString();
        }

        public static string GetReport() {
            var sb = new StringBuilder();
            sb.AppendLine("audio-spectrogram PowerShell audio endpoint report");
            sb.AppendLine();

            IMMDeviceEnumerator enumerator = null;
            IMMDeviceCollection collection = null;
            try {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();

                ERole[] roles = { ERole.eConsole, ERole.eMultimedia, ERole.eCommunications };
                string[] labels = { "Default console render", "Default multimedia render", "Default communications render" };
                for (int i = 0; i < roles.Length; ++i) {
                    IMMDevice device;
                    int hr = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, roles[i], out device);
                    if (hr == 0 && device != null) {
                        try {
                            sb.Append(EndpointReport(device, labels[i], true));
                        } finally {
                            Marshal.ReleaseComObject(device);
                        }
                    } else {
                        sb.Append(labels[i]).AppendLine();
                        sb.Append("  Error: ").Append(HResultText(hr)).AppendLine();
                        sb.AppendLine();
                    }
                }

                int listHr = enumerator.EnumAudioEndpoints(
                    EDataFlow.eRender,
                    DeviceState.ACTIVE | DeviceState.DISABLED | DeviceState.NOTPRESENT | DeviceState.UNPLUGGED,
                    out collection);
                if (listHr != 0 || collection == null) {
                    sb.Append("Render endpoint enumeration failed: ").Append(HResultText(listHr)).AppendLine();
                    return sb.ToString();
                }

                uint count;
                Marshal.ThrowExceptionForHR(collection.GetCount(out count));
                sb.Append("Render endpoints (").Append(count).AppendLine(")");
                sb.AppendLine();
                for (uint i = 0; i < count; ++i) {
                    IMMDevice device;
                    Marshal.ThrowExceptionForHR(collection.Item(i, out device));
                    try {
                        sb.Append(EndpointReport(device, string.Format("Render endpoint #{0}", i + 1), false));
                    } finally {
                        Marshal.ReleaseComObject(device);
                    }
                }
                return sb.ToString();
            } finally {
                if (collection != null) Marshal.ReleaseComObject(collection);
                if (enumerator != null) Marshal.ReleaseComObject(enumerator);
            }
        }
    }
}
"@
}

[AudioSpectrogram.PowerShell.AudioEndpointProbe]::GetReport()
