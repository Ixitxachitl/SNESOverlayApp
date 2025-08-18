using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO.Ports;
using System.Management;
using System.Net.WebSockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using SharpDX.DirectInput;

namespace SNESOverlayApp
{
    public class OverlaySettings
    {
        public bool AlwaysOnTop { get; set; } = false;
        public bool SwapXYAB { get; set; } = false;
        public bool LeftStickMapsToDpad { get; set; } = false;
        public bool TriggersMapToBumpers { get; set; } = false;
        public string Source { get; set; } = "None";
        public string Skin { get; set; } = "default";
        public int ZoomFactor { get; set; } = 1;
    }

    public partial class OverlayForm : Form
    {
        [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)]
        private static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst,
            ref Point pptDst, ref Size psize, IntPtr hdcSrc, ref Point pptSrc, int crKey,
            ref BLENDFUNCTION pblend, int dwFlags);

        [DllImport("gdi32.dll", ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr CreateCompatibleDC(IntPtr hDC);

        [DllImport("gdi32.dll", ExactSpelling = true)]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll", ExactSpelling = true)]
        private static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);

        [DllImport("gdi32.dll", ExactSpelling = true)]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("xinput1_4.dll", EntryPoint = "XInputGetState")]
        private static extern int XInputGetState(int dwUserIndex, out XINPUT_STATE pState);

        private const int ULW_ALPHA = 0x00000002;
        private const byte AC_SRC_ALPHA = 0x01;

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct BLENDFUNCTION
        {
            public byte BlendOp;
            public byte BlendFlags;
            public byte SourceConstantAlpha;
            public byte AlphaFormat;
        }

        // --- Input/Device Management ---
        private string serialPortName;
        private bool useXInput = false;
        private CancellationTokenSource inputCancelToken;
        private Task inputTask;
        private Joystick activeDirectInputDevice;
        private SerialPort activeSerialPort;
        private UnifiedInputManager inputManager;
        private InputType currentInputType = InputType.None;
        private object currentInputParam = null;
        private int selectedXInputIndex = 0;
        private Guid selectedDirectInputGuid;
        private Qusb2SnesInputSource qusb;
        private Qusb2SnesInputSource qusb2snesSource;
        private bool isQusb2snesAvailable = false;
        private string qusbComPort = null;
        private List<DirectInputDeviceInfo> cachedDirectInputDevices = new();
        private List<string> cachedPortNames = new();
        private DateTime lastPortScan = DateTime.MinValue;
        private readonly TimeSpan cacheTimeout = TimeSpan.FromSeconds(5);
        private Dictionary<string, string> cachedPortPnpIds = new();
        private bool isPopulatingPorts = false;

        // --- Skin/Layout/Config ---
        private Dictionary<string, List<ButtonInfo>> buttons = new();
        private string currentSkinFolder = "default";
        private bool isGenericSkin = false;
        private string skinType = "snes";

        // --- UI/Menu ---
        private bool swapXYAB = false;
        private bool leftStickMapsToDpad = false;
        private bool TriggersMapToBumpers = false;
        private ToolStripMenuItem portSubmenu;
        private ToolStripMenuItem zoomSubmenu;
        private ContextMenuStrip contextMenu;
        private ToolStripMenuItem alwaysOnTopItem;
        private List<ToolStripMenuItem> portMenuItems = new();
        private Point dragStart;

        // --- Misc State ---
        private string lastQusbResult = null;
        private string lastKnownQusbLabel = null;
        private string lastKnownQusbComPort = null;
        private string lastPopulatedQusbComPort = null;

        // --- Settings ---
        private OverlaySettings settings;
        private static readonly string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");

        // --- Background & Animation State ---
        private Bitmap background;
        private List<Bitmap> backgroundFrames = null;
        private List<int> backgroundDelays = null;
        private int backgroundFrameIndex = 0;
        private DateTime backgroundFrameStart = DateTime.Now;
        private int backgroundLoopCount = 0;
        private List<Bitmap> originalBackgroundFrames = null;
        private Bitmap originalBackground;

        // --- Button/Overlay Image Data ---
        private Dictionary<string, List<Image>> buttonImages = new();
        private Dictionary<string, List<Image>> originalButtonImages = new();
        private Dictionary<string, List<List<Bitmap>>> animatedFrames = new();
        private Dictionary<string, List<int>> animatedDelays = new();
        private Dictionary<string, List<int>> animatedFrameIndices = new();

        // --- Animation Timer ---
        private System.Windows.Forms.Timer animationTimer;
        private DateTime lastFrameUpdateTime = DateTime.Now;

        // --- Composite/Display Logic ---
        private readonly object backgroundLock = new();
        private Point bitmapScreenPosition;

        // --- Button Animation State ---
        private Dictionary<string, DateTime> buttonStartTimes = new();
        private Dictionary<string, int> animatedLoopCounts = new();

        // --- Button Activity State ---
        private HashSet<string> activeButtons = new();
        private HashSet<string> previousActiveButtons = new();

        // --- Zoom ---
        private int zoomFactor = 1;

        // --- SetBitmap callback from main form (injected) ---
        private readonly Action<Bitmap> setBitmapCallback;

        private static readonly Dictionary<(string vid, string pid), string> ArduinoDevices = new()
        {
            { ("2341", "0043"), "Arduino Uno" },
            { ("2341", "0001"), "Arduino Uno (Rev1)" },
            { ("2341", "0069"), "Arduino Uno R4 Minima" },
            { ("2341", "8057"), "Arduino Uno R4 WiFi" },
            { ("2341", "1002"), "Arduino Uno R4 WiFi (Bootloader)" },
            { ("2341", "8036"), "Arduino Leonardo" },
            { ("2341", "8037"), "Arduino Micro" },
            { ("2341", "0042"), "Arduino Mega 2560" },
            { ("2341", "0010"), "Arduino Mega 2560 (Rev1)" },
            { ("2341", "0243"), "Arduino Due" },
            { ("2A03", "0043"), "Arduino Uno (old VID)" },
            // Add others as needed...
        };

        private void LoadSettings()
        {
            if (File.Exists(settingsPath))
            {
                try
                {
                    string json = File.ReadAllText(settingsPath);
                    Debug.WriteLine("Loaded JSON: " + json);
                    var loaded = JsonSerializer.Deserialize<OverlaySettings>(json);

                    if (loaded != null)
                    {
                        Debug.WriteLine("Loaded settings.Source: " + loaded.Source);
                        settings = loaded;
                    }
                    else
                    {
                        Debug.WriteLine("Deserialize returned null, using defaults.");
                        settings = new OverlaySettings();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Error loading settings: " + ex);
                    settings = new OverlaySettings();
                }
            }
            else
            {
                Debug.WriteLine("No settings file, using defaults.");
                settings = new OverlaySettings();
            }
            Debug.WriteLine("AFTER LoadSettings: settings.Source = " + settings.Source);
        }

        private void SaveSettings()
        {
            try
            {
                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(settingsPath, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[EXCEPTION] " + ex.ToString());
            }
        }

        public class ButtonInfo
        {
            public string Name;
            public int X, Y, Width, Height;
            public int OriginalX, OriginalY, OriginalWidth, OriginalHeight;
            public string ImagePath;
            public bool AlwaysVisible = false;
            public Image Image; // ✅ This is the new field
        }

        private void PopulateSkinsMenu(ToolStripMenuItem skinsMenu)
        {
            skinsMenu.DropDownItems.Clear();

            string skinsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Skins");
            if (!Directory.Exists(skinsDir))
                return;

            foreach (var dir in Directory.GetDirectories(skinsDir))
            {
                string folderName = Path.GetFileName(dir);
                string skinXmlPath = Path.Combine(dir, "skin.xml");

                if (!File.Exists(skinXmlPath))
                    continue;

                var item = new ToolStripMenuItem(folderName)
                {
                    Checked = folderName.Equals(currentSkinFolder, StringComparison.OrdinalIgnoreCase)
                };

                item.Click += (s, e) =>
                {
                    LoadSkin(Path.Combine(dir, "skin.xml"), folderName);
                };
                skinsMenu.DropDownItems.Add(item);
            }
        }

        private async Task RefreshDeviceListsAsync()
        {
            List<string> newPortNames = null;
            List<DirectInputDeviceInfo> newDirectInputDevices = null;

            await Task.Run(() =>
            {
                try
                {
                    newPortNames = SerialPort.GetPortNames().ToList();
                    newDirectInputDevices = DirectInputEnumerator.GetConnectedDevices().ToList();
                }
                catch (Exception ex) { 
                    Debug.WriteLine("[EXCEPTION] " + ex.ToString());
                }
            });

            bool changed = !cachedPortNames.SequenceEqual(newPortNames)
                || !cachedDirectInputDevices.Select(d => d.InstanceGuid).SequenceEqual(newDirectInputDevices.Select(d => d.InstanceGuid));

            if (changed)
            {
                cachedPortNames = newPortNames;
                cachedDirectInputDevices = newDirectInputDevices;
                lastPortScan = DateTime.Now;

                // Only update the menu if something changed
                if (portSubmenu.GetCurrentParent() is not null)
                    portSubmenu.GetCurrentParent().BeginInvoke(new Action(PopulatePortMenuSync));
            }
        }

        private void InitializeAnimation()
        {
            if (animationTimer != null)
            {
                //Debug.WriteLine("DISPOSING OLD TIMER");
                animationTimer.Stop();
                animationTimer.Dispose();
                animationTimer = null;
            }
            animationTimer = new System.Windows.Forms.Timer();
            animationTimer.Interval = 100; // ~10fps base tick
            animationTimer.Tick += (s, e) =>
            {
                try
                {
                    //Debug.WriteLine("TICK: " + DateTime.Now.ToString("HH:mm:ss.fff"));
                    var now = DateTime.Now;
                    var elapsed = (int)(now - lastFrameUpdateTime).TotalMilliseconds;
                    lastFrameUpdateTime = now;

                    foreach (var key in animatedFrames.Keys)
                    {
                        for (int i = 0; i < animatedFrames[key].Count; i++)
                        {
                            if (animatedDelays[key][i] <= 0) continue;

                            animatedFrameIndices[key][i]++;
                            if (animatedFrameIndices[key][i] >= animatedFrames[key][i].Count)
                                animatedFrameIndices[key][i] = 0;
                        }
                    }

                    // Animate background
                    if (backgroundFrames != null && backgroundFrames.Count > 1)
                    {
                        int bgElapsed = (int)(now - backgroundFrameStart).TotalMilliseconds;
                        while (bgElapsed >= backgroundDelays[backgroundFrameIndex])
                        {
                            bgElapsed -= backgroundDelays[backgroundFrameIndex];
                            backgroundFrameIndex++;
                            if (backgroundFrameIndex >= backgroundFrames.Count)
                                backgroundFrameIndex = 0; // Loop

                            // Safely assign background (do not dispose if same object)
                            var newFrame = backgroundFrames[backgroundFrameIndex];
                            lock (backgroundLock)
                            {
                                background = newFrame;
                            }
                            backgroundFrameStart = now.AddMilliseconds(-bgElapsed); // Maintain correct timing
                        }
                    }

                    CompositeAndDisplay();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[AnimationTimer] Exception: {ex}");
                    // Optionally: disable timer if you want to halt on error
                    // animationTimer.Stop();
                }
            };
            animationTimer.Start();
        }


        private List<Bitmap> ExtractGifFrames(Image gif, out List<int> frameDelays)
        {
            List<Bitmap> frames = new();
            frameDelays = new();

            var dimension = new FrameDimension(gif.FrameDimensionsList[0]);
            int frameCount = gif.GetFrameCount(dimension);

            byte[]? delayBytes = null;
            try
            {
                var delayItem = gif.GetPropertyItem(0x5100); // Frame delays (in 1/100s of a second)
                delayBytes = delayItem?.Value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[EXCEPTION] " + ex.ToString());
            }

            for (int i = 0; i < frameCount; i++)
            {
                gif.SelectActiveFrame(dimension, i);

                Bitmap rawFrame = new Bitmap(gif.Width, gif.Height, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(rawFrame))
                {
                    g.InterpolationMode = InterpolationMode.NearestNeighbor;
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.SmoothingMode = SmoothingMode.None;
                    g.CompositingMode = CompositingMode.SourceOver;
                    g.DrawImage(gif, new Rectangle(0, 0, rawFrame.Width, rawFrame.Height));
                }

                frames.Add(rawFrame);

                int delay = 100; // default 100ms
                if (delayBytes != null && delayBytes.Length >= 4 * (i + 1))
                {
                    delay = BitConverter.ToInt32(delayBytes, i * 4) * 10;
                    if (delay < 20) delay = 100; // fallback for broken GIFs
                }
                frameDelays.Add(delay);
            }

            return frames;
        }

        private static readonly string?[] BUTTONS_SNES =
        {
            "b", "y", "select", "start",
            "up", "down", "left", "right",
            "a", "x", "l", "r",
            null, null, null, null
        };
        private readonly Dictionary<string, string> SNES_TO_GENERIC = new()
        {
            ["b"] = "b0",
            ["y"] = "b2",
            ["x"] = "b3",
            ["a"] = "b1",
            ["l"] = "b4",
            ["r"] = "b5",
            ["select"] = "b6",
            ["start"] = "b7",
        };

        public class XInputDeviceInfo
        {
            public int UserIndex;
            public string Description;
        }

        public static class XInputDeviceEnumerator
        {
            private const uint DIGCF_PRESENT = 0x2;
            private const uint DIGCF_DEVICEINTERFACE = 0x10;
            private const int SPDRP_DEVICEDESC = 0x0;

            private static readonly Guid GUID_DEVINTERFACE_HID = new Guid("4D1E55B2-F16F-11CF-88CB-001111000030");

            [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetupDiGetClassDevs(
                ref Guid ClassGuid,
                IntPtr Enumerator,
                IntPtr hwndParent,
                uint Flags);

            [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern bool SetupDiEnumDeviceInfo(
                IntPtr DeviceInfoSet,
                int MemberIndex,
                ref SP_DEVINFO_DATA DeviceInfoData);

            [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern bool SetupDiGetDeviceRegistryProperty(
                IntPtr DeviceInfoSet,
                ref SP_DEVINFO_DATA DeviceInfoData,
                int Property,
                out uint PropertyRegDataType,
                byte[] PropertyBuffer,
                uint PropertyBufferSize,
                out uint RequiredSize);

            [DllImport("setupapi.dll", SetLastError = true)]
            private static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

            [StructLayout(LayoutKind.Sequential)]
            private struct SP_DEVINFO_DATA
            {
                public int cbSize;
                public Guid ClassGuid;
                public int DevInst;
                public IntPtr Reserved;
            }

            [DllImport("xinput1_4.dll")]
            private static extern int XInputGetState(int dwUserIndex, out XINPUT_STATE pState);

            [StructLayout(LayoutKind.Sequential)]
            private struct XINPUT_STATE
            {
                public uint dwPacketNumber;
                public XINPUT_GAMEPAD Gamepad;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct XINPUT_GAMEPAD
            {
                public ushort wButtons;
                public byte bLeftTrigger;
                public byte bRightTrigger;
                public short sThumbLX;
                public short sThumbLY;
                public short sThumbRX;
                public short sThumbRY;
            }

            public static List<XInputDeviceInfo> GetConnectedDevices()
            {
                var results = new List<XInputDeviceInfo>();

                for (int i = 0; i < 4; i++)
                {
                    if (XInputGetState(i, out _) == 0)
                    {
                        string desc = LookupDeviceDescription(i);
                        results.Add(new XInputDeviceInfo { UserIndex = i, Description = desc });
                    }
                }

                return results;
            }

            private static string LookupDeviceDescription(int userIndex)
            {
                Guid hidGuid = GUID_DEVINTERFACE_HID;
                IntPtr hDevInfo = SetupDiGetClassDevs(ref hidGuid, IntPtr.Zero, IntPtr.Zero, DIGCF_PRESENT);
                if (hDevInfo == IntPtr.Zero) return $"XInput Device {userIndex}";

                try
                {
                    SP_DEVINFO_DATA devInfoData = new SP_DEVINFO_DATA
                    {
                        cbSize = Marshal.SizeOf<SP_DEVINFO_DATA>()
                    };

                    for (int i = 0; ; i++)
                    {
                        if (!SetupDiEnumDeviceInfo(hDevInfo, i, ref devInfoData)) break;

                        byte[] buffer = new byte[1024];
                        if (SetupDiGetDeviceRegistryProperty(hDevInfo, ref devInfoData, SPDRP_DEVICEDESC, out _, buffer, (uint)buffer.Length, out _))
                        {
                            string name = Encoding.Unicode.GetString(buffer).TrimEnd('\0');
                            if (name.Contains("XInput", StringComparison.OrdinalIgnoreCase))
                                return name + $" (User {userIndex})";
                        }
                    }
                }
                finally
                {
                    SetupDiDestroyDeviceInfoList(hDevInfo);
                }

                return $"XInput Controller (User {userIndex})";
            }
        }

        public class DirectInputDeviceInfo
        {
            public int DeviceIndex;
            public string Description;
            public Guid InstanceGuid;
            public Guid ProductGuid;
            public string PnpId;
            public string InstanceName;
        }

        public static class DirectInputEnumerator
        {
            public static List<DirectInputDeviceInfo> GetConnectedDevices()
            {
                var results = new List<DirectInputDeviceInfo>();
                var directInput = new DirectInput();
                int index = 0;

                // Gather WMI PNP device info
                var wmiDevices = new Dictionary<string, string>(); // [Name] = PnpId
                try
                {
                    using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity");
                    foreach (var obj in searcher.Get())
                    {
                        string? name = obj["Name"]?.ToString();
                        string? pnpId = obj["PNPDeviceID"]?.ToString();
                        if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(pnpId))
                            wmiDevices[name] = pnpId;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[EXCEPTION] " + ex.ToString());
                }

                foreach (var device in directInput.GetDevices(DeviceClass.GameControl, DeviceEnumerationFlags.AttachedOnly))
                {
                    string matchedPnp = "unknown";

                    // First: exact match on name, USB-only
                    foreach (var kvp in wmiDevices)
                    {
                        if (!kvp.Value.StartsWith("USB\\", StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (kvp.Key.Equals(device.InstanceName, StringComparison.OrdinalIgnoreCase))
                        {
                            matchedPnp = kvp.Value;
                            break;
                        }
                    }

                    // Fallback: fuzzy match if exact match failed
                    if (matchedPnp == "unknown")
                    {
                        foreach (var kvp in wmiDevices)
                        {
                            if (!kvp.Value.StartsWith("USB\\", StringComparison.OrdinalIgnoreCase))
                                continue;

                            if (kvp.Key.Contains("iBuffalo", StringComparison.OrdinalIgnoreCase) ||
                                kvp.Key.Contains("2-axis 8-button", StringComparison.OrdinalIgnoreCase) ||
                                kvp.Key.Contains("HID-compliant game controller", StringComparison.OrdinalIgnoreCase) ||
                                kvp.Key.Contains(device.InstanceName, StringComparison.OrdinalIgnoreCase))
                            {
                                matchedPnp = kvp.Value;
                                break;
                            }
                        }
                    }

                    results.Add(new DirectInputDeviceInfo
                    {
                        DeviceIndex = index++,
                        Description = device.InstanceName,
                        InstanceName = device.InstanceName,
                        InstanceGuid = device.InstanceGuid,
                        ProductGuid = device.ProductGuid,
                        PnpId = matchedPnp
                    });
                }

                return results;
            }

        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x80000; // WS_EX_LAYERED
                return cp;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct XINPUT_STATE
        {
            public uint dwPacketNumber;
            public XINPUT_GAMEPAD Gamepad;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct XINPUT_GAMEPAD
        {
            public ushort wButtons;
            public byte bLeftTrigger;
            public byte bRightTrigger;
            public short sThumbLX;
            public short sThumbLY;
            public short sThumbRX;
            public short sThumbRY;
        }

        private readonly Dictionary<ushort, string> XINPUT_BUTTON_MAP = new()
        {
            [0x1000] = "b",       // A → SNES B
            [0x2000] = "a",       // B → SNES A
            [0x4000] = "y",       // X → SNES Y
            [0x8000] = "x",       // Y → SNES X
            [0x0100] = "l",       // LB → SNES L
            [0x0200] = "r",       // RB → SNES R
            [0x0020] = "select",  // Back
            [0x0010] = "start",   // Start
            [0x0001] = "up",
            [0x0002] = "down",
            [0x0004] = "left",
            [0x0008] = "right"
        };

        private void CacheDeviceLists()
        {
            cachedPortNames = SerialPort.GetPortNames().ToList();
            cachedDirectInputDevices = DirectInputEnumerator.GetConnectedDevices().ToList();
            lastPortScan = DateTime.Now;

            // Update PNP ID cache here:
            cachedPortPnpIds.Clear();
            foreach (var port in cachedPortNames)
            {
                string pnpId = GetPnpDeviceIdForPort(port);
                if (!string.IsNullOrEmpty(pnpId))
                    cachedPortPnpIds[port] = pnpId;
            }
        }

        public OverlayForm(string portName)
        {
            LoadSettings();
            CacheDeviceLists(); // Always before SetPort or device menu
            InitializeComponent(); // Setup controls before anything may fire events

            // Fields from settings (set BEFORE SetPort)
            swapXYAB = settings.SwapXYAB;
            leftStickMapsToDpad = settings.LeftStickMapsToDpad;
            TriggersMapToBumpers = settings.TriggersMapToBumpers;
            serialPortName = settings.Source;
            currentSkinFolder = settings.Skin;
            zoomFactor = settings.ZoomFactor;
            this.TopMost = settings.AlwaysOnTop;
            ApplyZoom();

            inputManager = new UnifiedInputManager(leftStickMapsToDpad, TriggersMapToBumpers);
            inputManager.OnInputReceived += (bitmask, _, _) =>
            {
                if (this.IsHandleCreated && !this.IsDisposed)
                {
                    if (this.InvokeRequired)
                        this.BeginInvoke(() => UpdateButtons(bitmask));
                    else
                        UpdateButtons(bitmask);
                }
            };

            SetupContextMenu();

            bool restored = false;

            // Now set the port (after everything above is initialized)
            if (!string.IsNullOrEmpty(settings.Source))
            {
                var inputType = ParseInputType(settings.Source);
                switch (inputType)
                {
                    case InputType.XInput:
                        int userIdx = 0;
                        var parts = settings.Source.Split(':');
                        if (parts.Length == 2 && int.TryParse(parts[1], out int idx))
                            userIdx = idx;
                        // check available XInput devices for userIdx...
                        SetPort(InputType.XInput, userIdx);
                        restored = true;
                        break;

                    case InputType.DirectInput:
                        var dev = cachedDirectInputDevices
                            .FirstOrDefault(d => d.InstanceGuid.ToString().Equals(settings.Source, StringComparison.OrdinalIgnoreCase));
                        if (dev != null)
                        {
                            SetPort(InputType.DirectInput, Tuple.Create(dev.InstanceGuid, dev.PnpId));
                            restored = true;
                        }
                        break;

                    case InputType.Com:
                        if (cachedPortNames.Contains(settings.Source))
                        {
                            SetPort(InputType.Com, settings.Source);
                            restored = true;
                        }
                        break;

                    case InputType.Qusb2Snes:
                        // Similar logic for QUSB2SNES
                        SetPort(InputType.Qusb2Snes, settings.Source);
                        restored = true;
                        break;

                    case InputType.None:
                    default:
                        break;
                }
            }

            if (!restored)
            {
                SetPort(InputType.None, "None");
                settings.Source = "None";
                SaveSettings();
            }

            // The rest of your field setup and events (icon, mouse, etc)
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = true;

            var assembly = Assembly.GetExecutingAssembly();
            using Stream stream = assembly.GetManifestResourceStream("SNESOverlayApp.Resources.icon.ico");
            if (stream != null) this.Icon = new Icon(stream);

            this.MouseDown += Form_MouseDown;
            this.MouseMove += Form_MouseMove;
            this.MouseUp += Form_MouseUp;

            string skinToLoad = settings.Skin ?? "default";
            string skinPath = Path.Combine("Skins", skinToLoad, "skin.xml");
            if (!File.Exists(skinPath))
            {
                skinPath = Path.Combine("Skins", "default", "skin.xml");
                skinToLoad = "default";
            }
            LoadSkin(skinPath, skinToLoad);
            InitializeAnimation();
        }

        private void SetBitmap(Bitmap bitmap)
        {
            if (!this.IsHandleCreated || this.IsDisposed)
                return;

            Graphics screenGraphics = null;
            IntPtr screenDC = IntPtr.Zero;
            IntPtr memDC = IntPtr.Zero;
            IntPtr hBitmap = IntPtr.Zero;
            IntPtr oldBitmap = IntPtr.Zero;
            try
            {
                screenGraphics = Graphics.FromHwnd(IntPtr.Zero);
                screenDC = screenGraphics.GetHdc();
                memDC = CreateCompatibleDC(screenDC);
                hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
                oldBitmap = SelectObject(memDC, hBitmap);

                Size size = bitmap.Size;
                Point source = new Point(0, 0);
                Point top = bitmapScreenPosition;

                BLENDFUNCTION blend = new BLENDFUNCTION
                {
                    BlendOp = 0,
                    BlendFlags = 0,
                    SourceConstantAlpha = 255,
                    AlphaFormat = AC_SRC_ALPHA
                };

                if (!this.IsDisposed && this.IsHandleCreated)
                {
                    UpdateLayeredWindow(this.Handle, screenDC, ref top, ref size, memDC,
                        ref source, 0, ref blend, ULW_ALPHA);
                }
            }
            catch (ObjectDisposedException)
            {
                // ignore
            }
            finally
            {
                if (oldBitmap != IntPtr.Zero)
                    SelectObject(memDC, oldBitmap);
                if (hBitmap != IntPtr.Zero)
                    DeleteObject(hBitmap);
                if (memDC != IntPtr.Zero)
                    DeleteDC(memDC);
                if (screenGraphics != null && screenDC != IntPtr.Zero)
                    screenGraphics.ReleaseHdc(screenDC);
                if (screenGraphics != null)
                    screenGraphics.Dispose();
            }
        }

        private void CompositeAndDisplay()
        {
            if (background == null)
            {
                //Debug.WriteLine("Background is null in CompositeAndDisplay");
                return;
            }
            try
            {
                var test = background.Width; // Will throw if disposed!
            }
            catch (Exception ex)
            {
                //Debug.WriteLine("Background is invalid/disposed in CompositeAndDisplay: " + ex);
                return; // Don't draw with it!
            }
            lock (backgroundLock)
            {
                if (background == null) return;

                Bitmap composite = new Bitmap(background.Width, background.Height, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(composite))
                {
                    g.InterpolationMode = InterpolationMode.NearestNeighbor;
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.SmoothingMode = SmoothingMode.None;

                    // Draw background using SourceCopy to initialize the canvas
                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.DrawImage(background, 0, 0);

                    // Switch to SourceOver for button layering
                    g.CompositingMode = CompositingMode.SourceOver;

                    foreach (var name in activeButtons)
                    {
                        if (!buttons.TryGetValue(name, out var btns)) continue;

                        foreach (var btn in btns)
                        {
                            var imgToDraw = btn.Image;

                            if (animatedFrames.TryGetValue(name, out var animLists) &&
                                animatedDelays.TryGetValue(name, out var delays) &&
                                buttonStartTimes.TryGetValue(name, out var startTime))
                            {
                                if (animLists.Count > 0 && animLists[0].Count > 0)
                                {
                                    var frameList = animLists[0];
                                    var frameDelays = delays;

                                    int totalDuration = frameDelays.Sum();
                                    int elapsed = (int)(DateTime.Now - startTime).TotalMilliseconds;

                                    int loopCount = animatedLoopCounts.TryGetValue(name, out var count) ? count : 0;

                                    if (loopCount == 0) // infinite
                                    {
                                        elapsed %= totalDuration;
                                    }
                                    else
                                    {
                                        int maxDuration = totalDuration * loopCount;
                                        if (elapsed >= maxDuration)
                                            elapsed = maxDuration - 1; // clamp to final frame
                                    }

                                    int cumulative = 0;
                                    int frameIndex = 0;
                                    while (frameIndex < frameDelays.Count)
                                    {
                                        cumulative += frameDelays[frameIndex];
                                        if (elapsed < cumulative)
                                            break;
                                        frameIndex++;
                                    }

                                    if (frameIndex >= frameList.Count)
                                        frameIndex = frameList.Count - 1;

                                    imgToDraw = frameList[frameIndex];
                                }
                            }

                            g.DrawImage(imgToDraw, btn.X, btn.Y, btn.Width, btn.Height);
                        }
                    }
                }

                SetBitmap(composite);
                composite.Dispose();
            }
        }

        private int GetGifLoopCount(string path)
        {
            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                using var br = new BinaryReader(fs);

                while (fs.Position < fs.Length)
                {
                    byte b = br.ReadByte();

                    if (b == 0x21) // Extension Introducer
                    {
                        byte label = br.ReadByte();
                        if (label == 0xFF) // Application Extension
                        {
                            int blockSize = br.ReadByte(); // should be 11
                            string appId = new string(br.ReadChars(8));
                            string authCode = new string(br.ReadChars(3));
                            string fullAppId = appId + authCode;

                            if (fullAppId == "NETSCAPE2.0" || fullAppId == "ANIMEXTS1.0")
                            {
                                byte subBlockSize = br.ReadByte(); // usually 3
                                byte subId = br.ReadByte();        // should be 1
                                ushort loops = br.ReadUInt16();    // loop count
                                br.ReadByte(); // block terminator

                                return loops; // 0 means infinite
                            }
                        }
                        else
                        {
                            // skip other extensions
                            SkipSubBlocks(br);
                        }
                    }
                    else if (b == 0x3B) // Trailer
                    {
                        break;
                    }
                    else
                    {
                        // Image block or other data — skip over
                        SkipImageOrBlock(br, b);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[EXCEPTION] " + ex.ToString());
            }

            return -1; // unknown
        }

        private void SkipSubBlocks(BinaryReader br)
        {
            byte size;
            while ((size = br.ReadByte()) != 0)
                br.ReadBytes(size);
        }

        private void SkipImageOrBlock(BinaryReader br, byte introducer)
        {
            if (introducer == 0x2C) // Image Descriptor
            {
                br.ReadBytes(9); // descriptor
                byte packed = br.ReadByte();
                if ((packed & 0x80) != 0)
                {
                    int tableSize = 3 * (1 << ((packed & 0x07) + 1));
                    br.ReadBytes(tableSize);
                }
                br.ReadByte(); // LZW min code size
                SkipSubBlocks(br);
            }
            else if (introducer == 0x21) // Extension Introducer
            {
                br.ReadByte(); // Skip label
                SkipSubBlocks(br);
            }
            else
            {
                // Instead of throwing, try to safely skip, or just return
                // br.ReadByte(); // Optional: advance one byte and hope for resync
                // Or, just return;
            }
        }

        private Bitmap ScaleBitmapNearest(Bitmap src, int width, int height)
        {
            Bitmap dest = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(dest))
            {
                g.InterpolationMode = InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.SmoothingMode = SmoothingMode.None;
                g.CompositingMode = CompositingMode.SourceOver; // Allow proper blend layering
                g.CompositingQuality = CompositingQuality.HighSpeed;
                g.DrawImage(src, new Rectangle(0, 0, width, height));
            }
            return dest;
        }

        private void LoadSkin(string xmlPath, string folderName)
        {
            animationTimer?.Stop();
            currentSkinFolder = folderName;
            settings.Skin = currentSkinFolder;

            string skinDir = Path.GetDirectoryName(xmlPath)!;

            var doc = XDocument.Load(xmlPath);
            var root = doc.Root;
            skinType = root?.Attribute("type")?.Value?.ToLowerInvariant() ?? "snes";

            // 🧹 Clear state
            buttons.Clear();
            originalButtonImages.Clear();
            activeButtons.Clear();
            previousActiveButtons.Clear();

            // Dispose old background animation frames
            if (backgroundFrames != null)
            {
                foreach (var bmp in backgroundFrames)
                {
                    //Debug.WriteLine($"Disposing bitmap: {bmp.GetHashCode()} - (is current background? {object.ReferenceEquals(bmp, background)})");
                    bmp.Dispose();
                }
                backgroundFrames = null;
            }
            if (originalBackgroundFrames != null)
            {
                foreach (var bmp in originalBackgroundFrames)
                {
                    //Debug.WriteLine($"Disposing bitmap: {bmp.GetHashCode()} - (is current background? {object.ReferenceEquals(bmp, background)})");
                    bmp.Dispose();
                }
                originalBackgroundFrames = null;
            }
            backgroundDelays = null;
            backgroundFrameIndex = 0;

            // Dispose old animatedFrames (button GIFs)
            foreach (var animList in animatedFrames.Values)
                foreach (var frameList in animList)
                    foreach (var frame in frameList)
                        frame?.Dispose();
            animatedFrames.Clear();
            animatedDelays.Clear();
            animatedFrameIndices.Clear();
            animatedLoopCounts.Clear();

            // Dispose old button images
            foreach (var imgList in buttonImages.Values)
                foreach (var img in imgList)
                    img?.Dispose();
            buttonImages.Clear();

            // Dispose old original button images
            foreach (var imgList in originalButtonImages.Values)
                foreach (var img in imgList)
                    img?.Dispose();
            originalButtonImages.Clear();

            if (background != null && background != originalBackground)
            {
                lock (backgroundLock)
                {
                    background.Dispose();
                }
            }
            background = null;
            originalBackground = null;

            // Dispose old background animation frames
            if (backgroundFrames != null)
            {
                foreach (var bmp in backgroundFrames)
                {
                    bmp?.Dispose();
                }
                backgroundFrames = null;
                backgroundDelays = null;
                backgroundFrameIndex = 0;
            }

            var bgPath = GetBackgroundImagePath(xmlPath);
            if (File.Exists(bgPath))
            {
                using (Image bgImg = Image.FromFile(bgPath))
                {
                    if (ImageAnimator.CanAnimate(bgImg))
                    {
                        // Extract original GIF frames
                        List<int> delays;
                        originalBackgroundFrames = ExtractGifFrames(bgImg, out delays); // original size
                        backgroundDelays = delays;
                        backgroundLoopCount = GetGifLoopCount(bgPath); // 0 = infinite
                        backgroundFrameIndex = 0;
                        backgroundFrameStart = DateTime.Now;

                        // Now, scale each frame to current zoom and set as backgroundFrames
                        int newWidth = (int)(originalBackgroundFrames[0].Width * zoomFactor);
                        int newHeight = (int)(originalBackgroundFrames[0].Height * zoomFactor);

                        backgroundFrames = new List<Bitmap>();
                        foreach (var frame in originalBackgroundFrames)
                        {
                            Bitmap scaled = new Bitmap(frame, newWidth, newHeight);
                            backgroundFrames.Add(scaled);
                        }
                        // Immediately display the first frame at correct zoom
                        lock (backgroundLock)
                        {
                            background = backgroundFrames[0];
                        }
                        this.ClientSize = background.Size;
                    }
                    else
                    {
                        // Not animated, dispose old originals
                        originalBackgroundFrames?.ForEach(bmp => bmp.Dispose());
                        originalBackgroundFrames = null;
                        backgroundFrames?.ForEach(bmp => bmp.Dispose());
                        backgroundFrames = null;

                        originalBackground = new Bitmap(bgImg); // store the original
                        int newWidth = (int)(originalBackground.Width * zoomFactor);
                        int newHeight = (int)(originalBackground.Height * zoomFactor);

                        background?.Dispose();
                        background = new Bitmap(originalBackground, newWidth, newHeight);
                        this.ClientSize = background.Size;
                    }
                }
            }

            int requestedZoom = zoomFactor;
            zoomFactor = 1;

            var elements = new[] { "button", "stick", "analog" };
            foreach (var type in elements)
            {
                foreach (var elem in doc.Descendants(type))
                {
                    string name = elem.Attribute("name")?.Value
                                ?? $"{elem.Attribute("xname")?.Value}_{elem.Attribute("yname")?.Value}";
                    if (string.IsNullOrEmpty(name)) continue;

                    string relPath = elem.Attribute("image")?.Value;
                    if (string.IsNullOrEmpty(relPath)) continue;

                    string imagePath = Path.Combine(Path.GetDirectoryName(xmlPath), relPath);
                    if (!File.Exists(imagePath)) continue;

                    int x = int.Parse(elem.Attribute("x")?.Value ?? "0");
                    int y = int.Parse(elem.Attribute("y")?.Value ?? "0");

                    Image img = Image.FromFile(imagePath);
                    int width = int.TryParse(elem.Attribute("width")?.Value, out var w) ? w : img.Width;
                    int height = int.TryParse(elem.Attribute("height")?.Value, out var h) ? h : img.Height;

                    var buttonInfo = new ButtonInfo
                    {
                        Name = name,
                        X = x,
                        Y = y,
                        Width = width,
                        Height = height,
                        OriginalX = x,
                        OriginalY = y,
                        OriginalWidth = width,
                        OriginalHeight = height,
                        ImagePath = imagePath,
                        AlwaysVisible = type != "button"
                    };

                    if (!buttons.ContainsKey(name))
                        buttons[name] = new List<ButtonInfo>();
                    buttons[name].Add(buttonInfo);

                    if (!originalButtonImages.ContainsKey(name))
                        originalButtonImages[name] = new List<Image>();

                    if (!buttonImages.ContainsKey(name))
                    {
                        if (ImageAnimator.CanAnimate(img))
                        {
                            List<int> delays;
                            var rawFrames = ExtractGifFrames(img, out delays);
                            var scaledFrames = rawFrames.Select(f => ScaleBitmapNearest(f, (int)(width * zoomFactor), (int)(height * zoomFactor))).ToList();

                            int loopCount = GetGifLoopCount(imagePath);
                            if (loopCount == -1) loopCount = 1; // Default to 1 if unknown

                            animatedFrames[name] = new List<List<Bitmap>> { scaledFrames };
                            animatedDelays[name] = new List<int>(delays);
                            animatedLoopCounts[name] = loopCount;

                            animatedFrameIndices[name] = new List<int> { 0 };

                            var fallback = new Bitmap((int)(width * zoomFactor), (int)(height * zoomFactor), PixelFormat.Format32bppArgb);
                            using (Graphics g = Graphics.FromImage(fallback))
                            {
                                g.InterpolationMode = InterpolationMode.NearestNeighbor;
                                g.PixelOffsetMode = PixelOffsetMode.Half;
                                g.SmoothingMode = SmoothingMode.None;
                                g.CompositingMode = CompositingMode.SourceOver;
                                g.DrawImage(scaledFrames[0], 0, 0);
                            }

                            buttonImages[name] = new List<Image> { fallback };
                            originalButtonImages[name].Add(img);
                        }
                        else
                        {
                            var scaled = ScaleBitmapNearest((Bitmap)img, (int)(width * zoomFactor), (int)(height * zoomFactor));
                            buttonInfo.Image = scaled;
                            originalButtonImages[name].Add(img);
                        }
                    }

                    if (buttonInfo.AlwaysVisible)
                        activeButtons.Add(name);
                    if (type != "button")
                        activeButtons.Add(name);
                }
            }
            zoomFactor = requestedZoom;
            ApplyZoom();
            CompositeAndDisplay();
            InitializeAnimation();
            animationTimer?.Start();
            if (!this.Visible) // only center if it's the first time (startup)
                CenterOnScreen();
        }

        private string GetBackgroundImagePath(string xmlPath)
        {
            var doc = XDocument.Load(xmlPath);
            var bg = doc.Descendants("background").FirstOrDefault();
            string? rel = bg?.Attribute("image")?.Value;
            return string.IsNullOrEmpty(rel) ? string.Empty : Path.Combine(Path.GetDirectoryName(xmlPath), rel);
        }

        private void UpdateButtons(bool[] bitmask, float leftStickX = 0, float leftStickY = 0)
        {
            const float deadzone = 0.2f;

            if (leftStickMapsToDpad)
            {
                if (leftStickX < -deadzone) bitmask[6] = true; // Left
                if (leftStickX > deadzone) bitmask[7] = true; // Right
                if (leftStickY < -deadzone) bitmask[4] = true; // Up
                if (leftStickY > deadzone) bitmask[5] = true; // Down
            }

            var newActive = new HashSet<string>(
                buttons.SelectMany(kvp => kvp.Value)
                       .Where(info => info.AlwaysVisible)
                       .Select(info => info.Name));

            for (int i = 0; i < BUTTONS_SNES.Length && i < bitmask.Length; i++)
            {
                if (!bitmask[i]) continue;
                string? logical = BUTTONS_SNES[i];
                if (string.IsNullOrEmpty(logical)) continue;

                string key = logical;

                // Apply XYAB swap if enabled
                if (swapXYAB)
                {
                    if (key == "x") key = "y";
                    else if (key == "y") key = "x";
                    else if (key == "a") key = "b";
                    else if (key == "b") key = "a";
                }

                // Apply generic skin mapping (uses original logical name, not swapped key)
                if (skinType == "generic" && SNES_TO_GENERIC.TryGetValue(logical, out var mapped))
                    key = mapped;

                if (!previousActiveButtons.Contains(key))
                    buttonStartTimes[key] = DateTime.Now;

                if (buttons.ContainsKey(key))
                    newActive.Add(key);
            }

            if (!newActive.SetEquals(previousActiveButtons))
            {
                activeButtons = newActive;
                previousActiveButtons = new HashSet<string>(newActive);
                CompositeAndDisplay();
            }

            foreach (var key in buttonStartTimes.Keys.ToList())
            {
                if (!activeButtons.Contains(key))
                    buttonStartTimes.Remove(key);
            }
        }

        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                dragStart = e.Location;
        }

        private void Form_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                bitmapScreenPosition.X += e.X - dragStart.X;
                bitmapScreenPosition.Y += e.Y - dragStart.Y;
                CompositeAndDisplay();
            }
        }

        private void Form_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
                contextMenu.Show(this, e.Location);
        }

        private InputType ParseInputType(string source)
        {
            if (source.StartsWith("XInput")) return InputType.XInput;
            if (Guid.TryParse(source, out _)) return InputType.DirectInput;
            if (source.StartsWith("QUSB2SNES")) return InputType.Qusb2Snes;
            if (source.StartsWith("COM")) return InputType.Com;
            if (source == "None") return InputType.None;
            return InputType.None;
        }

        private void SetupContextMenu()
        {
            contextMenu = new ContextMenuStrip();

            contextMenu.Opening += async (object? s, CancelEventArgs e) => {
                animationTimer?.Stop();
                PopulatePortMenuSync();
                await LookupAndCacheQusb2SnesAsync();
            };

            contextMenu.Closed += (s, e) =>
            {
                //Debug.WriteLine("STARTED: " + DateTime.Now.ToString("HH:mm:ss.fff"));
                animationTimer?.Start();
                SaveSettings();
            };

            alwaysOnTopItem = new ToolStripMenuItem("Always on Top")
            {
                CheckOnClick = true,
                Checked = this.TopMost
            };

            alwaysOnTopItem.CheckedChanged += (s, e) =>
            {
                settings.AlwaysOnTop = alwaysOnTopItem.Checked;
                this.TopMost = alwaysOnTopItem.Checked;
                SaveSettings();
            };

            var swapItem = new ToolStripMenuItem("Swap X/Y and A/B")
            {
                CheckOnClick = true,
                Checked = swapXYAB
            };
            swapItem.CheckedChanged += (s, e) =>
            {
                settings.SwapXYAB = swapItem.Checked;
                swapXYAB = swapItem.Checked;
                SaveSettings();
            };

            var stickToDpadItem = new ToolStripMenuItem("Map Left Stick to D-Pad")
            {
                Checked = leftStickMapsToDpad,
                CheckOnClick = true
            };
            stickToDpadItem.CheckedChanged += (s, e) =>
            {
                settings.LeftStickMapsToDpad = stickToDpadItem.Checked;
                leftStickMapsToDpad = stickToDpadItem.Checked;
                RestartInputManager();
                SaveSettings();
            };

            var triggersToBumpersItem = new ToolStripMenuItem("Map Triggers to Bumpers")
            {
                Checked = settings.TriggersMapToBumpers,
                CheckOnClick = true
            };
            triggersToBumpersItem.CheckedChanged += (s, e) =>
            {
                settings.TriggersMapToBumpers = triggersToBumpersItem.Checked;
                TriggersMapToBumpers = triggersToBumpersItem.Checked;
                RestartInputManager(); // will pass updated setting
                SaveSettings();
            };

            portSubmenu = new ToolStripMenuItem("Select Input");
            portSubmenu.DropDownOpening += async (s, e) =>
            {
                PopulatePortMenuSync(); // Always shows cached device list instantly (no lag)
                await RefreshDeviceListsAsync(); // Runs scan in background and updates menu when done
            };

            var skinsMenu = new ToolStripMenuItem("Skins");
            PopulateSkinsMenu(skinsMenu);
            skinsMenu.DropDownOpening += (s, e) => PopulateSkinsMenu(skinsMenu);

            zoomSubmenu = new ToolStripMenuItem("Zoom");
            AddZoomOption("1x", 1);
            AddZoomOption("2x", 2);
            AddZoomOption("4x", 4);

            var closeItem = new ToolStripMenuItem("Close");
            closeItem.Click += (s, e) => this.Close();

            contextMenu.Items.Add(alwaysOnTopItem);
            contextMenu.Items.Add(swapItem);
            contextMenu.Items.Add(stickToDpadItem);
            contextMenu.Items.Add(triggersToBumpersItem);
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(portSubmenu);
            contextMenu.Items.Add(skinsMenu);
            contextMenu.Items.Add(zoomSubmenu);
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(closeItem);
        }

        private void RestartInputManager()
        {
            inputManager?.StopCurrent();
            inputManager = new UnifiedInputManager(leftStickMapsToDpad, TriggersMapToBumpers);
            inputManager.OnInputReceived += (bitmask, _, _) =>
            {
                if (this.IsHandleCreated && !this.IsDisposed)
                {
                    if (this.InvokeRequired)
                        this.BeginInvoke(() => UpdateButtons(bitmask));
                    else
                        UpdateButtons(bitmask);
                }
            };

            switch (ParseInputType(serialPortName))
            {
                case InputType.XInput:
                    inputManager.SetSource(InputType.XInput, serialPortName, selectedXInputIndex);
                    break;

                case InputType.DirectInput:
                    // we already know serialPortName is a GUID string
                    var pnp = GetPnpDeviceIdForPort(serialPortName);
                    inputManager.SetSource(InputType.DirectInput, serialPortName, Tuple.Create(selectedDirectInputGuid, pnp));
                    break;

                case InputType.Qusb2Snes:
                    inputManager.SetSource(InputType.Qusb2Snes, serialPortName);
                    break;

                case InputType.Com:
                    inputManager.SetSource(InputType.Com, serialPortName);
                    break;

                default: // InputType.None
                         // nothing to do
                    break;
            }
        }


        private void PopulatePortMenuSync()
        {
            portMenuItems.Clear();
            portSubmenu.DropDownItems.Clear();

            void Add(string label, Action onClick, bool isChecked, string tag = null)
            {
                var item = new ToolStripMenuItem(label)
                {
                    Checked = isChecked
                };
                if (tag != null)
                    item.Tag = tag;

                item.Click += (s, e) =>
                {
                    onClick();
                    RefreshPortMenuChecks();
                };

                portMenuItems.Add(item);
            }

            // None option
            Add("None", () => SetPort(InputType.None, "None"), serialPortName == "None", "None");

            // XInput devices
            foreach (var device in XInputDeviceEnumerator.GetConnectedDevices())
            {
                int userIndex = device.UserIndex;
                string label = device.Description;
                string tag = $"XInput_{userIndex}";
                bool isChecked = serialPortName == $"XInput:{userIndex}";
                Add(label, () => SetPort(InputType.XInput, userIndex), isChecked, tag);
            }

            var qusbPorts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> seenPorts = new(StringComparer.OrdinalIgnoreCase);

            // QUSB2SNES
            if (!string.IsNullOrEmpty(lastKnownQusbComPort))
            {
                qusbPorts.Add(lastKnownQusbComPort);
                string tag = $"QUSB2SNES_{lastKnownQusbComPort}";
                Add(
                    $"QUSB2SNES (FXPak Pro) {lastKnownQusbComPort}",
                    () => {
                        isQusb2snesAvailable = true;
                        qusbComPort = lastKnownQusbComPort;
                        SetPort(InputType.Qusb2Snes, lastKnownQusbComPort);
                    },
                    serialPortName == lastKnownQusbComPort,
                    tag
                );
            }

            // COM Ports
            foreach (string port in cachedPortNames)
            {
                if (seenPorts.Contains(port)) continue;
                seenPorts.Add(port);

                cachedPortPnpIds.TryGetValue(port, out string pnpId);
                string label;

                if (!string.IsNullOrEmpty(pnpId))
                {
                    string vid = ExtractValue(pnpId, "VID_");
                    string pid = ExtractValue(pnpId, "PID_");
                    string arduinoName = GetCommonName(vid, pid);

                    label = !string.IsNullOrEmpty(arduinoName) ? $"{port} ({arduinoName})" : port;
                }
                else
                {
                    label = port;
                }

                if (qusbPorts.Contains(port))
                    continue;

                Add(
                    label,
                    () => SetPort(InputType.Com, port),
                    serialPortName == port,
                    $"COM_{port}"
                );
            }

            // DirectInput
            foreach (var dev in cachedDirectInputDevices)
            {
                string label = $"{dev.InstanceName} [{dev.InstanceGuid.ToString()[..8]}]";
                Guid guidCopy = dev.InstanceGuid;
                string pnpCopy = dev.PnpId;
                string tag = $"DirectInput_{guidCopy}";

                Add(label,
                    () => {
                        selectedDirectInputGuid = guidCopy;
                        serialPortName = guidCopy.ToString();
                        SetPort(InputType.DirectInput, Tuple.Create(guidCopy, pnpCopy));
                        SaveSettings();
                        RefreshPortMenuChecks();
                    },
                    Guid.TryParse(serialPortName, out var guid) && guid == guidCopy,
                    tag
                );
            }

            portSubmenu.DropDownItems.Clear();
            portSubmenu.DropDownItems.AddRange(portMenuItems.ToArray());
        }

        private async Task LookupAndCacheQusb2SnesAsync()
        {
            string foundDeviceName = null;
            string foundComPort = null;
            string foundLabel = null;

            try
            {
                using var ws = new ClientWebSocket();
                var cts = new CancellationTokenSource(1000);
                await ws.ConnectAsync(new Uri("ws://localhost:8080"), cts.Token);

                string json = JsonSerializer.Serialize(new { Opcode = "DeviceList", Space = "SNES" });
                byte[] buffer = Encoding.UTF8.GetBytes(json);
                await ws.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Text, true, CancellationToken.None);

                byte[] readBuffer = new byte[2048];
                var result = await ws.ReceiveAsync(new ArraySegment<byte>(readBuffer), CancellationToken.None);
                var response = Encoding.UTF8.GetString(readBuffer, 0, result.Count);
                var doc = JsonDocument.Parse(response);

                var results = doc.RootElement.GetProperty("Results").EnumerateArray().Select(x => x.GetString()).Where(x => !string.IsNullOrEmpty(x)).ToList();
                if (results.Count > 0)
                {
                    foundDeviceName = results[0];
                    var match = Regex.Match(foundDeviceName, @"COM\d+", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        foundComPort = match.Value;
                        foundLabel = $"QUSB2SNES ({foundDeviceName})";
                    }
                }
            }
            catch
            {
                // Not found, do nothing
            }

            if (!string.IsNullOrEmpty(foundLabel) && !string.IsNullOrEmpty(foundComPort))
            {
                bool needsUpdate = foundComPort != lastKnownQusbComPort || foundLabel != lastKnownQusbLabel;

                lastKnownQusbComPort = foundComPort;
                lastKnownQusbLabel = foundLabel;

                if (needsUpdate)
                {
                    portSubmenu.GetCurrentParent()?.BeginInvoke(new Action(() =>
                    {
                        // Only add if not already present
                        bool exists = portSubmenu.DropDownItems.Cast<ToolStripMenuItem>().Any(i => i.Text == foundLabel);
                        if (!exists)
                        {
                            var item = new ToolStripMenuItem(foundLabel);
                            item.Click += (s, e) =>
                            {
                                isQusb2snesAvailable = true;
                                SetPort(InputType.Qusb2Snes, foundComPort);
                                RefreshPortMenuChecks();
                            };
                            portSubmenu.DropDownItems.Insert(1, item);
                            RefreshPortMenuChecks();
                        }
                    }));
                }
            }
        }

        // 4. Always update menu checks
        private void RefreshPortMenuChecks()
        {
            foreach (ToolStripMenuItem item in portSubmenu.DropDownItems)
            {
                string label = item.Text;
                bool shouldBeChecked =
                    (label == "None" && serialPortName == "None") ||
                    (label.Contains("XInput") && serialPortName.StartsWith("XInput")) ||
                    (label.Contains("QUSB2SNES") && serialPortName == lastKnownQusbComPort) ||
                    (label.StartsWith("COM") && serialPortName == label);

                // DirectInput universal check: menu label short-guid matches serialPortName
                if (label.Contains("[") && Guid.TryParse(serialPortName, out var guid))
                {
                    var shortGuid = guid.ToString().Substring(0, 8);
                    if (label.EndsWith($"[{shortGuid}]"))
                        shouldBeChecked = true;
                }

                item.Checked = shouldBeChecked;
            }
        }

        private string GetPnpDeviceIdForPort(string portName)
        {
            // DeviceID is like "COM3"
            using var searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_SerialPort WHERE DeviceID = '{portName}'");
            foreach (var obj in searcher.Get())
            {
                var pnpId = obj["PNPDeviceID"]?.ToString();
                if (!string.IsNullOrEmpty(pnpId))
                    return pnpId;
            }
            return null;
        }

        private void AddZoomOption(string label, int factor)
        {
            var item = new ToolStripMenuItem(label)
            {
                Checked = (zoomFactor == factor),
                CheckOnClick = true
            };
            item.Click += (s, e) =>
            {
                zoomFactor = factor;
                settings.ZoomFactor = factor;
                foreach (ToolStripMenuItem m in zoomSubmenu.DropDownItems)
                    m.Checked = m == item;
                ApplyZoom();
            };
            zoomSubmenu.DropDownItems.Add(item);
        }

        private void ApplyZoom()
        {
            int newWidth = 1, newHeight = 1; // Always initialize to prevent unassigned variable error

            // --- Animated background frames ---
            if (originalBackgroundFrames != null && originalBackgroundFrames.Count > 0)
            {
                animationTimer?.Stop();

                // Build new scaled frames
                newWidth = (int)(originalBackgroundFrames[0].Width * zoomFactor);
                newHeight = (int)(originalBackgroundFrames[0].Height * zoomFactor);
                List<Bitmap> newFrames = new();
                foreach (var frame in originalBackgroundFrames)
                {
                    Bitmap scaled = new Bitmap(newWidth, newHeight);
                    using (Graphics g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode = InterpolationMode.NearestNeighbor;
                        g.PixelOffsetMode = PixelOffsetMode.Half;
                        g.DrawImage(frame, new Rectangle(0, 0, newWidth, newHeight));
                    }
                    newFrames.Add(scaled);
                }

                // Remember old frames and background
                var oldFrames = backgroundFrames;
                Bitmap oldBackground = background;

                // Assign new frames
                backgroundFrames = newFrames;
                backgroundFrameIndex = Math.Min(backgroundFrameIndex, backgroundFrames.Count - 1);

                lock (backgroundLock)
                {
                    background = backgroundFrames[backgroundFrameIndex];
                }

                // Dispose old frames, but NEVER any frame currently referenced by 'background'
                if (oldFrames != null)
                {
                    foreach (var bmp in oldFrames)
                    {
                        // Only dispose if it's not any of the new frames, and not the one assigned
                        if (bmp != null && !backgroundFrames.Contains(bmp) && bmp != background)
                            bmp.Dispose();
                    }
                }
                if (oldBackground != null && !backgroundFrames.Contains(oldBackground) && oldBackground != background)
                    oldBackground.Dispose();

                this.ClientSize = background.Size;

                // Animation timer restart moved to end
            }
            // --- Static background ---
            else if (originalBackground != null)
            {
                newWidth = (int)(originalBackground.Width * zoomFactor);
                newHeight = (int)(originalBackground.Height * zoomFactor);

                Bitmap oldBackground = background;

                lock (backgroundLock)
                {
                    background = ScaleBitmapNearest(originalBackground, newWidth, newHeight);
                }

                if (oldBackground != null && oldBackground != background)
                    oldBackground.Dispose();

                this.ClientSize = background.Size;
            }

            // --- Button image scaling ---
            foreach (var imgList in buttonImages.Values)
                foreach (var img in imgList)
                    img.Dispose();
            buttonImages.Clear();

            buttonImages = originalButtonImages.ToDictionary(
                kvp => kvp.Key,
                kvp => kvp.Value.Select(orig =>
                {
                    Bitmap bmp = new Bitmap((int)(orig.Width * zoomFactor), (int)(orig.Height * zoomFactor));
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.InterpolationMode = InterpolationMode.NearestNeighbor;
                        g.PixelOffsetMode = PixelOffsetMode.Half;
                        g.DrawImage(orig, new Rectangle(0, 0, bmp.Width, bmp.Height));
                    }
                    return (Image)bmp;
                }).ToList()
            );

            foreach (var kvp in animatedFrames)
            {
                var key = kvp.Key;
                var animLists = kvp.Value;
                var newAnimLists = new List<List<Bitmap>>();

                foreach (var frameList in animLists)
                {
                    var newFrames = frameList.Select(frame =>
                    {
                        var bmp = new Bitmap((int)(frame.Width * zoomFactor), (int)(frame.Height * zoomFactor));
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.InterpolationMode = InterpolationMode.NearestNeighbor;
                            g.PixelOffsetMode = PixelOffsetMode.Half;
                            g.SmoothingMode = SmoothingMode.None;
                            g.CompositingMode = CompositingMode.SourceOver;
                            g.DrawImage(frame, new Rectangle(0, 0, bmp.Width, bmp.Height));
                        }
                        return bmp;
                    }).ToList();
                    newAnimLists.Add(newFrames);
                }

                // Dispose old, except any in use (background, or now part of newAnimLists)
                foreach (var frameList in animLists)
                {
                    foreach (var frame in frameList)
                    {
                        if (frame != background && !newAnimLists.SelectMany(list => list).Contains(frame))
                            frame.Dispose();
                    }
                }

                animatedFrames[key] = newAnimLists;
            }

            // Rescale button layout
            foreach (var list in buttons.Values)
            {
                foreach (var button in list)
                {
                    button.X = button.OriginalX * zoomFactor;
                    button.Y = button.OriginalY * zoomFactor;
                    button.Width = button.OriginalWidth * zoomFactor;
                    button.Height = button.OriginalHeight * zoomFactor;
                }
            }

            this.Invalidate();
            CompositeAndDisplay();
            animationTimer?.Start();
        }


        private void CenterOnScreen()
        {
            var bounds = Screen.PrimaryScreen.WorkingArea;
            bitmapScreenPosition = new Point(
                bounds.Left + (bounds.Width - this.ClientSize.Width) / 2,
                bounds.Top + (bounds.Height - this.ClientSize.Height) / 2
            );
        }

        private async void SetPort(InputType type, object param)
        {
            inputTask = null;
            qusb2snesSource?.Stop();
            qusb2snesSource = null;

            if (activeDirectInputDevice != null)
            {
                try
                {
                    activeDirectInputDevice.Unacquire();
                    activeDirectInputDevice.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[EXCEPTION] " + ex.ToString());
                }
                activeDirectInputDevice = null;
            }

            if (activeSerialPort != null)
            {
                try
                {
                    if (activeSerialPort.IsOpen) activeSerialPort.Close();
                    activeSerialPort.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[EXCEPTION] " + ex.ToString());
                }
                activeSerialPort = null;
            }

            // Pick string label for display tracking
            serialPortName = type switch
            {
                InputType.XInput => $"XInput:{param}",
                InputType.DirectInput => param is Tuple<Guid, string> t ? t.Item1.ToString() : "DirectInput",
                InputType.Qusb2Snes => param as string,
                InputType.Com => param as string,
                InputType.None => "None",
                _ => "Unknown"
            };

            settings.Source = serialPortName;

            if (type == InputType.DirectInput && param is Tuple<Guid, string> tuple)
            {
                selectedDirectInputGuid = tuple.Item1;
                serialPortName = tuple.Item1.ToString();
            }
            settings.Source = serialPortName;

            // Actually apply source
            inputManager.SetSource(type, serialPortName, param);
            settings.Source = serialPortName;

            PopulatePortMenuSync();
        }

        public static string ExtractValue(string pnpId, string prefix)
        {
            if (string.IsNullOrEmpty(pnpId)) return "";
            int start = pnpId.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
            if (start == -1) return "";
            start += prefix.Length;
            int end = pnpId.IndexOf('&', start);
            return end == -1 ? pnpId.Substring(start) : pnpId.Substring(start, end - start);
        }

        private string GetCommonName(string vid, string pid)
        {
            vid = vid.ToLower();
            pid = pid.ToLower();
            if (ArduinoDevices.TryGetValue((vid, pid), out var name))
            {
                return name;
            }
            if (vid == "2341" || vid == "2a03")
            {
                return "Unknown Arduino";
            }
            return null;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SaveSettings();
            inputCancelToken?.Cancel();
            inputCancelToken?.Dispose();
            inputCancelToken = null;

            try
            {
                activeDirectInputDevice?.Unacquire();
                activeDirectInputDevice?.Dispose();
            }
            catch (Exception ex)
            {
                //Debug.WriteLine($"[Dispose] DirectInput error: {ex.Message}");
            }

            try
            {
                if (activeSerialPort?.IsOpen == true)
                    activeSerialPort.Close();
                activeSerialPort?.Dispose();
            }
            catch (Exception ex)
            {
                //Debug.WriteLine($"[Dispose] SerialPort error: {ex.Message}");
            }
            animationTimer?.Stop();
            animationTimer?.Dispose();
            animationTimer = null;
            qusb2snesSource?.Stop();
            qusb2snesSource = null;
            base.OnFormClosing(e);
        }
    }
}
