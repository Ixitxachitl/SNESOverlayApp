using SharpDX.DirectInput;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

public class DirectInputInputSource
{
    private readonly Guid instanceGuid;
    private readonly string pnpId;
    private readonly bool mapLeftStickToDpad;
    private readonly bool mapTriggersToBumpers;

    private Joystick device;
    private CancellationTokenSource cts;
    private Task pollingTask;

    public event Action<bool[], float, float> OnInputReceived;

    public DirectInputInputSource(Guid instanceGuid, string pnpId, bool mapLeftStickToDpad = false, bool mapTriggersToBumpers = false)
    {
        this.instanceGuid = instanceGuid;
        this.pnpId = pnpId;
        this.mapLeftStickToDpad = mapLeftStickToDpad;
        this.mapTriggersToBumpers = mapTriggersToBumpers;
    }

    private static bool IsIBuffalo(string pnpId, Guid productGuid)
    {
        string vid = ExtractValue(pnpId, "VID_");
        string pid = ExtractValue(pnpId, "PID_");

        if (!string.IsNullOrEmpty(vid) && vid.Equals("0583", StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrEmpty(pid) && pid.Equals("2060", StringComparison.OrdinalIgnoreCase))
            return true;

        return productGuid.ToString().Equals("20600583-0000-0000-0000-504944564944", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsXInputXbox(string pnpId, Guid productGuid)
    {
        string vid = ExtractValue(pnpId, "VID_");
        string pid = ExtractValue(pnpId, "PID_");

        // Trust PNP when available
        if (!string.IsNullOrEmpty(vid) && vid.Equals("045E", StringComparison.OrdinalIgnoreCase))
            return true;

        // Fallback: known Xbox 360 product GUID
        return productGuid.ToString().Equals("028e045e-0000-0000-0000-504944564944", StringComparison.OrdinalIgnoreCase);
    }

    private static bool Is8BitDo(string pnpId, Guid productGuid)
    {
        string vid = ExtractValue(pnpId, "VID_");
        string pid = ExtractValue(pnpId, "PID_");

        // Known 8BitDo Vendor ID
        if (!string.IsNullOrEmpty(vid) && vid.Equals("2DC8", StringComparison.OrdinalIgnoreCase))
            return true;

        // Fallback: match against known 8BitDo Product GUIDs (add more as needed)
        string[] knownGuids =
        {
            "61022dc8-0000-0000-0000-504944564944",  //8BitDo SN30 Pro + Wireless
            "60022dc8-0000-0000-0000-504944564944",  //8BitDo SN30 Pro + Wired
            "301d2dc8-0000-0000-0000-504944564944"   //8BitDo Ultimate 2C Wired
        };

        return knownGuids.Contains(productGuid.ToString().ToLowerInvariant());
    }

    private static void LogDeviceInfoToFile(Joystick device, string pnpId, bool isIBuffalo, bool isXbox, bool is8bitdo)
    {
        try
        {
            var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DirectInputLog.txt");

            using var writer = new StreamWriter(logPath, append: true);
            writer.WriteLine("=== DirectInput Device Selected ===");
            writer.WriteLine($"Timestamp: {DateTime.Now}");
            writer.WriteLine($"Description: {device.Properties.ProductName}");
            writer.WriteLine($"InstanceName: {device.Properties.InstanceName}");
            writer.WriteLine($"Product GUID: {device.Information.ProductGuid}");
            writer.WriteLine($"Instance GUID: {device.Information.InstanceGuid}");
            writer.WriteLine($"PNP ID: {pnpId}");
            writer.WriteLine($"Button Count: {device.Capabilities.ButtonCount}");
            writer.WriteLine($"POVs: {device.Capabilities.PovCount}");
            writer.WriteLine($"Axes: {device.Capabilities.AxeCount}");
            writer.WriteLine($"Flags: isIBuffalo={isIBuffalo}, isXbox={isXbox}, is8bitdo={is8bitdo}");
            writer.WriteLine();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to write log: {ex}");
        }
    }

    public void Start()
    {
        var directInput = new DirectInput();
        device = new Joystick(directInput, instanceGuid);
        device.Properties.BufferSize = 128;
        device.Acquire();

        bool isIBuffalo = IsIBuffalo(pnpId, device.Information.ProductGuid);
        bool isXbox = IsXInputXbox(pnpId, device.Information.ProductGuid);
        bool is8bitdo = Is8BitDo(pnpId, device.Information.ProductGuid);
        LogDeviceInfoToFile(device, pnpId, isIBuffalo, isXbox, is8bitdo);

        cts = new CancellationTokenSource();
        var token = cts.Token;

        pollingTask = Task.Run(() =>
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    device.Poll();
                    var state = device.GetCurrentState();
                    var pressed = state.Buttons;
                    var bitmask = new bool[12];
                    /*                    for (int i = 0; i < pressed.Length; i++)
                                        {
                                            if (pressed[i])
                                                System.Diagnostics.Debug.WriteLine($"Button {i} pressed");
                                        }*/
                    if (isIBuffalo)
                    {
                        // Unify iBuffalo & RetroFlag face layout:
                        // DirectInput buttons 0..3 are Y, A, X, B (observed on these pads)
                        if (pressed.Length > 0 && pressed[0]) bitmask[8] = true;  // A
                        if (pressed.Length > 1 && pressed[1]) bitmask[0] = true;  // B
                        if (pressed.Length > 2 && pressed[2]) bitmask[9] = true;  // X
                        if (pressed.Length > 3 && pressed[3]) bitmask[1] = true;  // Y

                        // Should be the same across both
                        if (pressed.Length > 4 && pressed[4]) bitmask[10] = true; // L
                        if (pressed.Length > 5 && pressed[5]) bitmask[11] = true; // R
                        if (pressed.Length > 6 && pressed[6]) bitmask[2] = true; // Select
                        if (pressed.Length > 7 && pressed[7]) bitmask[3] = true; // Start
                    }
                    else if (isXbox)
                    {
                        if (pressed.Length > 0 && pressed[0]) bitmask[0] = true;  // A
                        if (pressed.Length > 1 && pressed[1]) bitmask[8] = true;  // B
                        if (pressed.Length > 2 && pressed[2]) bitmask[1] = true;  // X
                        if (pressed.Length > 3 && pressed[3]) bitmask[9] = true;  // Y
                        if (pressed.Length > 4 && pressed[4]) bitmask[10] = true; // LB -> L
                        if (pressed.Length > 5 && pressed[5]) bitmask[11] = true; // RB -> R
                        if (pressed.Length > 6 && pressed[6]) bitmask[2] = true;  // Back -> Select
                        if (pressed.Length > 7 && pressed[7]) bitmask[3] = true;  // Start
                    }
                    else if (is8bitdo)
                    {
                        if (pressed.Length > 0 && pressed[0]) bitmask[8] = true;  // A
                        if (pressed.Length > 1 && pressed[1]) bitmask[0] = true;  // B
                        if (pressed.Length > 3 && pressed[3]) bitmask[9] = true;  // X
                        if (pressed.Length > 4 && pressed[4]) bitmask[1] = true;  // Y
                        if (pressed.Length > 6 && pressed[6]) bitmask[10] = true; // L
                        if (pressed.Length > 7 && pressed[7]) bitmask[11] = true; // R
                        if (pressed.Length > 10 && pressed[10]) bitmask[2] = true;  // Select
                        if (pressed.Length > 11 && pressed[11]) bitmask[3] = true;  // Start
                    }
                    else
                    {
                        if (pressed.Length > 0 && pressed[0]) bitmask[1] = true; // Y
                        if (pressed.Length > 1 && pressed[1]) bitmask[0] = true; // B
                        if (pressed.Length > 2 && pressed[2]) bitmask[8] = true; // A
                        if (pressed.Length > 3 && pressed[3]) bitmask[9] = true; // X
                        if (pressed.Length > 4 && pressed[4]) bitmask[10] = true; // L
                        if (pressed.Length > 5 && pressed[5]) bitmask[11] = true; // R
                        if (pressed.Length > 8 && pressed[8]) bitmask[2] = true; // Select
                        if (pressed.Length > 9 && pressed[9]) bitmask[3] = true; // Start
                    }

                    int pov = state.PointOfViewControllers.Length > 0 ? state.PointOfViewControllers[0] : -1;
                    if (pov >= 0)
                    {
                        if (pov <= 4500 || pov >= 31500) bitmask[4] = true; // Up
                        if (pov >= 4500 && pov <= 13500) bitmask[7] = true; // Right
                        if (pov >= 13500 && pov <= 22500) bitmask[5] = true; // Down
                        if (pov >= 22500 && pov <= 31500) bitmask[6] = true; // Left
                    }

                    float normX = (state.X - 32767f) / 32767f;
                    float normY = (state.Y - 32767f) / 32767f;

                    if (mapLeftStickToDpad)
                    {
                        const float deadzone = 0.35f;
                        if (normX < -deadzone) bitmask[6] = true; // Left
                        if (normX > deadzone) bitmask[7] = true; // Right
                        if (normY < -deadzone) bitmask[4] = true; // Up
                        if (normY > deadzone) bitmask[5] = true; // Down
                    }

                    var pressed_bumper = state.Buttons;

                    // bitmask[10] = L, bitmask[11] = R (bumpers)
                    if (mapTriggersToBumpers)
                    {
                        bool isL = false, isR = false;

                        // Only synthesize bumpers from analog triggers on devices that have them.
                        // This avoids Start/Select (Buttons[6]/[7]) being treated as bumpers on iBuffalo/RetroFlag.
                        if (isXbox || is8bitdo)
                        {
                            // Many XInput-like devices expose LT/RT on Z (0..65535). Use wide thresholds.
                            int z = state.Z;
                            if (z < 20000) isL = true;     // LT
                            if (z > 45000) isR = true;     // RT

                            // If your specific device uses RotationZ instead, uncomment:
                            // int rz = state.RotationZ;
                            // if (rz < 20000) isL = true;
                            // if (rz > 45000) isR = true;
                        }

                        if (isL) bitmask[10] = true;
                        if (isR) bitmask[11] = true;
                    }

                    OnInputReceived?.Invoke(bitmask, normX, normY);
                }
                catch
                {

                }

                Thread.Sleep(16);
            }
        }, token);
    }

    public void Stop()
    {
        try
        {
            cts?.Cancel();
            pollingTask?.Wait(100);
            device?.Unacquire();
            device?.Dispose();
            cts?.Dispose();
        }
        catch { }
    }

    private static string ExtractValue(string pnpId, string prefix)
    {
        if (string.IsNullOrEmpty(pnpId)) return "";
        int start = pnpId.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
        if (start == -1) return "";
        start += prefix.Length;
        int end = pnpId.IndexOf('&', start);
        return end == -1 ? pnpId[start..] : pnpId[start..end];
    }
}
