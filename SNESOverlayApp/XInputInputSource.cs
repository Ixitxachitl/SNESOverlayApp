using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

public class XInputInputSource
{
    private CancellationTokenSource cts;
    private Task pollingTask;
    private readonly int userIndex;
    private readonly bool mapLeftStickToDpad;
    private readonly bool triggersMapToBumpers;

    public event Action<bool[], float, float> OnInputReceived;

    public XInputInputSource(int userIndex, bool mapLeftStickToDpad = false, bool triggersMapToBumpers = false)
    {
        this.userIndex = userIndex;
        this.mapLeftStickToDpad = mapLeftStickToDpad;
        this.triggersMapToBumpers = triggersMapToBumpers;
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

    [DllImport("xinput1_4.dll")]
    private static extern int XInputGetState(int dwUserIndex, out XINPUT_STATE pState);

    private static readonly Dictionary<ushort, int> ButtonMap = new()
    {
        [0x1000] = 0,  // A → SNES B
        [0x4000] = 1,  // X → SNES Y
        [0x0020] = 2,  // Back → Select
        [0x0010] = 3,  // Start
        [0x0001] = 4,  // Up
        [0x0002] = 5,  // Down
        [0x0004] = 6,  // Left
        [0x0008] = 7,  // Right
        [0x2000] = 8,  // B → SNES A
        [0x8000] = 9,  // Y → SNES X
        [0x0100] = 10, // RB → R
        [0x0200] = 11  // LB → L
    };

    public void Start()
    {
        cts = new CancellationTokenSource();
        var token = cts.Token;

        pollingTask = Task.Run(() =>
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (XInputGetState(userIndex, out var state) == 0)
                    {
                        var bitmask = new bool[16];
                        foreach (var kvp in ButtonMap)
                        {
                            if ((state.Gamepad.wButtons & kvp.Key) != 0)
                                bitmask[kvp.Value] = true;
                        }

                        float normLX = Math.Clamp(state.Gamepad.sThumbLX / 32767f, -1f, 1f);
                        float normLY = Math.Clamp(-state.Gamepad.sThumbLY / 32767f, -1f, 1f);

                        if (mapLeftStickToDpad)
                        {
                            const float deadzone = 0.35f;
                            if (!bitmask[4] && normLY < -deadzone) bitmask[4] = true;
                            if (!bitmask[5] && normLY > deadzone) bitmask[5] = true;
                            if (!bitmask[6] && normLX < -deadzone) bitmask[6] = true;
                            if (!bitmask[7] && normLX > deadzone) bitmask[7] = true;
                        }

                        if (triggersMapToBumpers)
                        {
                            if (state.Gamepad.bLeftTrigger > 30) bitmask[10] = true; // Treat LT as LB
                            if (state.Gamepad.bRightTrigger > 30) bitmask[11] = true; // Treat RT as RB
                        }

                        try
                        {
                            OnInputReceived?.Invoke(bitmask, normLX, normLY);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("[XInput] OnInputReceived error: " + ex.Message);
                        }
                    }

                    Thread.Sleep(16);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[XInput Polling Thread] Fatal: " + ex.Message);
            }
        }, token);
    }

    public void Stop()
    {
        try
        {
            cts?.Cancel();
            pollingTask?.Wait(100);
            cts?.Dispose();
        }
        catch { }
    }
}
