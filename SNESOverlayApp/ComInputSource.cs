using System;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Threading;
using System.Threading.Tasks;

public class ComInputSource
{
    private readonly string portName;
    private CancellationTokenSource cts;
    public event Action<bool[], float, float> OnInputReceived;

    public ComInputSource(string portName)
    {
        this.portName = portName;
    }

    public void Start()
    {
        cts = new CancellationTokenSource();
        var token = cts.Token;

        _ = Task.Run(() =>
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    int baudRate = DetectCp210x(portName) ? 115200 : 9600;

                    using var port = new SerialPort(portName, baudRate)
                    {
                        DtrEnable = true,
                        RtsEnable = true
                    };

                    port.Open();
                    port.DiscardInBuffer();
                    Thread.Sleep(250);

                    System.Diagnostics.Debug.WriteLine($"Opened {portName} at {baudRate} baud");

                    while (!token.IsCancellationRequested)
                    {
                        var buffer = new List<byte>();
                        while (true)
                        {
                            if (token.IsCancellationRequested) return;
                            int b = port.ReadByte();
                            if (b == -1) continue;
                            if (b == '\n') break;
                            buffer.Add((byte)b);
                        }

                        if (buffer.Count == 16)
                        {
                            bool[] bitmask = buffer.Select(b => b == (byte)'1').ToArray();
                            OnInputReceived?.Invoke(bitmask, 0, 0);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[Serial Error] {ex.Message}");
                    Thread.Sleep(2000);
                }
            }
        }, token);
    }

    private bool DetectCp210x(string comPort)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%({comPort})%'");

            foreach (var device in searcher.Get())
            {
                string pnpId = device["PNPDeviceID"]?.ToString();
                if (!string.IsNullOrEmpty(pnpId) && pnpId.Contains("VID_10C4", StringComparison.OrdinalIgnoreCase))
                {
                    System.Diagnostics.Debug.WriteLine($"Detected CP210x on {comPort} via PNP ID: {pnpId}");
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[WMI Error] {ex.Message}");
        }

        return false;
    }

    public void Stop()
    {
        cts?.Cancel();
        cts?.Dispose();
    }
}
