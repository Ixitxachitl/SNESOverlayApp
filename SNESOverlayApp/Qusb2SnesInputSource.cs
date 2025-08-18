using System;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Linq;

public class Qusb2SnesInputSource
{
    private const string QUSB2SNES_URI = "ws://localhost:8080";
    private ClientWebSocket socket;
    private CancellationTokenSource cts;
    private string deviceName;

    public event Action<bool[], float, float> OnInputReceived;

    public async Task StartAsync()
    {
        cts = new CancellationTokenSource();
        socket = new ClientWebSocket();
        socket.Options.KeepAliveInterval = TimeSpan.FromSeconds(30);

        await socket.ConnectAsync(new Uri(QUSB2SNES_URI), cts.Token);
        await SendOpcodeAsync("DeviceList");
        var devices = await ReceiveJsonArrayAsync();
        if (devices.Length == 0) throw new Exception("No QUSB2SNES devices found.");
        deviceName = devices[0].GetString();
        await SendOpcodeAsync("Attach", new Dictionary<string, object> { ["Operands"] = new[] { deviceName } });
        await SendOpcodeAsync("Info", new Dictionary<string, object> { ["Space"] = "SNES", ["Operands"] = new[] { deviceName } });
        await Task.Delay(300);
        _ = Task.Run(async () =>
        {
            try
            {
                await PollLoop();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"QUSB2SNES polling failed: {ex}");
            }
        });
    }

    private async Task PollLoop()
    {
        while (true)
        {
            // Read low byte from F50DA4
            await SendOpcodeAsync("GetAddress", new Dictionary<string, object>
            {
                ["Space"] = "SNES",
                ["Operands"] = new object[] { "F50DA4", "1" },
                ["Flags"] = new[] { "R" }
            });
            byte[] lo = await ReceiveExactBytesAsync(1);

            // Read high byte from F50DA2
            await SendOpcodeAsync("GetAddress", new Dictionary<string, object>
            {
                ["Space"] = "SNES",
                ["Operands"] = new object[] { "F50DA2", "1" },
                ["Flags"] = new[] { "R" }
            });
            byte[] hi = await ReceiveExactBytesAsync(1);

            if (lo.Length < 1 || hi.Length < 1)
            {
                //System.Diagnostics.Debug.WriteLine("[QUSB2SNES] Data too short, retrying...");
                await Task.Delay(120);
                continue;
            }

            ushort input = (ushort)((hi[0] << 8) | lo[0]);
            //System.Diagnostics.Debug.WriteLine($"[QUSB2SNES] Input word: {input:X4}");

            // lo = $0DA4, hi = $0DA2
            bool[] bitmask = new bool[12];
            bitmask[0] = (lo[0] & 0x80) != 0; // A 
            bitmask[1] = (lo[0] & 0x40) != 0; // X 
            bitmask[2] = (lo[0] & 0x20) != 0; // Select
            bitmask[3] = (lo[0] & 0x10) != 0; // Start
            bitmask[4] = (lo[0] & 0x08) != 0; // Up
            bitmask[5] = (lo[0] & 0x04) != 0; // Down
            bitmask[6] = (lo[0] & 0x02) != 0; // Left
            bitmask[7] = (lo[0] & 0x01) != 0; // Right
            bitmask[8] = (hi[0] & 0x80) != 0; // B   
            bitmask[9] = (hi[0] & 0x40) != 0; // Y   
            bitmask[10] = (hi[0] & 0x20) != 0; // R 
            bitmask[11] = (hi[0] & 0x10) != 0; // L 

            //System.Diagnostics.Debug.WriteLine($"[QUSB2SNES] Bitmask: {string.Join(',', bitmask.Select(b => b ? '1' : '0'))}");

            OnInputReceived?.Invoke(bitmask, 0, 0);

            await Task.Delay(30);
        }
    }

    private async Task SendOpcodeAsync(string opcode, Dictionary<string, object>? additionalFields = null)
    {
        if (socket == null || socket.State != WebSocketState.Open) return;
        var obj = new Dictionary<string, object> { ["Opcode"] = opcode };
        if (additionalFields != null)
            foreach (var kvp in additionalFields) obj[kvp.Key] = kvp.Value;
        string json = JsonSerializer.Serialize(obj);
        byte[] bytes = Encoding.UTF8.GetBytes(json);
        var buffer = new ArraySegment<byte>(bytes);
        await socket.SendAsync(buffer, WebSocketMessageType.Text, true, cts.Token);
    }

    private async Task<JsonElement[]> ReceiveJsonArrayAsync()
    {
        var buffer = new byte[2048];
        var result = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), cts.Token);
        if (result.Count == 0) return Array.Empty<JsonElement>();
        string json = Encoding.UTF8.GetString(buffer, 0, result.Count).Trim();
        var doc = JsonDocument.Parse(json);
        if (doc.RootElement.TryGetProperty("Results", out var results) && results.ValueKind == JsonValueKind.Array)
            return results.EnumerateArray().ToArray();
        return Array.Empty<JsonElement>();
    }

    private async Task<byte[]> ReceiveExactBytesAsync(int expectedByteCount)
    {
        var buffer = new byte[2048];
        using var ms = new MemoryStream();
        int received = 0;
        var timeout = Task.Delay(2000);
        while (received < expectedByteCount)
        {
            var receiveTask = socket.ReceiveAsync(new ArraySegment<byte>(buffer), cts.Token);
            var completed = await Task.WhenAny(receiveTask, timeout);
            if (completed == timeout) break;
            var result = receiveTask.Result;
            if (result.Count == 0) break;
            ms.Write(buffer, 0, result.Count);
            received += result.Count;
        }
        var fullData = ms.ToArray();
        return fullData.Length != expectedByteCount ? new byte[0] : fullData;
    }

    public void Stop()
    {
        try
        {
            cts?.Cancel();
            socket?.Abort();
            socket?.Dispose();
            cts?.Dispose();
        }
        catch { }
    }
}
