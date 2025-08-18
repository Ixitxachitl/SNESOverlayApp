using System;
using System.Windows.Forms;

public enum InputType
{
    None,
    Com,
    XInput,
    DirectInput,
    Qusb2Snes
}

public class UnifiedInputManager
{
    private object currentSource;
    private InputType currentType = InputType.None;
    private bool leftStickMapsToDpad;
    private bool triggersMapToBumpers;
    public event Action<bool[], float, float> OnInputReceived;

    public UnifiedInputManager(bool leftStickMapsToDpad = false, bool triggersMapToBumpers = false)
    {
        this.leftStickMapsToDpad = leftStickMapsToDpad;
        this.triggersMapToBumpers = triggersMapToBumpers;
    }

    public void SetSource(InputType type, string label, object extra = null)
    {
        if (type == currentType && IsSameSource(label, extra))
            return;

        StopCurrent();

        switch (type)
        {
            case InputType.Com:
                var com = new ComInputSource(label);
                com.OnInputReceived += HandleInput;
                com.Start();
                currentSource = com;
                break;

            case InputType.XInput:
                var xinput = new XInputInputSource((int)extra, leftStickMapsToDpad, triggersMapToBumpers);
                xinput.OnInputReceived += HandleInput;
                xinput.Start();
                currentSource = xinput;
                break;

            case InputType.DirectInput:
                if (extra is not Tuple<Guid, string> guidAndPnp)
                    throw new ArgumentException("DirectInput requires Tuple<Guid, string> as extra");
                var direct = new DirectInputInputSource(guidAndPnp.Item1, guidAndPnp.Item2, leftStickMapsToDpad, triggersMapToBumpers);
                direct.OnInputReceived += HandleInput;
                direct.Start();
                currentSource = direct;
                break;

            case InputType.Qusb2Snes:
                var qusb = new Qusb2SnesInputSource();
                qusb.OnInputReceived += HandleInput;
                qusb.StartAsync().ContinueWith(t =>
                {
                    if (t.IsFaulted)
                        Console.WriteLine($"QUSB2SNES failed to start: {t.Exception?.InnerException?.Message}");
                });
                currentSource = qusb;
                break;

            case InputType.None:
                currentSource = null;
                break;
        }

        currentType = type;
    }

    public void StopCurrent()
    {
        switch (currentSource)
        {
            case ComInputSource com:
                com.OnInputReceived -= HandleInput;
                com.Stop();
                break;
            case XInputInputSource xinput:
                xinput.OnInputReceived -= HandleInput;
                xinput.Stop();
                break;
            case DirectInputInputSource direct:
                direct.OnInputReceived -= HandleInput;
                direct.Stop();
                break;
            case Qusb2SnesInputSource qusb:
                qusb.OnInputReceived -= HandleInput;
                qusb.Stop();
                break;
        }

        currentSource = null;
        currentType = InputType.None;
    }

    private void HandleInput(bool[] bitmask, float normX, float normY)
    {
        OnInputReceived?.Invoke(bitmask, normX, normY);
    }

    private bool IsSameSource(string label, object extra)
    {
        // Optional optimization: define logic to avoid redundant restarts
        return false;
    }
}
