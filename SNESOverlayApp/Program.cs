// Updated Program.cs to remove port selection on startup
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;

namespace SNESOverlayApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var mainForm = new OverlayForm("None"); // default to None, user picks port from menu

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using Stream iconStream = assembly.GetManifestResourceStream("SNESOverlayApp.Resources.icon.ico");
                if (iconStream != null)
                {
                    mainForm.Icon = new Icon(iconStream);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Icon] Failed to load embedded icon: {ex.Message}");
            }

            Application.Run(mainForm);
        }
    }
}