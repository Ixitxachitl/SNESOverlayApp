using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;
using static SNESOverlayApp.OverlayForm;

namespace SNESOverlayApp
{
    public class DisplayManager : IDisposable
    {
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

        public DisplayManager(Action<Bitmap> setBitmapCallback)
        {
            this.setBitmapCallback = setBitmapCallback;
        }

        public void Dispose()
        {
            animationTimer?.Stop();
            animationTimer?.Dispose();
            animationTimer = null;

            // Dispose bitmaps
            background?.Dispose();
            originalBackground?.Dispose();
            if (backgroundFrames != null) foreach (var bmp in backgroundFrames) bmp?.Dispose();
            if (originalBackgroundFrames != null) foreach (var bmp in originalBackgroundFrames) bmp?.Dispose();
            // TODO: Dispose all button images/animated frames as needed
        }

        public void SetBitmap(Bitmap bitmap)
        {
            setBitmapCallback?.Invoke(bitmap);
        }

        public void CompositeAndDisplay(
    Dictionary<string, List<ButtonInfo>> buttons,
    HashSet<string> activeButtons)
        {
            if (background == null)
                return;

            try { var test = background.Width; }
            catch { return; }

            lock (backgroundLock)
            {
                if (background == null) return;

                Bitmap composite = new Bitmap(background.Width, background.Height, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(composite))
                {
                    g.InterpolationMode = InterpolationMode.NearestNeighbor;
                    g.PixelOffsetMode = PixelOffsetMode.Half;
                    g.SmoothingMode = SmoothingMode.None;
                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.DrawImage(background, 0, 0);

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
                                            elapsed = maxDuration - 1;
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


    }
}