using System;
using GameOverlay.Drawing;
using GameOverlay.Windows;

namespace Metro
{
    public class Overlay : IDisposable
    {
        private GameOverlay.Windows.OverlayWindow _window;
        private GameOverlay.Drawing.Graphics _graphics;
        private GameOverlay.Drawing.SolidBrush _red;
        private GameOverlay.Drawing.Font _font;
        private GameOverlay.Drawing.SolidBrush _black;
        private GameOverlay.Drawing.Graphics gfx;

        public Overlay()
        {
            // it is important to set the window to visible (and topmost) if you want to see it!
            _window = new OverlayWindow(0, 0, System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width, System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height)
            {
                IsTopmost = true,
                IsVisible = true
            };
            // handle this event to resize your Graphics surface
            //_window.SizeChanged += _window_SizeChanged;

            // initialize a new Graphics object
            // set everything before you call _graphics.Setup()
            _graphics = new GameOverlay.Drawing.Graphics
            {
                //MeasureFPS = true,
                Height = _window.Height,
                PerPrimitiveAntiAliasing = true,
                TextAntiAliasing = true,
                UseMultiThreadedFactories = false,
                VSync = true,
                Width = _window.Width,
                WindowHandle = IntPtr.Zero
            };

            _window.Create();
            _graphics.WindowHandle = _window.Handle; // set the target handle before calling Setup()         
            _graphics.Setup();

            _red = _graphics.CreateSolidBrush(GameOverlay.Drawing.Color.Red); // those are the only pre defined Colors                                                              
            _font = _graphics.CreateFont("Arial", 25); // creates a simple font with no additional style
            _black = _graphics.CreateSolidBrush(GameOverlay.Drawing.Color.Transparent);

            gfx = _graphics; // little shortcut
        }

        public void Clear()
        {
            gfx.BeginScene(); // call before you start any drawing
            gfx.ClearScene();
            gfx.EndScene();
        }
        public void DrawRectangle(int x, int y, int w, int h)
        {
            gfx.BeginScene();
            gfx.DrawRoundedRectangle(_red, RoundedRectangle.Create(x, y, w, h , 6), 2);
            gfx.EndScene();
        }

        ~Overlay()
        {
            Dispose(false);
        }

        #region IDisposable Support
        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                _window.Dispose();

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
