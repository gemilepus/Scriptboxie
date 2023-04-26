using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace Metro
{
    public partial class App : Application
    {
        private Mutex _mutex;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        public App()
        {
            // Try to grab mutex
            bool createdNew;
            _mutex = new Mutex(true, "Scriptboxie", out createdNew);

            if (!createdNew)
            {
                MessageBox.Show("Scriptboxie already started");
                Application.Current.Shutdown();
            }
            else
            {
                // Add Event handler to exit event.
                Exit += CloseMutexHandler;
            }
        }

        protected virtual void CloseMutexHandler(object sender, EventArgs e)
        {
            _mutex?.Close();
        }
    }

}
