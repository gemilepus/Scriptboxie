using System;
using System.Threading;
using System.Windows.Forms;

namespace Metro
{
    public partial class App : System.Windows.Application
    {
        private Mutex _mutex;

        public App()
        {
            // Try to grab mutex
            bool createdNew;
            _mutex = new Mutex(true, "Scriptboxie", out createdNew);

            if (!createdNew)
            {
                MessageBox.Show("Scriptboxie already started", "Scriptboxie", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1,MessageBoxOptions.DefaultDesktopOnly);

                System.Windows.Application.Current.Shutdown();
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

        private void Application_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message, "Scriptboxie", MessageBoxButtons.OK,
                    MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

            System.Windows.Application.Current.Shutdown();
        }
    }

}
