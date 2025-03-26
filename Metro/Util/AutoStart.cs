using Microsoft.Win32;
using System;
using System.Diagnostics;

namespace Metro
{
    public static class AutoStart
    {
        const string subkey = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
        const string key = "Scriptboxie";

        public static void SetAutoStart(bool enable)
        {
            string val = '"' + System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + '\\' + key + ".exe" + '"';

            if (enable)
            {
                try
                {
                    RegistryKey RegK = Registry.CurrentUser.OpenSubKey(subkey, true);
                    if (RegK == null)
                    {
                        RegK = Registry.CurrentUser.CreateSubKey(subkey, true);
                    }
                    RegK.SetValue(key, val);
                    RegK.Close();
                }
                catch
                {
                    Console.WriteLine("Failed to set registry");
                }
            }
            else {
                try
                {
                    RegistryKey RegK = Registry.CurrentUser.OpenSubKey(subkey, true);
                    if (RegK != null)
                    {
                        RegK.DeleteValue(key, false);
                        RegK.Close();
                    }
                }
                catch
                {
                    Console.WriteLine("Failed to delete registry");
                }
            }

        }
    }
}
