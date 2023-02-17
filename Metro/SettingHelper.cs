using IniParser;
using IniParser.Model;

namespace Metro
{
    public class SettingHelper
    {
        public string OnOff_Hotkey, Run_Hotkey, Stop_Hotkey;

        public SettingHelper()
        {
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            try
            {
                data = parser.ReadFile("user.ini");
            }
            catch
            {
                parser.WriteFile("user.ini", new IniData());
            }

            // From location
            if (data["Def"]["x"] == null)
            {
                data["Def"]["x"] = "0";
                data["Def"]["y"] = "0";
            }

            //  WindowTitle setting
            if (data["Def"]["WindowTitle"] == null)
            {
                data["Def"]["WindowTitle"] = "";
            }

            //  Script
            if (data["Def"]["Script"] == null)
            {
                data["Def"]["Script"] = "";
            }

            // ScaleX, ScaleY, OffsetX, OffsetY
            if (data["Def"]["ScaleX"] == null
                && data["Def"]["ScaleY"] == null
                && data["Def"]["OffsetX"] == null
                && data["Def"]["OffsetY"] == null)
            {
                data["Def"]["ScaleX"] = "1";
                data["Def"]["ScaleY"] = "1";
                data["Def"]["OffsetX"] = "0";
                data["Def"]["OffsetY"] = "0";
              
            }

            //  ShowBalloon
            if (data["Def"]["ShowBalloon"] == null)
            {
                data["Def"]["ShowBalloon"] = "";
            }

            //  ON_OFF Hotkey
            if (data["Def"]["OnOff_Hotkey"] == null)
            {
                data["Def"]["OnOff_Hotkey"] = "Oem7";// '
            }

            //  Run Hotkey
            if (data["Def"]["Run_Hotkey"] == null)
            {
                data["Def"]["Run_Hotkey"] = "OemOpenBrackets"; // [
            }

            //  Stop Hotkey
            if (data["Def"]["Stop_Hotkey"] == null)
            {
                data["Def"]["Stop_Hotkey"] = "Oem6"; // ]
            }

            // TestMode
            if (data["Def"]["TestMode"] == null)
            {
                data["Def"]["TestMode"] = "0";
            }

            parser.WriteFile("user.ini", data);


            OnOff_Hotkey = data["Def"]["OnOff_Hotkey"];
            Run_Hotkey = data["Def"]["Run_Hotkey"];
            Stop_Hotkey = data["Def"]["Stop_Hotkey"];

        }

        public void End(MainWindow MainWindow)
        {
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");

            data["Def"]["x"] = MainWindow.Left.ToString();
            data["Def"]["y"] = MainWindow.Top.ToString();
            data["Def"]["OnOff_Hotkey"] = OnOff_Hotkey;
            data["Def"]["Run_Hotkey"] = Run_Hotkey;
            data["Def"]["Stop_Hotkey"] = Stop_Hotkey;

            parser.WriteFile("user.ini", data);
        }
    }
}
