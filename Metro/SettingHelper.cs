using IniParser;
using IniParser.Model;
using System.Windows;

namespace Metro
{
    public class SettingHelper
    {
        public string OnOff_Hotkey, OnOff_CrtlKey, Run_Hotkey, Run_CrtlKey , Stop_Hotkey ,Stop_CrtlKey,
            TestMode , HideOnSatrt , Language;
        public bool ShowBalloon = false, Topmost;

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
                data["Def"]["ShowBalloon"] = "0";
            }

            //  ON_OFF Hotkey
            if (data["Def"]["OnOff_CrtlKey"] == null)
            {
                data["Def"]["OnOff_CrtlKey"] = "1";
            }
            if (data["Def"]["OnOff_Hotkey"] == null)
            {
                data["Def"]["OnOff_Hotkey"] = "Oem7";// '
            }

            //  Run Hotkey
            if (data["Def"]["Run_CrtlKey"] == null)
            {
                data["Def"]["Run_CrtlKey"] = "1";
            }
            if (data["Def"]["Run_Hotkey"] == null)
            {
                data["Def"]["Run_Hotkey"] = "OemOpenBrackets"; // [
            }

            //  Stop Hotkey
            if (data["Def"]["Stop_CrtlKey"] == null)
            {
                data["Def"]["Stop_CrtlKey"] = "1";
            }
            if (data["Def"]["Stop_Hotkey"] == null)
            {
                data["Def"]["Stop_Hotkey"] = "Oem6"; // ]
            }

            // HideOnSatrt
            if (data["Def"]["HideOnSatrt"] == null)
            {
                data["Def"]["HideOnSatrt"] = "0";
            }

            // TestMode
            if (data["Def"]["TestMode"] == null)
            {
                data["Def"]["TestMode"] = "0";
            }

            // Topmost
            if (data["Def"]["Topmost"] == null)
            {
                data["Def"]["Topmost"] = "0";
            }

            // Language
            if (data["Def"]["Language"] == null)
            {
                data["Def"]["Language"] = "en";
            }
            
            parser.WriteFile("user.ini", data);

            if (data["Def"]["ShowBalloon"].ToString().Equals("1")) {
                ShowBalloon = true;
            }

            OnOff_CrtlKey = data["Def"]["OnOff_CrtlKey"];
            OnOff_Hotkey = ConvertHelper.ConvertKeyCode(data["Def"]["OnOff_Hotkey"]);

            Run_CrtlKey = data["Def"]["Run_CrtlKey"];
            Run_Hotkey = ConvertHelper.ConvertKeyCode(data["Def"]["Run_Hotkey"]);

            Stop_CrtlKey = data["Def"]["Stop_CrtlKey"];
            Stop_Hotkey = ConvertHelper.ConvertKeyCode(data["Def"]["Stop_Hotkey"]);

            HideOnSatrt = data["Def"]["HideOnSatrt"];
            TestMode = data["Def"]["TestMode"];
            Topmost = data["Def"]["Topmost"].Equals("1") ? true : false;
            Language = data["Def"]["Language"];
        }

        public void End(MainWindow MainWindow)
        {
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");

            data["Def"]["x"] = MainWindow.Left.ToString();
            data["Def"]["y"] = MainWindow.Top.ToString();

            data["Def"]["OnOff_CrtlKey"] = OnOff_CrtlKey;
            data["Def"]["OnOff_Hotkey"] =  OnOff_Hotkey;
            data["Def"]["Run_CrtlKey"] = Run_CrtlKey;
            data["Def"]["Run_Hotkey"] =  Run_Hotkey;
            data["Def"]["Stop_CrtlKey"] = Stop_CrtlKey;
            data["Def"]["Stop_Hotkey"] =  Stop_Hotkey;

            data["Def"]["HideOnSatrt"] = HideOnSatrt;
            data["Def"]["TestMode"] = TestMode;
            data["Def"]["Topmost"] = Topmost ? "1" : "0";
            data["Def"]["Language"] = Language;

            parser.WriteFile("user.ini", data);
        }
    }
}
