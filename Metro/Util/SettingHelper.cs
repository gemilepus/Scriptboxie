using IniParser;
using IniParser.Model;

namespace Metro
{
    public class SettingHelper
    {
        public string TestMode, Language,
            OnOff_Hotkey,Run_Hotkey,Stop_Hotkey,
            TypeOfKeyboardInput;

        public bool ShowBalloon = false, Topmost, HideOnSatrt,
            OnOff_AltKey, Run_AltKey, Stop_AltKey,
            OnOff_CrtlKey, Run_CrtlKey, Stop_CrtlKey;

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
            if (data["Def"]["OnOff_AltKey"] == null)
            {
                data["Def"]["OnOff_AltKey"] = data["Def"]["OnOff_CrtlKey"] == null ? "1" : "0";
            }
            if (data["Def"]["OnOff_CrtlKey"] == null)
            {
                data["Def"]["OnOff_CrtlKey"] = "0";
            }
            if (data["Def"]["OnOff_Hotkey"] == null)
            {
                data["Def"]["OnOff_Hotkey"] = "Oem7";// '
            }

            //  Run Hotkey
            if (data["Def"]["Run_AltKey"] == null)
            {
                data["Def"]["Run_AltKey"] = data["Def"]["Run_CrtlKey"] == null ? "1" : "0";
            }
            if (data["Def"]["Run_CrtlKey"] == null)
            {
                data["Def"]["Run_CrtlKey"] = "0";
            }
            if (data["Def"]["Run_Hotkey"] == null)
            {
                data["Def"]["Run_Hotkey"] = "OemOpenBrackets"; // [
            }

            //  Stop Hotkey
            if (data["Def"]["Stop_AltKey"] == null)
            {
                data["Def"]["Stop_AltKey"] = data["Def"]["Stop_CrtlKey"] == null ? "1" : "0";
            }
            if (data["Def"]["Stop_CrtlKey"] == null)
            {
                data["Def"]["Stop_CrtlKey"] = "0";
            }
            if (data["Def"]["Stop_Hotkey"] == null)
            {
                data["Def"]["Stop_Hotkey"] = "Oem6"; // ]
            }

            if (data["Def"]["TypeOfKeyboardInput"] == null)
            {
                data["Def"]["TypeOfKeyboardInput"] = "Game";
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

            OnOff_AltKey  = data["Def"]["OnOff_AltKey"].Equals("1") ? true : false;
            OnOff_CrtlKey = data["Def"]["OnOff_CrtlKey"].Equals("1") ? true : false;
            OnOff_Hotkey  = ConvertHelper.ConvertKeyCode(data["Def"]["OnOff_Hotkey"]);

            Run_AltKey  = data["Def"]["Run_AltKey"].Equals("1") ? true : false;
            Run_CrtlKey = data["Def"]["Run_CrtlKey"].Equals("1") ? true : false;
            Run_Hotkey  = ConvertHelper.ConvertKeyCode(data["Def"]["Run_Hotkey"]);

            Stop_AltKey  = data["Def"]["Stop_AltKey"].Equals("1") ? true : false;
            Stop_CrtlKey = data["Def"]["Stop_CrtlKey"].Equals("1") ? true : false;
            Stop_Hotkey  = ConvertHelper.ConvertKeyCode(data["Def"]["Stop_Hotkey"]);

            TypeOfKeyboardInput = data["Def"]["TypeOfKeyboardInput"];
            ShowBalloon = data["Def"]["ShowBalloon"].Equals("1") ? true : false;
            HideOnSatrt = data["Def"]["HideOnSatrt"].Equals("1") ? true : false;
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

            data["Def"]["OnOff_AltKey"]  = OnOff_AltKey ? "1" : "0";
            data["Def"]["OnOff_CrtlKey"] = OnOff_CrtlKey ? "1" : "0";
            data["Def"]["OnOff_Hotkey"]  =  OnOff_Hotkey;

            data["Def"]["Run_AltKey"]  = Run_AltKey ? "1" : "0";
            data["Def"]["Run_CrtlKey"] = Run_CrtlKey ? "1" : "0";
            data["Def"]["Run_Hotkey"]  =  Run_Hotkey;

            data["Def"]["Stop_AltKey"]  = Stop_AltKey ? "1" : "0";
            data["Def"]["Stop_CrtlKey"] = Stop_CrtlKey ? "1" : "0";
            data["Def"]["Stop_Hotkey"]  =  Stop_Hotkey;

            data["Def"]["TypeOfKeyboardInput"] = TypeOfKeyboardInput;
            data["Def"]["HideOnSatrt"] = HideOnSatrt ? "1" : "0";
            data["Def"]["TestMode"] = TestMode;
            data["Def"]["Topmost"] = Topmost ? "1" : "0";
            data["Def"]["Language"] = Language;

            parser.WriteFile("user.ini", data);
        }
    }
}
