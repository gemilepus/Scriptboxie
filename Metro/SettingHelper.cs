using IniParser;
using IniParser.Model;

namespace Metro
{
    public class SettingHelper
    {

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

            parser.WriteFile("user.ini", data);

        }

        public void End(int x)
        {


        }
    }
}
