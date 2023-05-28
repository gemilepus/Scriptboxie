using System;
using System.Runtime.InteropServices;
using WindowsInput.Native;
using System.Collections;
using System.Windows.Forms;
using System.Text.RegularExpressions;

public class ConvertHelper
{
    // VkKeyScan Char to 0x00
    [DllImport("user32.dll")]
    private static extern byte VkKeyScan(char ch);

    public int MakeLParam(int LoWord, int HiWord)
    {
        return ((HiWord << 16) | (LoWord & 0xffff));
    }

    public Keys ConvertCharToVirtualKey(char ch)
    {
        short vkey = VkKeyScan(ch);
        Keys retval = (Keys)(vkey & 0xff);
        int modifiers = vkey >> 8;
        if ((modifiers & 1) != 0) retval |= Keys.Shift;
        if ((modifiers & 2) != 0) retval |= Keys.Control;
        if ((modifiers & 4) != 0) retval |= Keys.Alt;
        return retval;
    }

    public int StringToVirtualKeyCode(String str)
    {
        str = str.ToUpper();
        int value = 0;
        Array enumValueArray = Enum.GetValues(typeof(VirtualKeyCode));
        foreach (int enumValue in enumValueArray)
        {
            if (Enum.GetName(typeof(VirtualKeyCode), enumValue).Equals(str))
            {
                value = enumValue;
            }
        }
        return value;
    }
    public string ConvertKeyString(string KeyValue)
    {
        switch (KeyValue)
        {
            case "ESC":
                KeyValue = "ESCAPE";
                break;
            case "ALT":
                KeyValue = "MENU";
                break;
            case "CTRL":
                KeyValue = "CONTROL";
                break;
            case "ENTER":
                KeyValue = "RETURN";
                break;
            case "WIN":
            case "WINDOWS":
                KeyValue = "LWIN";
                break;
            case ";":
                KeyValue = "OEM_1";
                break;
            case "=":
                KeyValue = "OEM_PLUS";
                break;
            case ",":
                KeyValue = "OEM_COMMA";
                break;
            case "-":
                KeyValue = "OEM_MINUS";
                break;
            case ".":
                KeyValue = "OEM_PERIOD";
                break;
            case "/":
                KeyValue = "OEM_2";
                break;
            case "`":
                KeyValue = "OEM_3";
                break;
            case "\\":
                KeyValue = "OEM_5";
                break;
            case "[":
                KeyValue = "OEM_4";
                break;
            case "]":
                KeyValue = "OEM_6";
                break;
            case "'":
                KeyValue = "OEM_7";
                break;

        }
        return KeyValue;
    }
    public static string ConvertKeyCode(string mKeyCode)
    {
        // Number
        if (Regex.IsMatch(mKeyCode, @"^D[0-9]{1}")) {
            return mKeyCode.Replace("D","");
        }

        switch (mKeyCode)
        {
            case "Return":
                mKeyCode = "ENTER";
                break;
            case "Escape":
                mKeyCode = "ESC";
                break;
            case "LMenu":
            case "RMenu":
                mKeyCode = "ALT";
                break;
            case "LShiftKey":
            case "RShiftKey":
                mKeyCode = "SHIFT";
                break;
            case "LControlKey":
            case "RControlKey":
                mKeyCode = "CTRL";
                break;
            case "LWin":
            case "RWin":
                mKeyCode = "WIN";
                break;
            case "Oem1":
                mKeyCode = ";";
                break;
            case "Oemplus":
                mKeyCode = "=";
                break;
            case "Oemcomma":
                mKeyCode = ",";
                break;
            case "OemMinus":
                mKeyCode = "-";
                break;
            case "OemPeriod":
                mKeyCode = ".";
                break;
            case "OemQuestion":
                mKeyCode = "/";
                break;
            case "Oemtilde":
                mKeyCode = "`";
                break;
            case "Oem5":
                mKeyCode = "\\";
                break;
            case "OemOpenBrackets":
                mKeyCode = "[";
                break;
            case "Oem6":
                mKeyCode = "]";
                break;
            case "Oem7":
                mKeyCode = "'";
                break;

            //
            case "PageUp":
                mKeyCode = "PRIOR";
                break;
            case "PrintScreen":
                mKeyCode = "SNAPSHOT";
                break; 
        }
        return mKeyCode;
    }
    public void GetEnumVirtualKeyCodeValues()
    {
        Array enumValueArray = Enum.GetValues(typeof(VirtualKeyCode));
        ArrayList myArrayList = new ArrayList();
        myArrayList.AddRange(enumValueArray);
        //string m= myArrayList[0].ToString();
        foreach (int enumValue in enumValueArray)
        {
            Console.WriteLine("Name: " + Enum.GetName(typeof(VirtualKeyCode), enumValue) + " , Value: " + enumValue);
        }
    }

    public void ParseEnum()
    {
        VirtualKeyCode ms = (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), "MENU");
        Console.WriteLine(ms.ToString());
        Array enumValueArray = Enum.GetValues(typeof(VirtualKeyCode));
    }
}