using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using WindowsInput.Native;
using System.Collections;
using System.Windows.Forms;

public class V
{
    public static string[] Get_Split(string CommandData)
    {

        string[] Split;
        Split = CommandData.Split(',');

        return Split;
    }

    public static string[] Get_EventValue(SortedList mSortedList, string Event)
    {

        string[] EventValue = null;
        // Check Key
        if (mSortedList.IndexOfKey(Event) != -1)
        {
            EventValue = mSortedList.GetByIndex(mSortedList.IndexOfKey(Event)).ToString().Split(',');
        }

        return EventValue;
    }
}