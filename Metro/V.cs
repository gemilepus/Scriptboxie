﻿using System;
using System.Collections;

public class V
{
    public string[] Get_Split(string CommandData)
    {
        string[] Split;
        Split = CommandData.Split(',');

        return Split;
    }

    public string[] Get_EventValue(SortedList mSortedList, string Event)
    {
        string[] EventValue = null;
        bool NOT = Event.StartsWith("(") && Event.EndsWith(")");
        if (NOT) {
            Event = Event.Replace("(","").Replace(")","");
        }

        // Check Key
        if (mSortedList.IndexOfKey(Event) != -1)
        {
            if (!NOT)
            {
                EventValue = mSortedList.GetByIndex(mSortedList.IndexOfKey(Event)).ToString().Split(',');
            }
        }
        else {
            if (NOT)
            {
                EventValue = new string[] { "" };
            }
        }

        return EventValue;
    }

    public int Get_ValueX(int x, float ScaleX, int OffsetX)
    {
        x = (int)(x * ScaleX) + OffsetX;

        return x;
    }

    public int Get_ValueY(int y, float ScaleY, int OffsetY)
    {
        y = (int)(y * ScaleY) + OffsetY;

        return y;
    }
    public Boolean NOT_Check(string Event, Boolean mBoolean)
    {
        if (Event.Substring(0, 1).Equals("!"))
        {
            mBoolean = !mBoolean;
        }

        return mBoolean;
    }
}