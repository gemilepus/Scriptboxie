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
        // Check Key
        if (mSortedList.IndexOfKey(Event) != -1)
        {
            EventValue = mSortedList.GetByIndex(mSortedList.IndexOfKey(Event)).ToString().Split(',');
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
}