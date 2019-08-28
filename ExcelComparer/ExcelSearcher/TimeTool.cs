using System;

namespace ExcelSearcher
{
    public class TimeTool
    {
        static DateTime start = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
        public static DateTime UnixTimeToDateTime(uint sec)
        {
            return start.AddSeconds(sec);
        }

        public static uint DateTimeToUnixTime(DateTime time)
        {
            return (uint)(time - start).TotalSeconds;
        }
    }
}