using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnoSharp
{
    public static class Utils
    {
        private static int GetDaysFromOrigin(DateTime date)
        {
            int y = date.Year;
            int m = date.Month;
            int d = date.Day;

            if (m <= 2)
            {
                --y;
                m += 12;
            }
            int dy = 365 * (y - 1); // 経過年数×365日
            int c = y / 100;
            int dl = (y >> 2) - c + (c >> 2); // うるう年分
            int dm = (m * 979 - 1033) >> 5; // 1月1日から m 月1日までの日数
            return dy + dl + dm + d - 1;
        }

        public static double ConvertDateToValue(DateTime nullDateTime, DateTime value)
        {
            int nullOffset = GetDaysFromOrigin(nullDateTime);
            int valOffset = GetDaysFromOrigin(value);

            return
                (valOffset - nullOffset)
                + (value.Hour / 24d)
                + (value.Minute / 24d / 60d)
                + (value.Second / 24d / 60d / 60d);

        }

        public static DateTime ConvertValueToDate(DateTime nullDateTime, double value)
        {
            var rtnDate = new DateTime(nullDateTime.Year, nullDateTime.Month, nullDateTime.Day);
            rtnDate = rtnDate.AddDays((int)value);

            double time = value % 1;
            rtnDate = rtnDate.AddHours((int)(time * 24) % 24);
            rtnDate = rtnDate.AddMinutes((int)(time * 24 * 60) % 60);
            rtnDate = rtnDate.AddSeconds((int)Math.Round(time * 24 * 60 * 60) % 60);

            return rtnDate;
        }

        public static string CovertIndexToColumnLabel(int col)
        {
            StringBuilder colRef = new StringBuilder(4);

            col = col + 1;

            while (col > 0)
            {
                int thisPart = col % 26;
                if (thisPart == 0) { thisPart = 26; }

                col = (col - thisPart) / 26;

                colRef.Insert(0, (char)(thisPart + 64));
            }

            return colRef.ToString();
        }

        public static int ConvertColumnLabelToIndex(string lbl)
        {
            int retval = 0;
            foreach (var thechar in lbl.ToUpper())
            {
                retval = (retval * 26) + (thechar - 'A' + 1);

            }

            return retval - 1;
        }
    }
}
