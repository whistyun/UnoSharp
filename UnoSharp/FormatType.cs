using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.util;

namespace UnoSharp
{
    public enum FormatType
    {
        Unknown,
        DateTime,
        Date,
        Time,
        Number,
        Boolean,
        Text
    }
    public static class FormatTypeExt
    {
        private static bool hasFlag(short input, short checkFlg)
        {
            return (input & checkFlg) == checkFlg;
        }
        private static bool hasFlag(short input, params short[] checkFlgs)
        {
            foreach (var flg in checkFlgs)
                if (hasFlag(input, flg)) return true;

            return false;
        }

        public static FormatType ConvertFromNumberFormat(short numberFormatFlags)
        {
            if (hasFlag(numberFormatFlags, NumberFormat.DATETIME))
                return FormatType.DateTime;

            if (hasFlag(numberFormatFlags, NumberFormat.DATE))
                return FormatType.Date;

            if (hasFlag(numberFormatFlags, NumberFormat.DURATION, NumberFormat.TIME))
                return FormatType.Time;

            if (hasFlag(numberFormatFlags, NumberFormat.CURRENCY, NumberFormat.FRACTION, NumberFormat.NUMBER, NumberFormat.PERCENT, NumberFormat.SCIENTIFIC))
                return FormatType.Number;

            if (hasFlag(numberFormatFlags, NumberFormat.LOGICAL))
                return FormatType.Boolean;

            if (hasFlag(numberFormatFlags, NumberFormat.TEXT))
                return FormatType.Text;

            return FormatType.Unknown;
        }

        public static short ConvertToNumberFormat(this FormatType type)
        {
            switch (type)
            {
                case FormatType.Unknown:
                    return NumberFormat.ALL;

                case FormatType.DateTime:
                    return NumberFormat.DATETIME;

                case FormatType.Date:
                    return NumberFormat.DATE;

                case FormatType.Time:
                    return NumberFormat.TIME;

                case FormatType.Number:
                    return NumberFormat.NUMBER;

                case FormatType.Boolean:
                    return NumberFormat.LOGICAL;

                case FormatType.Text:
                    return NumberFormat.TEXT;

            }
            throw new InvalidOperationException();
        }
    }
}
