using System;
using System.Linq;
using System.Text;
using uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.table;
using Locale = unoidl.com.sun.star.lang.Locale;
using XNumberFormatTypes = unoidl.com.sun.star.util.XNumberFormatTypes;

namespace UnoSharp
{
    public class Range
    {
        public Workbook Workbook { get => Worksheet.Workbook; }

        public Worksheet Worksheet { get; }

        private XCellRange Peer { get; }

        public int Row0 { get; }

        public int Column0 { get; }

        public int RowCount { get; }

        public int ColumnCount { get; }

        public Range(Worksheet wksheet, int row01, int col01, int row02, int col02)
        {
            this.Worksheet = wksheet;

            this.Row0 = row01;
            this.Column0 = col01;

            this.RowCount = row02 - row01 + 1;
            this.ColumnCount = col02 - col01 + 1;

            this.Peer = wksheet.Peer.getCellRangeByPosition(col01, row01, col02, row02);
        }

        public Cell this[int r, int c]
        {
            get { return CellAt(r, c); }
        }

        public virtual Cell CellAt(int row0, int col0)
        {
            if (row0 < 0 || col0 < 0 || row0 >= RowCount || col0 >= ColumnCount)
                throw new IndexOutOfRangeException();

            return new Cell(Worksheet, Row0 + row0, Column0 + col0);
        }

        public virtual Cell Offset(int row0, int col0)
        {
            return new Cell(Worksheet, Row0 + row0, Column0 + col0);
        }
        public short FormatTypeBit
        {
            set
            {
                // https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Applying_Number_Formats

                var nft = (XNumberFormatTypes)Workbook.FormatsSupplier.getNumberFormats();
                var fmt = nft.getStandardFormat(value, new Locale());
                ((XPropertySet)Peer).setPropertyValue("NumberFormat", new Any(fmt));
            }
        }

        public FormatType FormatType
        {
            set
            {
                FormatTypeBit = value.ConvertToNumberFormat();
            }
        }


        public virtual object[][] Values
        {
            set => SetValue(value, false);
            get => GetValue(false);
        }

        internal void SetValue(object[][] value, bool ignoreFormat)
        {
            var nullDate = Workbook.NullDate;

            // check a count of rows
            if (value.Length != RowCount)
                throw new IndexOutOfRangeException("The count of rows");


            var anys = new Any[RowCount][];

            for (var r = 0; r < RowCount; ++r)
            {
                var line = value[r];

                // check a count of rows and columns
                if (line.Length != ColumnCount)
                    throw new IndexOutOfRangeException("The count of columns");

                var anyLine = new Any[ColumnCount];
                anys[r] = anyLine;

                for (var c = 0; c < ColumnCount; ++c)
                {
                    var elmnt = line[c];
                    // null
                    if (elmnt == null)
                    {
                        anyLine[c] = Any.VOID;
                    }
                    // Boolean
                    else if (elmnt is bool)
                    {
                        // Workaround: when use Any(bool), "setDataArray" throw RuntimeException.
                        if (!ignoreFormat) CellAt(Row0 + r, Column0 + c).FormatType = FormatType.Boolean;
                        anyLine[c] = new Any((bool)elmnt ? 1 : 0);
                    }
                    // Long
                    else if (elmnt is long)
                    {
                        // Workaround: when use Any(long), "setDataArray" throw RuntimeException.
                        var lngVal = (long)elmnt;

                        if (lngVal == (int)lngVal)
                            anyLine[c] = new Any((int)lngVal);

                        else
                            anyLine[c] = new Any(lngVal.ToString());
                    }
                    // Number
                    else if (elmnt is short | elmnt is int | elmnt is float | elmnt is double)
                    {
                        anyLine[c] = new Any(elmnt.GetType(), elmnt);
                    }
                    // Text
                    else if (elmnt is String | elmnt is StringBuilder)
                    {
                        anyLine[c] = new Any(elmnt.ToString());
                    }
                    // DateTime
                    else if (elmnt is DateTime)
                    {
                        if (!ignoreFormat) CellAt(Row0 + r, Column0 + c).FormatType = FormatType.DateTime;
                        anyLine[c] = new Any(Utils.ConvertDateToValue(nullDate, (DateTime)elmnt));
                    }
                    // Unsupport
                    else
                    {
                        throw new ArgumentException(
                            String.Format("<{0},{1}> '{2}' is not support", r, c, elmnt.GetType().Name));
                    }
                }
            }

            var dPeer = (XCellRangeData)Peer;
            dPeer.setDataArray(anys);
        }

        internal object[][] GetValue(bool ignoreFormat)
        {
            var nullDate = Workbook.NullDate;

            var dPeer = (XCellRangeData)Peer;
            var anys = dPeer.getDataArray();
            var vals = anys
                .Select(line => line.Select(v => v.hasValue() ? v.Value : null).ToArray())
                .ToArray();

            if (!ignoreFormat)
            {
                // check data type for Date and Boolean.
                for (var rOfst = 0; rOfst < vals.Length; ++rOfst)
                {
                    var line = vals[rOfst];

                    for (var cOfst = 0; cOfst < line.Length; ++cOfst)
                    {
                        var elmnt = line[cOfst];

                        if (elmnt is double)
                        {
                            // check num fmt
                            switch (CellAt(Row0 + rOfst, Column0 + cOfst).FormatType)
                            {
                                case FormatType.Boolean:
                                    line[cOfst] = ((double)elmnt) != 0;
                                    break;

                                case FormatType.Date:
                                case FormatType.DateTime:
                                case FormatType.Time:
                                    line[cOfst] = Utils.ConvertValueToDate(nullDate, (double)elmnt);
                                    break;
                            }
                        }
                    }
                }
            }
            return vals;
        }
    }
}
