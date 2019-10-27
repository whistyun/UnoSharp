using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.container;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.util;

namespace UnoSharp
{
    public class Worksheet
    {
        /// <summary>"A1" pattern</summary>
        private Regex ptn1 = new Regex(@"^\$?([A-Za-z]+)\$?([0-9]+)$");
        /// <summary>"A1:A1" pattern</summary>
        private Regex ptn2 = new Regex(@"^\$?([A-Za-z]+)\$?([0-9]+):\$?([A-Za-z]+)\$?([0-9]+)$");

        public Worksheet(Workbook owner, XSpreadsheet xsheet)
        {
            this.Workbook = owner;
            this.Peer = xsheet;
        }

        public XSpreadsheet Peer { get; }

        public Workbook Workbook { get; }

        public string Name
        {
            set { ((XNamed)Peer).setName(value); }
            get { return ((XNamed)Peer).getName(); }
        }

        public Range UsedRange
        {
            get
            {
                var cursor = Peer.createCursor();

                var xuacursor = (XUsedAreaCursor)cursor;
                xuacursor.gotoStartOfUsedArea(false);
                xuacursor.gotoEndOfUsedArea(true);

                var address = ((XCellRangeAddressable)cursor).getRangeAddress();
                // 注：Excelと異なり、0-base
                // see: https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Cells_and_Ranges
                return new Range(this,
                        address.StartRow, address.StartColumn,
                        address.EndRow, address.EndColumn);
            }
        }

        public Cell LastCell
        {
            get
            {
                var cursor = Peer.createCursor();

                var xuacursor = (XUsedAreaCursor)cursor;
                xuacursor.gotoStartOfUsedArea(false);
                xuacursor.gotoEndOfUsedArea(true);

                var address = ((XCellRangeAddressable)cursor).getRangeAddress();
                // 注：Excelと異なり、0-base
                // see: https://wiki.openoffice.org/wiki/Documentation/BASIC_Guide/Cells_and_Ranges
                return CellAt(address.EndRow, address.EndColumn);
            }
        }

        public Range this[string address]
        {
            get { return Range(address); }
        }

        public Range Range(string address)
        {
            var match = ptn2.Match(address);
            if (match.Success)
            {
                return Range(
                    int.Parse(match.Groups[2].Value) - 1,
                    Utils.ConvertColumnLabelToIndex(match.Groups[1].Value),
                    int.Parse(match.Groups[4].Value) - 1,
                    Utils.ConvertColumnLabelToIndex(match.Groups[3].Value));
            }

            match = ptn1.Match(address);
            if (match.Success)
            {
                int c = Utils.ConvertColumnLabelToIndex(match.Groups[1].Value);
                int r = int.Parse(match.Groups[2].Value) - 1;

                return Range(r, c, r, c);
            }

            throw new FormatException();
        }

        public Range Range(int row01, int col01, int row02, int col02)
        {
            return new Range(this, row01, col01, row02, col02);
        }

        public Cell CellAt(string address)
        {
            var match = ptn1.Match(address);
            if (match.Success)
            {
                return CellAt(
                    int.Parse(match.Groups[2].Value) - 1,
                    Utils.ConvertColumnLabelToIndex(match.Groups[1].Value));
            }
            else throw new FormatException();

        }

        public Cell CellAt(int row0, int col0)
        {
            return new Cell(this, row0, col0);
        }

        public Cell this[int r, int c]
        {
            get { return CellAt(r, c); }
        }

        public TypedRange BuildRange(string address, bool ignoreFormat, params Type[] columnTypes)
        {
            var match = ptn1.Match(address);
            if (match.Success)
            {
                return BuildRange(
                    int.Parse(match.Groups[2].Value) - 1,
                    Utils.ConvertColumnLabelToIndex(match.Groups[1].Value),
                    ignoreFormat,
                    columnTypes);
            }
            else throw new FormatException();
        }

        public TypedRange BuildRange(string address, params Type[] columnTypes)
        {
            return BuildRange(address, false, columnTypes);
        }

        public TypedRange BuildRange(int row0, int col0, params Type[] columnTypes)
        {
            return BuildRange(row0, col0, false, columnTypes);
        }

        public TypedRange BuildRange(int row0, int col0, bool ignoreFormat, params Type[] columnTypes)
        {
            return new TypedRange(this, row0, col0, columnTypes, ignoreFormat);
        }
    }
}
