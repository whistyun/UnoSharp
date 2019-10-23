using uno;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.util;

namespace UnoSharp
{
    public class Cell
    {
        public Workbook Workbook { get => Worksheet.Workbook; }

        public Worksheet Worksheet { get; }

        public XCell Peer { get; }

        public int Row0 { get; }

        public int Column0 { get; }

        public Cell(Worksheet wksheet, int row, int col)
        {
            this.Worksheet = wksheet;
            this.Row0 = row;
            this.Column0 = col;

            Peer = wksheet.Peer.getCellByPosition(Column0, Row0);
        }

        public virtual Cell Offset(int row0, int col0)
        {
            return new Cell(Worksheet, Row0 + row0, Column0 + col0);
        }

        public CellType Type
        {
            get
            {
                return (CellType)Peer.getType();
            }
        }

        public string FormatString
        {
            get
            {
                var key = (int)((XPropertySet)Peer).getPropertyValue("NumberFormat").Value;

                var propValSet = Workbook.FormatsSupplier.getNumberFormats().getByKey(key);
                return (string)propValSet.getPropertyValue("FormatString").Value;

            }
        }

        public short FormatTypeBit
        {
            set
            {
                if ((FormatTypeBit & value) != value)
                {
                    // https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Applying_Number_Formats

                    var nft = (XNumberFormatTypes)Workbook.FormatsSupplier.getNumberFormats();
                    var fmt = nft.getStandardFormat(value, new Locale());
                    ((XPropertySet)Peer).setPropertyValue("NumberFormat", new Any(fmt));
                }
            }
            get
            {
                var key = (int)((XPropertySet)Peer).getPropertyValue("NumberFormat").Value;

                var propValSet = Workbook.FormatsSupplier.getNumberFormats().getByKey(key);
                var propVal = propValSet.getPropertyValue("Type").Value;

                return (short)propVal;
            }
        }

        public FormatType FormatType
        {
            set { FormatTypeBit = value.ConvertToNumberFormat(); }
            get { return FormatTypeExt.ConvertFromNumberFormat(FormatTypeBit); }
        }

        public string Formula
        {
            set
            {
                Peer.setFormula(value);
            }
            get
            {
                return Peer.getFormula();
            }
        }

        public string Text
        {
            set
            {
                var txtPeer = (XText)Peer;
                txtPeer.setString(value);
            }
            get
            {
                return ((XText)Peer).getString();
            }
        }

        public double Value
        {
            set
            {
                Peer.setValue(value);
            }
            get
            {
                return Peer.getValue();
            }
        }

        public bool Flag
        {
            set
            {
                FormatTypeBit = NumberFormat.LOGICAL;
                Value = value ? 1 : 0;
            }
            get { return Value == 1; }
        }

        public System.DateTime Date
        {
            set
            {
                FormatTypeBit = NumberFormat.DATE;
                Value = (int)Utils.ConvertDateToValue(Workbook.NullDate, value);
            }
            get { return Utils.ConvertValueToDate(Workbook.NullDate, Value); }
        }

        public System.DateTime DateTime
        {
            set
            {
                FormatTypeBit = NumberFormat.DATETIME;
                Value = Utils.ConvertDateToValue(Workbook.NullDate, value);
            }
            get { return Utils.ConvertValueToDate(Workbook.NullDate, Value); }
        }
    }
}
