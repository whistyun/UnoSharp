using unoidl.com.sun.star.table;

namespace UnoSharp
{
    public enum CellType
    {
        Empty = CellContentType.EMPTY,
        Formula = CellContentType.FORMULA,
        Text = CellContentType.TEXT,
        Value = CellContentType.VALUE,
    }
}
