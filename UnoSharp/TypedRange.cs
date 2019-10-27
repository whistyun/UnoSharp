using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.table;

namespace UnoSharp
{
    public class TypedRange : IEnumerable<object[]>
    {
        private readonly Type[] SupportTypes = new[]{
            typeof(string), typeof(DateTime), typeof(bool),
            typeof(long), typeof(int), typeof(short),
            typeof(double), typeof(float)    };

        public const int Forward = 200;
        public const int ChunkSize = 1000;

        private bool Wroted;

        public Workbook Workbook { get => Worksheet.Workbook; }
        public Worksheet Worksheet { get; }

        public int Row0 { get; }
        public int Column0 { get; }
        public int RowCount { get; private set; }
        public int ColumnCount { get; }

        private Type[] ColumnTypes { get; }

        private bool IgnoreFormat { get; }

        private DateTime nullDate;

        private int loadedRngBgn;
        private int loadedRngEnd;
        private List<object[]> stored;

        public TypedRange(Worksheet wksheet, int row0, int col0, Type[] columnTypes, bool ignoreFormat)
        {
            this.Worksheet = wksheet;

            this.Row0 = row0;
            this.Column0 = col0;

            this.ColumnCount = columnTypes.Length;

            this.ColumnTypes = columnTypes;

            this.IgnoreFormat = ignoreFormat;

            this.nullDate = wksheet.Workbook.NullDate;

            foreach (var type in columnTypes)
                if (!CheckType(type, SupportTypes))
                    throw new ArgumentException("unknown type: " + type.FullName);

            UpdateLastRow();

            loadedRngBgn = -ChunkSize;
            loadedRngEnd = -ChunkSize;
            stored = new List<object[]>();
            ReadChunkAt(Row0);
        }

        #region private method

        private void UpdateLastRow()
        {
            var lastRow = Worksheet.LastCell.Row0;

            var startRow = (lastRow - Row0) / ChunkSize * ChunkSize + Row0;
            for (
                var bgnRngRow = startRow;
                bgnRngRow >= Row0;
                bgnRngRow -= ChunkSize)
            {
                var endRngRow = bgnRngRow == startRow ? lastRow : bgnRngRow + ChunkSize;

                var values = Worksheet.Range(
                        bgnRngRow, Column0,
                        endRngRow, Column0 + ColumnCount)
                    .GetValue(true);

                for (var ridx = values.Length - 1; ridx >= 0; --ridx)
                {
                    for (var cidx = 0; cidx < ColumnCount; ++cidx)
                    {
                        var val = values[ridx][cidx];
                        if (val != null && !(val is string && string.IsNullOrEmpty((string)val)))
                        {
                            RowCount = bgnRngRow + ridx - Row0 + 1;
                            return;
                        }
                    }
                }
            }
            RowCount = 0;
        }

        private void ReadChunkAt(int idx)
        {
            var reqRngBgn = Math.Max(Row0, idx - Forward);
            var reqRngEnd = Math.Min(reqRngBgn + ChunkSize - 1, Row0 + Math.Max(1, RowCount - 1));

            if (Wroted) loadedRngBgn = loadedRngEnd = -1;

            if (reqRngBgn == loadedRngBgn) return;

            if (reqRngBgn < loadedRngBgn && loadedRngBgn < reqRngEnd)
            {
                var values = Worksheet.Range(
                        reqRngBgn, Column0,
                        loadedRngBgn - 1, Column0 + ColumnCount - 1)
                    .GetValue(true);

                var deleteAt = reqRngEnd - loadedRngBgn;
                stored.RemoveRange(deleteAt, ChunkSize - deleteAt);

                stored.InsertRange(0, ConvertLines(reqRngBgn, values));
            }
            else if (loadedRngBgn < reqRngBgn && reqRngBgn < loadedRngEnd)
            {
                var values = Worksheet.Range(
                        loadedRngEnd + 1, Column0,
                        reqRngEnd, Column0 + ColumnCount - 1)
                    .GetValue(true);

                stored.RemoveRange(0, reqRngBgn - loadedRngBgn);
                stored.AddRange(ConvertLines(loadedRngEnd + 1, values));
            }
            else
            {
                var values = Worksheet.Range(
                        reqRngBgn, Column0,
                        reqRngEnd, Column0 + ColumnCount - 1)
                    .GetValue(true);

                stored.Clear();

                stored.AddRange(ConvertLines(reqRngBgn, values));
            }

            loadedRngBgn = reqRngBgn;
            loadedRngEnd = reqRngEnd;
        }

        private object[][] ConvertLines(int row, object[][] lines)
        {
            for (var ridx = 0; ridx < lines.Length; ++ridx)
            {
                var line = lines[ridx];

                for (var idx = 0; idx < ColumnTypes.Length; ++idx)
                {
                    var value = line[idx];
                    var type = ColumnTypes[idx];

                    if (typeof(string).IsAssignableFrom(type))
                    {
                        if (!IgnoreFormat && value is double)
                            line[idx] = Worksheet.CellAt(row + ridx, Column0 + idx).Text;

                        else line[idx] = value.ToString();

                        continue;
                    }

                    if (typeof(DateTime).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                            line[idx] = (value is string) ?
                                DateTime.Parse((string)value) :
                                Utils.ConvertValueToDate(nullDate, (double)value);

                        continue;
                    }

                    if (typeof(bool).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                            line[idx] = (value is string) ?
                                Boolean.Parse((string)value) :
                                (double)value != 0;

                        continue;
                    }

                    if (typeof(long).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                            line[idx] = (value is string) ?
                                long.Parse((string)value) :
                                (long)(double)value;

                        continue;
                    }

                    if (typeof(int).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                            line[idx] = (value is string) ?
                                Int32.Parse((string)value) :
                                (int)(double)value;

                        continue;
                    }

                    if (typeof(double).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                            line[idx] = (value is string) ?
                                Double.Parse((string)value) :
                                (double)value;

                        continue;
                    }

                    if (typeof(short).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                        {
                            var num = (value is string) ?
                                Int32.Parse((string)value) :
                                (int)(double)value;

                            line[idx] = (short)num;
                        }

                        continue;
                    }

                    if (typeof(float).IsAssignableFrom(type))
                    {
                        if (!string.Empty.Equals(value))
                        {
                            var num = (value is string) ?
                                Double.Parse((string)value) :
                                (double)value;

                            line[idx] = (float)num;
                        }

                        continue;
                    }

                    throw new ArgumentException("unknown type: " + type.FullName);
                }
            }

            return lines;
        }

        private static bool CheckType(Type target, params Type[] list)
        {
            foreach (var t in list) if (t.IsAssignableFrom(target)) return true;
            return false;
        }

        #endregion

        public object[] Line(int idx)
        {
            if (idx < 0 || idx >= RowCount) throw new IndexOutOfRangeException();

            var rowIdx = idx + Row0;

            if (Wroted || rowIdx < loadedRngBgn || loadedRngEnd < rowIdx)
            {
                ReadChunkAt(rowIdx);
                Wroted = false;
            }

            return stored[rowIdx - loadedRngBgn];
        }

        public void Write(IEnumerable<object[]> lines)
        {
            var buffer = new List<object[]>();

            foreach (var line in lines)
            {
                buffer.Add(line);

                if (buffer.Count == 1000)
                {
                    Write(buffer.ToArray());
                    buffer.Clear();
                }
            }

            if (buffer.Count != 0)
            {
                Write(buffer.ToArray());
                buffer.Clear();
            }
        }
        private void Write(object[][] datas)
        {
            var outBgnRow = Row0 + RowCount;

            var rng = Worksheet.Range(
                outBgnRow, Column0,
                outBgnRow + datas.Length - 1, Column0 + ColumnCount - 1);

            if (!IgnoreFormat)
            {
                for (var i = 0; i < ColumnTypes.Length; ++i)
                {
                    var ctype = ColumnTypes[i];
                    if (typeof(bool).IsAssignableFrom(ctype))
                    {
                        Worksheet.Range(
                            outBgnRow, Column0 + i,
                            outBgnRow + datas.Length - 1, Column0 + i)
                            .FormatType = FormatType.Boolean;
                    }
                    if (typeof(DateTime).IsAssignableFrom(ctype))
                    {
                        Worksheet.Range(
                            outBgnRow, Column0 + i,
                            outBgnRow + datas.Length - 1, Column0 + i)
                            .FormatType = FormatType.DateTime;
                    }
                }
            }

            rng.SetValue(datas, true);
            Wroted = true;
            RowCount += datas.Length;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public IEnumerator<object[]> GetEnumerator()
        {
            throw new NotImplementedException();
        }

    }

    class TypedRangeEnumerator : IEnumerator<object[]>
    {
        int idx;
        TypedRange owner;

        public TypedRangeEnumerator(TypedRange owner)
        {
            this.owner = owner;
            Reset();
        }

        object IEnumerator.Current => Current;
        public object[] Current => owner.Line(idx);

        public void Dispose() { }

        public bool MoveNext()
        {
            return ++idx < owner.RowCount;
        }

        public void Reset()
        {
            idx = -1;
        }
    }
}
