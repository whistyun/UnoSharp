using NUnit.Framework;
using UnoSharp;
using System;

namespace UnoSharp.Tests
{
    public class RangeTests : SetupBase
    {
        [Test]
        public void ReadTest()
        {
            using (var book = new Workbook("TestForRead.ods"))
            {
                var sheet = book.Worksheets[0];

                for (var i = 0; i < 4; ++i)
                {
                    Assert.AreEqual(i + 1, sheet.CellAt(i, 0).Value);
                    Assert.AreEqual(i + 1, sheet.CellAt("A" + (i + 1)).Value);

                }

                Assert.AreEqual(true, sheet.CellAt(0, 1).Flag);
                Assert.AreEqual(true, sheet.CellAt("B1").Flag);
                Assert.AreEqual(true, sheet.Range("B1")[0, 0].Flag);

                Assert.AreEqual(false, sheet.CellAt(0, 2).Flag);
                Assert.AreEqual(false, sheet.CellAt("C1").Flag);
                Assert.AreEqual(false, sheet.Range("C1")[0, 0].Flag);

                Assert.AreEqual("aiueo", sheet.CellAt(1, 1).Text);
                Assert.AreEqual("あいうえお", sheet.CellAt(1, 2).Text);

                Assert.AreEqual(-12.34, sheet.CellAt(2, 1).Value);
                Assert.AreEqual(new DateTime(2019, 10, 22), sheet.CellAt(2, 2).Date);

                Assert.AreEqual(
                    new DateTime(
                        book.NullDate.Year, book.NullDate.Month, book.NullDate.Day,
                        12, 12, 00),
                    sheet.CellAt(3, 2).Date);
            }

        }

        [Test]
        public void ReadRangeTest()
        {
            using (var book = new Workbook("TestForRead.ods"))
            {
                var sheet = book.Worksheets[0];
                var range = sheet.Range("A1:C4");

                var datas = range.Values;

                int y = book.NullDate.Year;
                int m = book.NullDate.Month;
                int d = book.NullDate.Day;

                var expecteds = new object[][] {
                    new object[]{ 1d,    true, false                          },
                    new object[]{ 2d, "aiueo", "あいうえお"                   },
                    new object[]{ 3d,  -12.34, new DateTime(2019, 10, 22)     },
                    new object[]{ 4d,    "", new DateTime(y,m,d,12, 12, 00) }
                };

                Assert.AreEqual(expecteds.Length, range.RowCount, "row");
                Assert.AreEqual(expecteds[0].Length, range.ColumnCount, "col");

                MatchArray2(expecteds, datas);
            }
        }

        [Test]
        public void WriteVariousTypeTest()
        {
            using (var book = new Workbook())
            {
                var sheet = book.Worksheets[0];
                var range = sheet.Range("A1:D2");

                range.Values = new object[][] {
                    new object[]{ (short)1,(long)2,(float)3,(double)4 },
                    new object[]{ new DateTime(2010,1,2),true,"",null },
                };
            }
        }

        [Test]
        public void ReadWriteRangeTest()
        {
            using (var book = new Workbook("TestForRead.ods"))
            {
                var sheet = book.Worksheets[0];
                var range = sheet.Range("A1:C4");

                var datas = range.Values;

                using (var book2 = new Workbook())
                {
                    var sheet2 = book2.Worksheets[0];
                    sheet2.Range("A1:C4").Values = datas;
                    book2.SaveAs("TestForWrite.ods");
                }

                using (var book2 = new Workbook("TestForWrite.ods"))
                {
                    var sheet2 = book2.Worksheets[0];

                    var datas2 = sheet2.Range("A1:C4").Values;
                    MatchArray2(datas, datas2);

                    sheet2.Range("A1:C4").Values = datas;

                    datas2 = sheet2.Range("A1:C4").Values;
                    MatchArray2(datas, datas2);
                }
            }
        }

        [Test]
        public void UseRangeTest()
        {
            using (var book = new Workbook("TestForRange.ods"))
            {
                var
                rnt = book.Worksheets["Sheet0"].UsedRange;
                Assert.AreEqual(2, rnt.Column0);
                Assert.AreEqual(7, rnt.Row0);
                Assert.AreEqual(4, rnt.ColumnCount);
                Assert.AreEqual(5, rnt.RowCount);

                rnt = book.Worksheets["Sheet1"].UsedRange;
                Assert.AreEqual(0, rnt.Column0);
                Assert.AreEqual(0, rnt.Row0);
                Assert.AreEqual(6, rnt.ColumnCount);
                Assert.AreEqual(24, rnt.RowCount);

                rnt = book.Worksheets["Sheet2"].UsedRange;
                Assert.AreEqual(0, rnt.Column0);
                Assert.AreEqual(0, rnt.Row0);
                Assert.AreEqual(3, rnt.ColumnCount);
                Assert.AreEqual(15, rnt.RowCount);

            }
        }

        private void MatchArray2(object[][] expecteds, object[][] inputs)
        {

            Assert.AreEqual(expecteds.Length, inputs.Length, "row");

            for (int r = 0; r < expecteds.Length; ++r)
            {
                var exLine = expecteds[r];
                var inLine = inputs[r];

                Assert.AreEqual(inLine.Length, inLine.Length, "col");

                for (int c = 0; c < exLine.Length; ++c)
                {
                    Assert.AreEqual(exLine[c], inLine[c], String.Format("<{0},{1}>", r, c));
                }
            }
        }
    }
}
