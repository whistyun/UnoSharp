using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace UnoSharp.Tests
{
    public class TypeRanteTests : SetupBase
    {
        [Test]
        public void ReadTest()
        {
            using (var book = new Workbook(@"TestForTypeRange.ods"))
            {

                var columnTypes = new[] { typeof(DateTime), typeof(int), typeof(double), typeof(string) };

                var sheet = book.Worksheets[0];
                var trange = sheet.BuildRange("A2", columnTypes);

                var values = new[] {
                    new object[]{ new DateTime(2019,1,1), 1, 1.1, "abc" },
                    new object[]{ new DateTime(2019,1,2), 2, 2.2, "1月1日" },
                    new object[]{ new DateTime(2019,1,3), 3, 3.3, "あいう" },
                    new object[]{ new DateTime(2019,1,4), 4, 4.4, "TRUE" },
                    new object[]{ new DateTime(2019,1,5), 5, 5.5, "FALSE" },
                    new object[]{ new DateTime(2019,1,6), 6, 6.6, "123" },
                };

                Assert.AreEqual(values.Length, trange.RowCount);
                for (var i = 0; i < values.Length; ++i)
                {
                    var expected = values[i];
                    var actual = trange.Line(i);

                    MatchArray1(expected, actual);
                }
            }
        }

        [Test]
        public void ReadBlockTest1()
        {
            using (var book = new Workbook(@"TestForTypeRange.ods"))
            {
                var columnTypes = new[] { typeof(string), typeof(double) };

                var sheet = book.Worksheets[1];
                var trange = sheet.BuildRange("A9", columnTypes);

                Assert.AreEqual(4000, trange.RowCount);

                for (var i = 0; i < 4000; ++i)
                {
                    var line = trange.Line(i);

                    Assert.AreEqual(
                        i == 10 ? new object[] { "", "" } :
                               new object[] { "TEST" + i.ToString("000"), i }
                        ,
                        line);
                }
            }
        }

        [Test]
        public void ReadBlockTest2()
        {
            using (var book = new Workbook(@"TestForTypeRange.ods"))
            {
                var columnTypes = new[] { typeof(string), typeof(double) };

                var sheet = book.Worksheets[1];
                var trange = sheet.BuildRange("A9", columnTypes);

                Assert.AreEqual(4000, trange.RowCount);

                foreach (var i in new[] { 0, 1001, 500, 3999, 0, 2000, 1800, 2799, 1799 })
                {
                    var line = trange.Line(i);

                    Assert.AreEqual(
                        i == 10 ? new object[] { "", "" } :
                               new object[] { "TEST" + i.ToString("000"), i }
                        ,
                        line);
                }
            }
        }

        [Test]
        public void WriteBlockTest()
        {
            using (var workbook = new Workbook())
            {
                var ndate = workbook.NullDate;
                var sht = workbook.Worksheets[0];
                var trng = sht.BuildRange("A1", new Type[] { typeof(bool), typeof(int), typeof(DateTime), typeof(string) });


                var inputting = new List<object[]>();
                for (var i = 0; i < 4000; ++i)
                {
                    inputting.Add(new object[] {
                        (i&1)==0,
                        i,
                        Utils.ConvertValueToDate(ndate, i),
                        "Txt"+i
                    });
                }

                trng.Write(inputting);
                trng.Write(inputting);

                workbook.SaveAs("WriteBlockTestResult.ods");

                Assert.AreEqual(inputting.Count * 2, trng.RowCount);
                for (var i = 0; i < trng.RowCount; ++i)
                {
                    MatchArray1(inputting[i % inputting.Count], trng.Line(i), "At" + i);
                }
            }
            var iii = 23;
        }

        private void MatchArray1(object[] expectedLine, object[] actualLine, string msg = "")
        {
            Assert.AreEqual(actualLine.Length, actualLine.Length, "col");

            for (int c = 0; c < expectedLine.Length; ++c)
            {
                Assert.AreEqual(expectedLine[c], actualLine[c], msg + ":" + String.Format("<{0}>", c));
            }
        }
    }
}
