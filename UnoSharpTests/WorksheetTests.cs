using NUnit.Framework;
using UnoSharp;

namespace UnoSharp.Tests
{
    public class WorksheetTests : SetupBase
    {
        [Test]
        public void SheetTest()
        {
            using (var book = new Workbook())
            {
                book.Worksheets.Add("Sheet2");
                book.Worksheets.Add("Sheet3");
                book.Worksheets.Add("Sheet4");

                for (int i = 0; i < 4; ++i)
                {
                    Assert.AreEqual("Sheet" + (i + 1), book.Worksheets[i].Name);

                }
                for (int i = 0; i < 4; ++i)
                {
                    Assert.AreEqual("Sheet" + (i + 1), book.Worksheets["Sheet" + (i + 1)].Name);

                }

                book.Worksheets[0].Name = "Worksheet0";
                book.Worksheets[1].Name = "Worksheet1";
                book.Worksheets[2].Name = "Worksheet2";
                book.Worksheets[3].Name = "Worksheet3";

                for (int i = 1; i < 4; ++i)
                {
                    Assert.AreEqual("Worksheet" + i, book.Worksheets[i].Name);

                }

                int idx = 0;
                foreach (var sht in book.Worksheets)
                {
                    Assert.AreEqual("Worksheet" + idx, book.Worksheets[idx].Name);
                    ++idx;
                }
            }
        }

        [Test]
        public void SheetTest2()
        {
            using (var book = new Workbook("TestForSheet.ods"))
            {
                for (var i = 0; i < 3; ++i)
                {
                    Assert.AreEqual(
                        book.Worksheets[i][0, 0].Text,
                        book.Worksheets["Sheet" + i].CellAt("$A$1").Text);

                    Assert.AreEqual(
                        book.Worksheets[i][0, 0].Text,
                        "Test" + i);
                }
            }
        }

    }
}
