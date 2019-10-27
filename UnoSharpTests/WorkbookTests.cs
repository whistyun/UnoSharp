using NUnit.Framework;
using System.IO;
using System.Threading;

namespace UnoSharp.Tests
{
    public class WorkbookTests: SetupBase
    {
        [Test]
        public void OpenTest()
        {
            using (var book = new Workbook()) ;
            using (var book = new Workbook("TestForSheet.ods")) ;

            try
            {
                using (var book = new Workbook("NotFoundFiles.ods")) ;

                Assert.Fail("Must be throw exception!");
            }
            catch (FileNotFoundException e) { }
        }

        [Test]
        public void SaveAsTest()
        {
            using (var book = new Workbook())
            {
                book.SaveAs("save.ods");
            }

            Assert.IsTrue(File.Exists("save.ods"));
        }

        [Test]
        public void SaveTest()
        {
            var firstLastWriteTime = File.GetLastWriteTime("TestForSave.ods");

            using (var book = new Workbook("TestForSave.ods"))
            {
                Thread.Sleep(1000);
                book.Save();
            }

            Assert.IsTrue(firstLastWriteTime < File.GetLastWriteTime("TestForSave.ods"));
        }
    }
}