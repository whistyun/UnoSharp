# UnoSharp

UnoSharp is the simple wrapper to operate LibreOffice by Universal Network Objects (hereinafter UNO).

Classes and Methods are somewhat similar to the Microsoft Office Excel model.

## Example

```cs
// new workbook
var book1 = new Workbook();
book1.Worksheets.Add("new sheet");
book1.SaveAs(@"D:\test1.ods");
book1.Close();

// open workbook
using (var book2 = new Workbook(@"D:\test2.ods"))
{
    Worksheet sheet = book2.Worksheets["Sheet1"];
    string text = sheet.CellAt("A1").Text;
    double value = sheet.CellAt("A1").Value;
}
```

## Setup LibreOffce

To use UnoSharp, We should install LibreOffice and LibreOffice SDK.
Please match 64bit or 32bit with the application to be executed: If you use 64bit application, choose 64bit LibreOffice.

If you execute AnyCPU application on 64bit system, and are not sure whether it will work on 64bit, I recommend to install both 32bit and 64bit.
