# UnoSharp

UnoSharpは、Universal Network Objects(以下、UNO)を使用し、
LibreOfficeを操作するための簡易レイヤーです。

各種クラス・メソッドはMicrosoft Office Excelのモデルに、ある程度は似せています。

## サンプル

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

## セットアップ

UnoSharpを使用するには、LibreOffice本体と、LibreOffice SDKが必要です。 UnoSharpを使用するアプリが32bitアプリの場合は32bit版のLibreOffice及びSDKを、64bitアプリの場合は64bitのものを使用してください。

### AnyCPUの場合

AnyCPUでアプリケーションをビルドする場合、64bit環境では、「32bit優先」でビルドするか否かにより、32bitで動くか64bitで動くか変わります。「32bit優先」でも、そのアプリを別のアプリが呼び出す場合、64bitで動作する可能性があります(例えば、別のアプリが64bitの場合)。

32bitと64bitの何方で動作するか確信が持てない場合は、32bit版と64bit版の両方のLibreOfficeとSDKをインストールしてください。