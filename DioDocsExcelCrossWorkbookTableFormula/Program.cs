// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("クロスワークブック数式で外部ワークブックのテーブルを参照する");

//Workbook.SetLicenseKey("");

// 集計用のワークブックを読み込み
Workbook workbook = new();
workbook.Open("sendai.xlsx");

// 外部参照を設定
workbook.Worksheets[0].Range["B4"].Formula = "=SUM('[aoba.xlsx]'!テーブル1[小計])";
workbook.Worksheets[0].Range["B5"].Formula = "=SUM('[izumi.xlsx]'!テーブル1[小計])";
workbook.Worksheets[0].Range["B6"].Formula = "=SUM('[miyagino.xlsx]'!テーブル1[小計])";
workbook.Worksheets[0].Range["B7"].Formula = "=SUM('[taihaku.xlsx]'!テーブル1[小計])";

// 外部ワークブックを読み込み（青葉区）
Workbook aoba = new();
aoba.Open("aoba.xlsx");

// 外部ワークブックを読み込み（泉区）
Workbook izumi = new();
izumi.Open("izumi.xlsx");

// 外部ワークブックを読み込み（宮城野区）
Workbook miyagino = new();
miyagino.Open("miyagino.xlsx");

// 外部ワークブックを読み込み（太白区）
Workbook taihaku = new();
taihaku.Open("taihaku.xlsx");

// 外部参照を更新
workbook.UpdateExcelLink("aoba.xlsx", aoba);
workbook.UpdateExcelLink("izumi.xlsx", izumi);
workbook.UpdateExcelLink("miyagino.xlsx", miyagino);
workbook.UpdateExcelLink("taihaku.xlsx", taihaku);
workbook.Calculate();

// EXCELファイル（.xlsx）に保存
workbook.Save("result.xlsx");
