// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("Excelワークブックからデータをインポート");

var sw1 = new System.Diagnostics.Stopwatch();
var workbook1 = new Workbook();

sw1.Start();

// Excelファイルを読み込む
workbook1.Open("random-numbers-array.xlsx");

// ワークシート追加
var worksheet1 = workbook1.Worksheets.Add();

// 新規シートにデータをコピー
workbook1.Worksheets[0].Range[0, 0, 1000, 1000].Copy(worksheet1.Range[0, 0, 1000, 1000]);

// 既存シートを削除
workbook1.Worksheets[0].Delete();

sw1.Stop();

// 結果の表示
Console.WriteLine("処理時間（Open、Add、Copy、Delete）");
Console.WriteLine($"　{sw1.ElapsedMilliseconds}ミリ秒");

workbook1.Save("result1.xlsx");

var sw2 = new System.Diagnostics.Stopwatch();
var workbook2 = new Workbook();

sw2.Start();

// Excelファイルからデータを読み込む
var data = Workbook.ImportData("random-numbers-array.xlsx", "Sheet1", 0, 0, 1000, 1000);
workbook2.Worksheets[0].Range[0, 0, 1000, 1000].Value = data;

sw2.Stop();

// 結果の表示
Console.WriteLine("処理時間（ImportData）");
Console.WriteLine($"　{sw2.ElapsedMilliseconds}ミリ秒");

workbook2.Save("result2.xlsx");


