// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");

using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

//ここからプログラミング

// これはテスト
Console.WriteLine("Excelテストデータの表示");
Console.WriteLine(Environment.CurrentDirectory);

// Excelデータの表示
const string path = @"C:\Users\myros\OneDrive\デスクトップ\授業メモ\卒業制作\卒業制作_Test\卒業制作　学年テストデータ.xlsx";
XLWorkbook book = new XLWorkbook(path);
var worksheet = book.Worksheet("Sheet1");

// カスタム日付の書式設定
var dateFormat = "yyyy/m/d/"; 

// 日付のセル範囲を指定
var dateRange = worksheet.Range("F2:F13");

// 日付のセル範囲に対して日付の書式を設定
dateRange.Style.NumberFormat.Format = dateFormat;

// 列ごとにデータを読み取り、列ごとに表示
for (int row = 1; row <= 13; row++)
{
    for (int column = 1; column <= 6; column++)
    {
        var cell = worksheet.Cell(row, column);
        var cellValue = cell.Value;

        Console.Write($"{cellValue}\t");
    }
    Console.WriteLine(); // 改行して次の行に移動
}

// Excelファイルを保存
book.Save();
