// See https://aka.ms/new-console-template for more information
using FlexCel.Core;
using FlexCel.XlsAdapter;

XlsFile xls = new XlsFile(1, TExcelFileFormat.v2023, true);

xls.SetCellValue(1,1, "Hello World");

xls.SetCellValue(2,1, 3);
xls.SetCellValue(2,2, 4);
xls.SetCellValue(2,3, new TFormula("=Sum(A2:A3)"));

xls.Save(System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "output.xlsx"));

