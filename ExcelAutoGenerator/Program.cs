// See https://aka.ms/new-console-template for more information
using Microsoft.AspNetCore.Mvc;
using System.IO;
using ExcelAutoGenerator;
using System.Collections.Generic;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Drawing;
using NPOI.XWPF.UserModel;
using NPOI.SS.Formula.Functions;
using Microsoft.AspNetCore.Routing;
using System.Diagnostics.CodeAnalysis;

Console.WriteLine("Application Started!");
string fromDate = string.Empty;//09/01/2023";
string toDate = string.Empty;//"09/10/2023";
DateTime fromDateT = DateTime.Now;
DateTime toDateT = DateTime.Now;
string OutpathFolder = @"C:\\Users\\Public\\Downloads\\ExcelGenerator\\" + fromDateT.ToString("MM") + "_" + fromDateT.ToString("YYYY");
Console.WriteLine("Enter from date(MM/DD/YYYY):");
fromDate = Console.ReadLine();
Console.WriteLine("Enter to date(MM/DD/YYYY):");
toDate = Console.ReadLine();
fromDateT = Convert.ToDateTime(fromDate);
toDateT = Convert.ToDateTime(toDate);
if (!Directory.Exists(OutpathFolder))
{
    Directory.CreateDirectory(OutpathFolder);
}
int dys = 1;
string inputPathFile = "Readychange.xls";
string OutpathFile = OutpathFolder + "\\" + fromDateT.ToString("MM") + "_" + fromDateT.ToString("yyyy") + ".xls";
//File.Copy(inputPathFile, OutpathFile, true);
using (var fs = new FileStream(OutpathFile, FileMode.Create, FileAccess.Write))
{
    IWorkbook workbook = new XSSFWorkbook();
    foreach (DateTime day in Helper.EachDay(fromDateT, toDateT))
    {
        ISheet sheet1 = workbook.CreateSheet(day.ToString("dd MMM"));
        sheet1.SetColumnWidth(0, 2000);
        sheet1.SetColumnWidth(1, 25000);

        
        IRow row;
        int Room = 101;
        int middleCounter = 0;
        for (int i = 0; i < 43; i++)
        {
            //style
            XSSFFont defaultFont = (XSSFFont)workbook.CreateFont();
            defaultFont.FontHeightInPoints = (short)12;
            defaultFont.FontName = "Arial";
            defaultFont.Color = IndexedColors.Black.Index;
            defaultFont.IsBold = true;
            XSSFCellStyle firstRowStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            firstRowStyle.SetFont(defaultFont);
            firstRowStyle.Alignment = HorizontalAlignment.Center;
            firstRowStyle.VerticalAlignment = VerticalAlignment.Center;
            firstRowStyle.WrapText = true;
            // create bordered cell style

            if (i == 0)
            {
                row = sheet1.CreateRow(i);
                row.Height = 320;
                row.CreateCell(0).SetCellValue("Room");
                row.CreateCell(1).SetCellValue(day.ToString("dd MMMM yyyy dddd"));
                firstRowStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                firstRowStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                firstRowStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                firstRowStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                row.GetCell(0).CellStyle = firstRowStyle;
                row.GetCell(1).CellStyle = firstRowStyle;
            }
            else
            {
                row = sheet1.CreateRow(i);
                //row.RowStyle = firstRowStyle;
                if (i > 1 && i % 3 == 0)
                {
                    CellRangeAddress cellMerge = new CellRangeAddress(i-2, i, 0, 0);
                    sheet1.AddMergedRegion(cellMerge);
                    Room = Room + 1;
                }
                row.Height = 335;
                row.CreateCell(0).SetCellValue(Room.ToString());
                row.CreateCell(1).SetCellValue("");
                row.GetCell(0).CellStyle = firstRowStyle;
                if (middleCounter == 0)
                {
                    firstRowStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
                    row.GetCell(1).CellStyle = firstRowStyle;
                    middleCounter++;
                }
                else if (middleCounter == 1)
                {
                    middleCounter++;
                    firstRowStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
                    firstRowStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Dotted;
                    row.GetCell(1).CellStyle = firstRowStyle;
                }
                else if (middleCounter > 1)
                {
                    middleCounter = 0;
                    firstRowStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Dotted;
                    firstRowStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    firstRowStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                    row.GetCell(1).CellStyle = firstRowStyle;
                }
            }
        }
        Console.WriteLine("Day added :" + dys +"/" + Helper.EachDay(fromDateT, toDateT).Count());
        dys++;
    }
    workbook.Write(fs);
}
Console.WriteLine("Excel Generated Successfully!");
Console.ReadLine();
Environment.Exit(0);