using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace TheFirstAddin
{
    public static class FillingForm
    {
        public static void FillForm(Excel.Workbook wb, List<Door> doorList)
        {
            foreach (var door in doorList)
            {
                Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
                //
                Excel.Range activeCell;
                //sh.Range[sh.Cells[4, 1], sh.Cells[6, 1]].HorizontalAlignment = 1;//TODO: посмотреть, зачем писалась эта строка
                sh.Range[sh.Cells[1, 1], sh.Cells[1, 1]].ColumnWidth = 15;
                sh.Range[sh.Cells[1, 2], sh.Cells[1, 2]].ColumnWidth = 1.4;
                sh.Range[sh.Cells[1, 3], sh.Cells[1, 3]].ColumnWidth = 9;
                sh.Range[sh.Cells[1, 4], sh.Cells[1, 4]].ColumnWidth = 36;
                sh.Range[sh.Cells[1, 5], sh.Cells[1, 5]].ColumnWidth = 18;
                sh.Range[sh.Cells[1, 6], sh.Cells[1, 6]].ColumnWidth = 1;
                sh.Range[sh.Cells[1, 7], sh.Cells[1, 7]].ColumnWidth = 8;
                double relHeight = Math.Truncate(0.9 * (sh.Range[sh.Cells[1, 4], sh.Cells[1, 4]].ColumnWidth + sh.Range[sh.Cells[1, 3], sh.Cells[1, 3]].ColumnWidth) / sh.Range[sh.Cells[1, 3], sh.Cells[1, 3]].ColumnWidth);
                //
                sh.Cells[1, 1] = "Упаковочный лист*";
                activeCell = sh.Range[sh.Cells[1, 1], sh.Cells[1, 7]];
                activeCell.Merge(true);
                activeCell.EntireRow.RowHeight = 31.5;
                activeCell.Font.Size = 24;
                activeCell.Font.Bold = true;
                activeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                activeCell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //
                sh.Cells[2, 1] = String.Concat(PassportNameSet.Dic[door.DoorType.PassportNameEnum]," №", door.NumberDoor);
                activeCell = sh.Range[sh.Cells[2, 1], sh.Cells[2, 7]];
                activeCell.Merge(true);
                activeCell.Font.Size = 14;
                activeCell.Font.Bold = true;
                activeCell.HorizontalAlignment = 3;
                activeCell.VerticalAlignment = 3;
                //
                int rowNumber = 4;
                int rowPosition = 1;
                foreach (var Internal in door.Internals)
                {
                    sh.Range[sh.Cells[rowNumber, 1], sh.Cells[rowNumber, 7]].WrapText = true;
                    sh.Cells[rowNumber, 1] = rowPosition;
                    sh.Cells[rowNumber, 2] = ".";
                    sh.Cells[rowNumber, 3] = Internal.Name;
                    sh.Range[sh.Cells[rowNumber, 3], sh.Cells[rowNumber, 4]].Merge(true);
                    var nextRowHeight = sh.Range[sh.Cells[rowNumber + 1, 1], sh.Cells[rowNumber + 1, 1]].RowHeight;
                    if (sh.Range[sh.Cells[rowNumber, 1], sh.Cells[rowNumber, 1]].RowHeight > nextRowHeight)
                    {
                        var tempHeight = sh.Range[sh.Cells[rowNumber, 1], sh.Cells[rowNumber, 1]].RowHeight/relHeight;
                        tempHeight = Math.Truncate(tempHeight/nextRowHeight)*nextRowHeight;
                        sh.Range[sh.Cells[rowNumber, 1], sh.Cells[rowNumber, 1]].RowHeight = tempHeight > 0 ? tempHeight : nextRowHeight;
                    }
                    sh.Cells[rowNumber, 5] = Internal.Count;
                    sh.Cells[rowNumber, 7] = Internal.Unit;
                    ++rowNumber;
                    ++rowPosition;
                }
                activeCell = sh.Range[sh.Cells[4, 1], sh.Cells[rowNumber, 7]];
                activeCell.VerticalAlignment = 1;
                //Подпись
                ++rowNumber;
                sh.Cells[rowNumber, 1] = "Упаковщик";
                sh.Range[sh.Cells[rowNumber, 2], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Cells[rowNumber, 5] = "Мосина Н.Г.";
                ++rowNumber;
                sh.Cells[rowNumber, 3] = "подпись";
                sh.Cells[rowNumber, 4] = "дата";
                sh.Range[sh.Cells[rowNumber, 4], sh.Cells[rowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[rowNumber, 5] = "Ф.И.О.";
                ++rowNumber;
                sh.Cells[rowNumber, 1] = "Контролёр";
                sh.Range[sh.Cells[rowNumber, 2], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Cells[rowNumber, 5] = "";
                ++rowNumber;
                sh.Cells[rowNumber, 3] = "подпись";
                sh.Cells[rowNumber, 4] = "дата";
                sh.Range[sh.Cells[rowNumber, 4], sh.Cells[rowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[rowNumber, 5] = "Ф.И.О.";
                ++rowNumber;
                sh.Cells[rowNumber, 1] = "Мастер участка";
                sh.Range[sh.Cells[rowNumber, 2], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Cells[rowNumber, 5] = "Зельнев А.Н.";
                ++rowNumber;
                sh.Cells[rowNumber, 3] = "подпись";
                sh.Cells[rowNumber, 4] = "дата";
                sh.Range[sh.Cells[rowNumber, 4], sh.Cells[rowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[rowNumber, 5] = "Ф.И.О.";
                //
                sh.Range[sh.Cells[1, 1], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[1, 1], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[1, 1], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[1, 1], sh.Cells[rowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                //
            }
            wb.Save();
        }
    }
}
