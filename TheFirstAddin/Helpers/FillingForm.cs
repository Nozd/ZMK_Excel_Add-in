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
                const double relHeight = 5.1;
                Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
                //
                var activeCell = sh.Range[sh.Cells[4, 1], sh.Cells[7, 7]];
                activeCell.WrapText = true;
                activeCell.VerticalAlignment = 1;
                sh.Range[sh.Cells[1, 1], sh.Cells[14, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[1, 1], sh.Cells[14, 7]].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[1, 1], sh.Cells[14, 7]].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[1, 1], sh.Cells[14, 7]].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.get_Range("A4", "A6").HorizontalAlignment = 1;
                sh.get_Range("A1", "A1").ColumnWidth = 7;
                sh.get_Range("B1", "B1").ColumnWidth = 7;
                sh.get_Range("C1", "C1").ColumnWidth = 7;
                sh.get_Range("D1", "D1").ColumnWidth = 45;
                sh.get_Range("E1", "E1").ColumnWidth = 7;
                sh.get_Range("F1", "F1").ColumnWidth = 7;
                sh.get_Range("G1", "G1").ColumnWidth = 7;
                //
                sh.Cells[1, 1] = "Упаковочный лист*";
                activeCell = sh.get_Range("A1", "G1");
                activeCell.Merge(true);
                activeCell.EntireRow.RowHeight = 31.5;
                activeCell.Font.Size = 24;
                activeCell.Font.Bold = true;
                activeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                activeCell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //
                sh.Cells[2, 1] = String.Concat(door.NameDoor," №", door.NumberDoor);
                activeCell = sh.get_Range("A2", "G2");
                activeCell.Merge(true);
                activeCell.Font.Size = 14;
                activeCell.Font.Bold = true;
                activeCell.HorizontalAlignment = 3;
                activeCell.VerticalAlignment = 3;
                //
                sh.Cells[4, 1] = "1";
                sh.Cells[4, 2] = ".";
                sh.Cells[4, 3] =
                    "Дверь в сборе с установленным замком, ригелем, петлевыми подшипниками и ответной планкой";
                sh.get_Range("C4", "D4").Merge(true);
                sh.get_Range("A4", "A4").RowHeight = Math.Truncate(sh.get_Range("A4", "A4").RowHeight/relHeight);
                sh.Cells[4, 5] = "1";
                sh.Cells[4, 7] = "к-кт";
                //
                sh.Cells[5, 1] = "2";
                sh.Cells[5, 2] = ".";
                sh.Cells[5, 3] = "Цилиндр замка с комплектом ключей и винтом";
                sh.get_Range("C5", "D5").Merge(true);
                sh.get_Range("A5", "A5").RowHeight = Math.Truncate(sh.get_Range("A5", "A5").RowHeight/relHeight);
                sh.Cells[5, 5] = "1";
                sh.Cells[5, 7] = "к-кт";
                //
                sh.Cells[6, 1] = "3";
                sh.Cells[6, 2] = ".";
                sh.Cells[6, 3] = "Ручка со стяжными винтами и накладками";
                sh.get_Range("C6", "D6").Merge(true);
                sh.get_Range("A6", "A6").RowHeight = Math.Truncate(sh.get_Range("A6", "A6").RowHeight/relHeight);
                sh.Cells[6, 5] = "1";
                sh.Cells[6, 7] = "к-кт";
                //

                //Подпись
                sh.Cells[8, 1] = "Упаковщик";
                sh.Range[sh.Cells[8, 2], sh.Cells[8, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Cells[8, 5] = "Мосина Н.Г.";
                sh.Cells[9, 3] = "подпись";
                sh.Cells[9, 4] = "дата";
                sh.Range[sh.Cells[9, 4], sh.Cells[9, 4]].HorizontalAlignment = 3;
                sh.Cells[9, 5] = "Ф.И.О.";
                //
                sh.Cells[10, 1] = "Контролёр";
                sh.Range[sh.Cells[10, 2], sh.Cells[10, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Cells[10, 5] = "";
                sh.Cells[11, 3] = "подпись";
                sh.Cells[11, 4] = "дата";
                sh.Range[sh.Cells[11, 4], sh.Cells[11, 4]].HorizontalAlignment = 3;
                sh.Cells[11, 5] = "Ф.И.О.";
                //
                sh.Cells[12, 1] = "Мастер участка";
                sh.Range[sh.Cells[12, 2], sh.Cells[12, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Cells[12, 5] = "Зельнев А.Н.";
                sh.Cells[14, 3] = "подпись";
                sh.Cells[14, 4] = "дата";
                sh.Range[sh.Cells[14, 4], sh.Cells[14, 4]].HorizontalAlignment = 3;
                sh.Cells[14, 5] = "Ф.И.О.";
                //
            }
            wb.Save();
        }
    }
}
