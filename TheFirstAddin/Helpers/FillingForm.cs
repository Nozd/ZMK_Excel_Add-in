using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
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
        public static void FillSheet(Excel.Workbook wb, List<Door> doorList, Excel.Application xl)
        {
            string configPacker;
            string configContoller;
            string configMaster;
            string currentDate;
            string customerAppConfigPath = AppConfig.GetCustomerAppConfigPath();

            if(!string.IsNullOrEmpty(customerAppConfigPath))
            {
                using (
                    AppConfig.Change(customerAppConfigPath)
                    )
                {
                    ApplyConfig(out configPacker, out configContoller, out configMaster, out currentDate);
                }
            }
            else
            {
                ApplyConfig(out configPacker, out configContoller, out configMaster, out currentDate);
            }


            int startRowNumber = 1;
            int sheetCount = 1;
            int doorCount = 0;
            foreach (var door in doorList)
            {
                ++doorCount;
                Excel.Worksheet sh;
                if (startRowNumber == 1 && sheetCount > wb.Sheets.Count)
                {
                    sh = (Excel.Worksheet)wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count], Count: 1, Type: Excel.XlSheetType.xlWorksheet);
                }
                else
                {
                    sh = wb.Sheets[sheetCount];
                }
                try
                {

                    sh.Name = startRowNumber == 1 ? door.NumberDoor : string.Concat(sh.Name, ",", door.NumberDoor);
                }
                catch (Exception e)
                {
                }
                //
                Excel.Range activeCell;
                int currentRowNumber = startRowNumber;
                if (currentRowNumber == 1)
                {
                    //sh.Range[sh.Cells[4, 1], sh.Cells[6, 1]].HorizontalAlignment = 1;//TODO: посмотреть, зачем писалась эта строка
                    sh.Range[sh.Cells[1, 1], sh.Cells[1, 1]].ColumnWidth = 15;
                    sh.Range[sh.Cells[1, 2], sh.Cells[1, 2]].ColumnWidth = 1.4;
                    sh.Range[sh.Cells[1, 3], sh.Cells[1, 3]].ColumnWidth = 9;
                    sh.Range[sh.Cells[1, 4], sh.Cells[1, 4]].ColumnWidth = 36;
                    sh.Range[sh.Cells[1, 5], sh.Cells[1, 5]].ColumnWidth = 18;
                    sh.Range[sh.Cells[1, 6], sh.Cells[1, 6]].ColumnWidth = 1;
                    sh.Range[sh.Cells[1, 7], sh.Cells[1, 7]].ColumnWidth = 8;
                }
                double relHeight =
                        Math.Truncate(0.95 *
                                      (sh.Range[sh.Cells[1, 4], sh.Cells[1, 4]].ColumnWidth +
                                       sh.Range[sh.Cells[1, 3], sh.Cells[1, 3]].ColumnWidth) /
                                      sh.Range[sh.Cells[1, 3], sh.Cells[1, 3]].ColumnWidth);//
                sh.Cells[currentRowNumber, 1] = "Упаковочный лист*";
                activeCell = sh.Range[sh.Cells[currentRowNumber, 1], sh.Cells[currentRowNumber, 7]];
                activeCell.Merge(true);
                activeCell.EntireRow.RowHeight = 31.5;
                activeCell.Font.Size = 24;
                activeCell.Font.Bold = true;
                activeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                activeCell.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 1] = String.Concat(PassportNameSet.Dic[door.DoorType.PassportNameEnum], " №", door.NumberDoor);
                activeCell = sh.Range[sh.Cells[currentRowNumber, 1], sh.Cells[currentRowNumber, 7]];
                activeCell.Merge(true);
                activeCell.Font.Size = 14;
                activeCell.Font.Bold = true;
                activeCell.HorizontalAlignment = 3;
                activeCell.VerticalAlignment = 3;
                //
                currentRowNumber +=2;
                int rowPosition = 1;
                foreach (var Internal in door.Internals)
                {
                    sh.Range[sh.Cells[currentRowNumber, 1], sh.Cells[currentRowNumber, 7]].WrapText = true;
                    sh.Cells[currentRowNumber, 1] = rowPosition;
                    sh.Cells[currentRowNumber, 2] = ".";
                    sh.Cells[currentRowNumber, 3] = Internal.Name;
                    sh.Range[sh.Cells[currentRowNumber, 3], sh.Cells[currentRowNumber, 4]].Merge(true);
                    var nextRowHeight = sh.Range[sh.Cells[currentRowNumber + 1, 1], sh.Cells[currentRowNumber + 1, 1]].RowHeight;
                    if (sh.Range[sh.Cells[currentRowNumber, 1], sh.Cells[currentRowNumber, 1]].RowHeight > nextRowHeight)
                    {
                        var tempHeight = sh.Range[sh.Cells[currentRowNumber, 1], sh.Cells[currentRowNumber, 1]].RowHeight / relHeight;
                        tempHeight = Math.Truncate(tempHeight/nextRowHeight)*nextRowHeight;
                        sh.Range[sh.Cells[currentRowNumber, 1], sh.Cells[currentRowNumber, 1]].RowHeight = tempHeight > 0 ? tempHeight : nextRowHeight;
                    }
                    sh.Cells[currentRowNumber, 5] = Internal.Count;
                    sh.Cells[currentRowNumber, 7] = Internal.Unit;
                    ++currentRowNumber;
                    ++rowPosition;
                }
                activeCell = sh.Range[sh.Cells[4, 1], sh.Cells[currentRowNumber, 7]];
                activeCell.VerticalAlignment = 1;
                //Подпись
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 1] = "Упаковщик";
                sh.Range[sh.Cells[currentRowNumber, 2], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[currentRowNumber, 4], sh.Cells[currentRowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[currentRowNumber, 4] = currentDate;
                sh.Cells[currentRowNumber, 5] = configPacker;
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 3] = "подпись";
                sh.Cells[currentRowNumber, 4] = "дата";
                sh.Range[sh.Cells[currentRowNumber, 4], sh.Cells[currentRowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[currentRowNumber, 5] = "Ф.И.О.";
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 1] = "Контролёр";
                sh.Range[sh.Cells[currentRowNumber, 2], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[currentRowNumber, 4], sh.Cells[currentRowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[currentRowNumber, 4] = currentDate;
                sh.Cells[currentRowNumber, 5] = configContoller;
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 3] = "подпись";
                sh.Cells[currentRowNumber, 4] = "дата";
                sh.Range[sh.Cells[currentRowNumber, 4], sh.Cells[currentRowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[currentRowNumber, 5] = "Ф.И.О.";
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 1] = "Мастер участка";
                sh.Range[sh.Cells[currentRowNumber, 2], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[currentRowNumber, 4], sh.Cells[currentRowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[currentRowNumber, 4] = currentDate;
                sh.Cells[currentRowNumber, 5] = configMaster;
                ++currentRowNumber;
                sh.Cells[currentRowNumber, 3] = "подпись";
                sh.Cells[currentRowNumber, 4] = "дата";
                sh.Range[sh.Cells[currentRowNumber, 4], sh.Cells[currentRowNumber, 4]].HorizontalAlignment = 3;
                sh.Cells[currentRowNumber, 5] = "Ф.И.О.";
                //
                sh.Range[sh.Cells[startRowNumber, 1], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[startRowNumber, 1], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[startRowNumber, 1], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                sh.Range[sh.Cells[startRowNumber, 1], sh.Cells[currentRowNumber, 7]].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    Excel.XlLineStyle.xlContinuous;
                if (startRowNumber != 1 || doorCount == doorList.Count)
                {
                    ++sheetCount;
                    string printArea = string.Concat("A1:", sh.Cells[currentRowNumber, 7].Address);
                    xl.PrintCommunication = false;
                    sh.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA3;
                    sh.PageSetup.PrintArea = printArea;
                    sh.PageSetup.FitToPagesWide = 1;
                    sh.PageSetup.FitToPagesTall = 1;
                    sh.PageSetup.LeftMargin = xl.InchesToPoints(0.2);
                    sh.PageSetup.RightMargin = xl.InchesToPoints(0.2);
                    sh.PageSetup.TopMargin = xl.InchesToPoints(0.2);
                    sh.PageSetup.BottomMargin = xl.InchesToPoints(0.2);
                    sh.PageSetup.HeaderMargin = xl.InchesToPoints(0.15);
                    sh.PageSetup.FooterMargin = xl.InchesToPoints(0.15);
                    xl.PrintCommunication = true;
                }
                startRowNumber = startRowNumber == 1 ? currentRowNumber + 1 : 1;
                
                //
            }
            wb.Save();
        }

        private static void ApplyConfig(out string configPacker, out string configContoller, out string configMaster,
            out string currentDate)
        {
            configPacker = ConfigurationManager.AppSettings["configPacker"];
            configPacker = configPacker ?? "";
            configContoller = ConfigurationManager.AppSettings["configContoller"];
            configContoller = configContoller ?? "";
            configMaster = ConfigurationManager.AppSettings["configMaster"];
            configMaster = configMaster ?? "";
            bool setCurrentDate = !string.IsNullOrEmpty(ConfigurationManager.AppSettings["configSetCurrentDate"])
                             &&
                             string.Equals(ConfigurationManager.AppSettings["configSetCurrentDate"].ToLower(),
                                 "да");
            currentDate = setCurrentDate ? string.Format("{0:dd.MM.yyyy}", DateTime.Now) : "";
        }
    }
}
