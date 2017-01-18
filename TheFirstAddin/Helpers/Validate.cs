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
    static class Validate
    {
        //Validation selection rows
        internal static bool ValidateRows(Excel.Range range)
        {
            if (range == null)
            //|| range.Value2 == null
            //|| !(range.Value2 is Array))
            {
                return ShowNonValidationMessage();
            }


            List<int> validRowNumer = new List<int>();
            List<int> noValidRowNumber = new List<int>();
            foreach (Excel.Range area in range.Areas)
            {
                foreach (Excel.Range row in area.Rows)
                {

                    if (row.Value2 == null
                        || !(row.Value2 is Array)
                        || row.Value2.GetLength(0) != 1
                        || row.Value2.GetLength(1) != 256)
                    {
                        noValidRowNumber.Add(row.Row);
                    }
                    else
                    {
                        validRowNumer.Add(row.Row);
                    }
                }
            }
            if (validRowNumer.Count(val => noValidRowNumber.Any(noVal => noVal == val)) != noValidRowNumber.Count)
            {
                return ShowNonValidationMessage();
            }
            //TODO: удалить row, к-рые содержатся в строках
            MessageBox.Show("Aasdll is Oak!");
            int rank;//rank of selection range
            rank = range.Value2.Rank;
            if (rank > 1)
            {
                var b1 = range.Value2.GetLength(0);
                var b2 = range.Value2.GetLength(1);
            }
            return true;
        }
        //Non validation selection rows message
        private static bool ShowNonValidationMessage()
        {
            MessageBox.Show("Error!" +
                            "\nВыделенная область имеет неверный формат или содеsdfржит недопустимые данные" +
                            "\nСборосьте выделение и выберите заново одну или несколько строк");
            return false;
        }
    }
}
