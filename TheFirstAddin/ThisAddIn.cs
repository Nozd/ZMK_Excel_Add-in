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
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new Ribbon();
            ribbon.ButtonClicked += ribbon_ButtonClicked;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        }

        //Passport Button click
        private void ribbon_ButtonClicked()
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (!ValidateRows(range))
            {
                return;
            }

            //int rowLength = 5;
            //for(int i = 0; i < rowLength; i++)
            ////foreach (object cell in range.Cells)
            //{
            //    try
            //    {
            //        MessageBox.Show(((Excel.Range)range.Cells[i]).Value2.ToString());
            //    }
            //    catch
            //    {
            //        MessageBox.Show("NULL VALUE");
            //    }
            //}
            
        }

        //Validation selection rows
        private bool ValidateRows(Excel.Range range)
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
                foreach (Excel.Range row in area.Rows){
                    
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
        private bool ShowNonValidationMessage()
        {
            MessageBox.Show("Error!" +
                            "\nВыделенная область имеет неверный формат или содеsdfржит недопустимые данные" +
                            "\nСборосьте выделение и выберите заново одну или несколько строк");
            return false;
        }
    }
}
