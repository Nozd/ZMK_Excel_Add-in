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
            if (!Validate.ValidateRows(range))
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
    }
}
