﻿using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Xsl;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using TheFirstAddin.Data;

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

#region Passport Button click
        private void ribbon_ButtonClicked()
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            if (!Validation.ValidateRows(range))
            {
                //return;
            }

            //Parser
            List<Door> doorList = Parser.ParseForm(Globals.ThisAddIn.Application);
            //

            //Create new workBook
            this.Application.Visible = true;
            Excel.Workbook newWorkbook = this.Application.Workbooks.Add(missing);
            newWorkbook.SaveAs(@"D:\Book1.xlsx", missing,
                missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);

            //Fill form
            FillingForm.FillForm(newWorkbook, doorList);

           

        }
#endregion
    }
}
