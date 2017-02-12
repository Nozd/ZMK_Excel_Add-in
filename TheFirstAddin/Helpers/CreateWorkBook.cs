using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TheFirstAddin
{
    public static partial class Helpers
    {
        public static Excel.Workbook CreateNewWorkBook(Excel.Application xl)
        {
            string fileName = string.Concat(Environment.UserName, 
                "-",
                string.Format("{0:yyyyMMdd_HHmmss}", DateTime.Now),
                ".xlsx");
            string customerAppConfigFilePath = AppConfig.GetCustomerAppConfigPath();
            string filePathDeault = @"D:\";
            string filePath;
            if (!string.IsNullOrEmpty(customerAppConfigFilePath))
            {
                using (AppConfig.Change(customerAppConfigFilePath))
                {
                    filePath = Directory.Exists(ConfigurationManager.AppSettings["configPackingListSheetFolder"])
                        ? ConfigurationManager.AppSettings["configPackingListSheetFolder"]
                        : filePathDeault;
                }
            }
            else
            {
                filePath = filePathDeault;}
            //this.Application.Visible = true;
            xl.Visible = true;
            xl.SheetsInNewWorkbook = 1;
            xl.Visible = true;
            Excel.Workbook newWorkbook = (Excel.Workbook)(xl.Workbooks.Add(Missing.Value));
            //Excel.Workbook newWorkbook = this.Application.Workbooks.Add(missing);
            newWorkbook.SaveAs(string.Concat(filePath, fileName), Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return newWorkbook;
        }
    }
}
