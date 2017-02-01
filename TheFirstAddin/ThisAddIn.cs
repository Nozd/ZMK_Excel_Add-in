using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
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
            DoorTypeSet dts = new DoorTypeSet();
            List<Door> doorList = new List<Door>();
            foreach (Excel.Range area in range.Areas)
            {
                foreach (Excel.Range row in area.Rows)
                {
                    Door door = new Door();

                    int x = row.Column;
                    int y = row.Row;
                    //Номер двери
                    door.NumberDoor = sheet.Cells[y, 4].Value;
                    //Парсим наименование
                    string graphWholeName = sheet.Cells[y, 5].Value.Trim(new[] { ' ' });
                    //Идентификация двери
                    string[] graphWholeNameDivided = graphWholeName.Split(new[] {' ', '(', ')', '.'});
                    foreach (var dtsItem in dts.DoorTS)
                    {
                        if (String.Equals(graphWholeNameDivided[0], dtsItem.GraphName))
                        {
                            door.DoorType = dtsItem;
                            break;
                        }
                    }
                    //Является ли двухстворчатой
                    if (PassportNameSet.Dic[door.DoorType.PassportNameEnum].Contains('2'))
                    {
                        door.DoorType.IsDouble = true;
                    }
                    //Размеры двери
                    int h, w, wwl;
                    door.Height = int.TryParse(graphWholeNameDivided[1], out h) ? h : 0;
                    door.Width = int.TryParse(graphWholeNameDivided[3], out w) ? w : 0;
                    if (door.DoorType.IsDouble)
                    {
                        if (Array.IndexOf(graphWholeNameDivided, "равные") == -1)
                        {
                            int indexWorkLeaf = Array.IndexOf(graphWholeNameDivided, "ств");
                            door.WidthWorkLeaf = indexWorkLeaf > (-1) &&
                                                 int.TryParse(graphWholeNameDivided[indexWorkLeaf + 1], out wwl)
                                ? wwl
                                : 0;
                        }
                        else
                        {
                            door.WidthWorkLeaf = (int)Math.Floor((double)(door.Width / 2));
                        }
                    }
                    //Определение кол-ва петель
                    if (door.DoorType.IsDouble)
                    {
                        door.IsThreeLoop = door.WidthWorkLeaf >= 1000;
                    }
                    else if (door.DoorType.IsAngular)
                    {
                        door.IsThreeLoop = door.Height >= 2200
                                           || door.Width >= 1000;
                                           //|| tempDoorType.IsDouble && door.WidthWorkLeaf >= 1000;
                    }
                    else
                    {
                        door.IsThreeLoop = door.Height >= 2250
                                           || door.Width >= 1100;
                        //|| tempDoorType.IsDouble && door.WidthWorkLeaf >= 1100;
                    }
                    //Описание основной/рабочей створки
                    door.Internals.Add(new Door.Internal(DescriptionMainLeafSet.Dic[door.DoorType.PassportNameEnum], 1, UnitSet.Dic[UnitSet.Enum.Thing]));
                    if (door.DoorType.IsDouble)
                    {
                        door.Internals.Add(new Door.Internal(DescriptionSecondLeafSet.Dic[door.DoorType.PassportNameEnum], 1, UnitSet.Dic[UnitSet.Enum.Thing]));
                    }
                    doorList.Add(door);
                }
            }
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
