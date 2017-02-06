using System;
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
                    //
#region Парсим остекление
                    string graphGlazing = sheet.Cells[y, 18].Value.Trim(new[] { ' ' });
                    graphGlazing = string.IsNullOrEmpty(graphGlazing) ? "" : graphGlazing.Trim(new[] { ' ' }).ToLower();
                    if (string.Equals(graphGlazing, "нет") || string.IsNullOrEmpty(graphGlazing))
                    {
                        door.DoorType.IsGlazed = false;
                    }
                    else
                    {
                        door.DoorType.IsGlazed = true;
                        graphGlazing = Regex.Replace(graphGlazing, "[^0-9\\*]", "");
                        string[] graphGlazingDivided = graphGlazing.Split(new[] {'*'});
                        int glathingWidth, glathingHeight;
                        door.GlazingWidth = int.TryParse(graphGlazingDivided[0], out glathingWidth) ? glathingWidth : 0;
                        door.GlazingHeight = int.TryParse(graphGlazingDivided[1], out glathingHeight) ? glathingHeight : 0;
                    }
#endregion

#region Парсим наименование
                    string graphWholeName = sheet.Cells[y, 5].Value;
                    graphWholeName = string.IsNullOrEmpty(graphWholeName) ? "" : graphWholeName.Trim(new[] { ' ' }).ToLower();
                    //Идентификация двери
                    string[] graphWholeNameDivided = graphWholeName.Split(new string[] { " ", "(", ")", ".", "mm", "мм" }, StringSplitOptions.None);
                    foreach (var dtsItem in dts.DoorTS)
                    {
                        if (String.Equals(graphWholeNameDivided[0], dtsItem.GraphName))
                        {
                            door.DoorType = dtsItem;
                            break;
                        }
                    }
                    //Является ли двухстворчатой
                    if (door.IsIdentified && PassportNameSet.Dic[door.DoorType.PassportNameEnum].Contains('2'))
                    {
                        door.DoorType.IsDouble = true;
                    }
   
#endregion
#region Парсим размеры двери
                    string graphDoorSize = sheet.Cells[y, 8].Value.Trim(new[] { ' ' });
                    graphDoorSize = string.IsNullOrEmpty(graphDoorSize) ? "" : graphDoorSize.Trim(new[] { ' ' }).ToLower();
                    string[] graphDoorSizeDivided = graphDoorSize.Split(new char[]{'*', 'x', 'х'});
                    int h, w, wwl;
                    door.Width = int.TryParse(graphDoorSizeDivided[0], out w) ? w : 0;
                    door.Height = int.TryParse(graphDoorSizeDivided[1], out h) ? h : 0;
                    if (door.DoorType.IsDouble)
                    {
                        //string[] equalLeafsName = {"равные"};
                        if (!graphWholeName.Contains("равн"))
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
#endregion
    #region Является ли разборной
                    door.IsCollapsible =
                        door.DoorType.IsAngular && (door.Width > 1450 || door.Height > 2250)
                        || !door.DoorType.IsAngular && (door.Width > 1550 || door.Height > 2300);
    #endregion
                    #region Определение кол-ва петель
                    door.IsThreeLoop =
                        door.DoorType.IsGlazed && (door.GlazingWidth > 300 || door.GlazingHeight > 400)
                        || door.DoorType.IsDouble
                        && (
                            door.DoorType.IsAngular && (door.WidthWorkLeaf >= 1000 || door.Height >= 2200)
                            || !door.DoorType.IsAngular && (door.WidthWorkLeaf >= 1050 || door.Height >= 2250)
                            )
                        || !door.DoorType.IsDouble
                        && (
                            door.DoorType.IsAngular && (door.Height >= 2200 || door.Width >= 1000)
                            || !door.DoorType.IsAngular && (door.Height >= 2250 || door.Width >= 1100)
                            );
                   
    #endregion
#region Парсим тип порога
                    string graphThreshold = sheet.Cells[y, 10].Value;
                    graphThreshold = string.IsNullOrEmpty(graphThreshold) ? "" : graphThreshold.Trim(new[] { ' ' }).ToLower();
                    foreach (KeyValuePair<ThresholdSet.Enum, string> item in ThresholdSet.Dic)
                    {
                        if (string.Equals(item.Value, graphThreshold))
                        {
                            door.DoorType.Threshold = item.Key;
                            break;
                        }
                    }
#endregion
#region Парсим тип запирающего механизма
                    string lockType = sheet.Cells[y, 14].Value;
                    lockType = string.IsNullOrEmpty(lockType) ? "" : lockType.Trim(new[] { ' ' }).ToLower();
                    foreach (KeyValuePair<LockSet.Enum, string> item in LockSet.Dic)
                    {
                        if (string.Equals(item.Value.ToLower(), lockType))
                        {
                            door.DoorType.Lock = item.Key;
                            break;
                        }
                    }
#endregion
                    //Описание основной/рабочей створки
                    door.Internals.Add(new Door.Internal(DescriptionMainLeafSet.Dic[door.DoorType.PassportNameEnum], 1, UnitSet.Dic[UnitSet.Enum.Kit]));
                    if (door.DoorType.IsDouble)
                    {
                        door.Internals.Add(new Door.Internal(DescriptionSecondLeafSet.Dic[door.DoorType.PassportNameEnum], 1, UnitSet.Dic[UnitSet.Enum.Kit]));
                    }
                    if (door.DoorType.IsDouble)
                    {
                        if (door.IsCollapsible)
                        {
                            door.Internals.Add(new Door.Internal("Коробка дверная", 1, UnitSet.Dic[UnitSet.Enum.Kit]));
                            door.Internals.Add(new Door.Internal("Болт М10х50", 2, UnitSet.Dic[UnitSet.Enum.Thing]));
                            switch (door.DoorType.Threshold)
                            {
                                case ThresholdSet.Enum.Hight:
                                case ThresholdSet.Enum.Low:
                                    door.Internals.Add(new Door.Internal("Гайка М10", 6, UnitSet.Dic[UnitSet.Enum.Thing]));
                                    door.Internals.Add(new Door.Internal("Шайба М10", 8, UnitSet.Dic[UnitSet.Enum.Thing]));
                                    break;
                                case ThresholdSet.Enum.Mounting:
                                    door.Internals.Add(new Door.Internal("Гайка М10", 2, UnitSet.Dic[UnitSet.Enum.Thing]));
                                    door.Internals.Add(new Door.Internal("Шайба М10", 4, UnitSet.Dic[UnitSet.Enum.Thing]));
                                    door.Internals.Add(new Door.Internal("Саморез 4,2х13 мм", 2,
                                        UnitSet.Dic[UnitSet.Enum.Thing]));
                                    break;
                                default:
                                    break;
                            }

                        }
                        else
                        {
                            door.Internals.Add(new Door.Internal("Коробка дверная в сборе", 1,
                                UnitSet.Dic[UnitSet.Enum.Kit]));
                        }
                    }
                    else
                    {
                        door.Internals.Add(new Door.Internal("Коробка дверная в сборе", 1, UnitSet.Dic[UnitSet.Enum.Kit]));
                    }
                    door.Internals.Add(new Door.Internal("Цилиндр замка с комплектом ключей и винтом", 1, UnitSet.Dic[UnitSet.Enum.Kit]));
                    door.Internals.Add(new Door.Internal("Ручка со стяжными винтами и накладками", 1, UnitSet.Dic[UnitSet.Enum.Kit]));
                    door.Internals.Add(new Door.Internal(SquareSet.Dic[door.DoorType.Lock], 1, UnitSet.Dic[UnitSet.Enum.Thing]));
                    int cap72Count = 0;//Кол-во заглушек 72
                    int cap112Count = 0;//Кол-во заглушек 112
                    if (door.DoorType.IsDouble)
                    {
                        switch (door.DoorType.Threshold)
                        {
                            case ThresholdSet.Enum.Hight:
                            case ThresholdSet.Enum.Low:
                                cap72Count = 2;
                                break;
                            default:
                                cap72Count = 0;
                                break;
                        }
                        cap112Count = door.Height < 1700 ? 6 : 8;
                    }
                    else
                    {
                        switch (door.DoorType.Threshold)
                        {
                            case ThresholdSet.Enum.Hight:
                            case ThresholdSet.Enum.Low:
                                cap72Count = 1;
                                break;
                            default:
                                cap72Count = 0;
                                break;
                        }
                        cap112Count = door.Height < 1700 ? 4 : 6;
                        door.Internals.Add(new Door.Internal("Дюбель рамный 10х112", cap112Count, UnitSet.Dic[UnitSet.Enum.Thing]));
                    }
                    if (cap72Count > 0)
                    {
                        door.Internals.Add(new Door.Internal("Дюбель рамный 10х72", cap72Count, UnitSet.Dic[UnitSet.Enum.Thing]));
                    }
                    if (cap112Count > 0)
                    {
                        door.Internals.Add(new Door.Internal("Дюбель рамный 10х112", cap112Count,
                            UnitSet.Dic[UnitSet.Enum.Thing]));
                    }
                    if (cap72Count + cap112Count > 0)
                    {
                        door.Internals.Add(new Door.Internal("Заглушка Ø16 мм", cap72Count + cap112Count, UnitSet.Dic[UnitSet.Enum.Thing]));
                    }
                    door.Internals.Add(new Door.Internal("Шайба регулировочная", (door.IsThreeLoop ? 3 : 2)*(door.DoorType.IsDouble ? 2 : 1)*5,
                            UnitSet.Dic[UnitSet.Enum.Thing]));
                    if (door.DoorType.IsDouble)
                    {
                        door.Internals.Add(new Door.Internal("Подшипник петлевой",
                            (door.IsThreeLoop ? 3 : 2)*(door.DoorType.IsDouble ? 2 : 1), UnitSet.Dic[UnitSet.Enum.Thing]));
                    }
                    door.Internals.Add(new Door.Internal("Паспорт", 1, UnitSet.Dic[UnitSet.Enum.Thing]));
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
