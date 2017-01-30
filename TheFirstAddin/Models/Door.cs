using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace TheFirstAddin
{
    public class Door
    {
        public string NameDoor { get; set; }//Название двери
        public string NumberDoor { get; set; }//Номер двери
        //public bool IsDouble { get; set; }//Является двухстворчатой
        public int Height { get; set; }//Высота коробки
        public int Width { get; set; }//Ширина коробки
        public int WidthWorkLeaf { get; set; }//Ширина рабочей створки
        public bool IsThreeLoop { get; set; }//Является трёхпетлевой
    }
}
