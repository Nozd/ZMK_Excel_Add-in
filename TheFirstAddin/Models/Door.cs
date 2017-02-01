using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace TheFirstAddin
{
    public class Door
    {
        public Door()
        {
            Internals = new List<Internal>();
        }
        //public string NameDoor { get; set; }//Название двери
        public doorType DoorType { get; set; }
        public string NumberDoor { get; set; }//Номер двери
        public int Height { get; set; }//Высота коробки
        public int Width { get; set; }//Ширина коробки
        public int WidthWorkLeaf { get; set; }//Ширина рабочей створки
        public bool IsThreeLoop { get; set; }//Является трёхпетлевой
        public List<Internal> Internals { get; set; } 

        public class Internal
        {
            public Internal(string Name, int Count, string Unit)
            {
                this.Name = Name;
                this.Count = Count;
                this.Unit = Unit;
            }
            public string Name { get; set; }
            public int Count { get; set; }
            public string Unit { get; set; }
        }
    }
}
