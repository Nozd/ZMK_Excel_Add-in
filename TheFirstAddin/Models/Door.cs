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
            _doorType = new doorType();
        }
        public bool IsIdentified { get; set; }//Дверь определена

        private doorType _doorType;
        public doorType DoorType
        {
            get { return _doorType; }
            set
            {
                _doorType = value;
                IsIdentified = true;
            } 
        }

        public string NumberDoor { get; set; }//Номер двери
        public int Height { get; set; }//Высота коробки
        public int Width { get; set; }//Ширина коробки
        public int WidthWorkLeaf { get; set; }//Ширина рабочей створки
        public bool IsThreeLoop { get; set; }//Является трёхпетлевой
        public bool IsCollapsible { get; set; }//Является разборной
        public int GlazingHeight { get; set; }//Высота остекления
        public int GlazingWidth { get; set; }//Ширина остекления
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
