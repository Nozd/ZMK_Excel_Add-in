using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin
{
    public partial class DoorTypeSet
    {
        private List<doorType> doorTS;
        public List<doorType> DoorTS { 
            get { return doorTS; }
            set { doorTS = value; }
        }

        public DoorTypeSet()
        {
            doorTS = new List<doorType>
            {
                //Одностворки
                {new doorType
                {
                    GraphName = "ДМ-100-ЛУ",
                    PassportNameEnum = PassportNameSet.Enum.DM_100
                }},
                {new doorType
                {
                    GraphName = "ДМ-100-ПУ",
                    PassportNameEnum = PassportNameSet.Enum.DM_100
                }},
                {new doorType
                {
                    GraphName = "ДПМ-130-ЛУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_01_30k
                }},
                {new doorType
                {
                    GraphName = "ДПМ-130-ПУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_01_30k
                }},
                //Двухстворки
                {new doorType
                {
                    GraphName = "ДПМ-230-ЛУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_02_30k
                }},
                {new doorType
                {
                    GraphName = "ДПМ-230-ПУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_02_30k
                }},
            };
        }
    }
    
}
