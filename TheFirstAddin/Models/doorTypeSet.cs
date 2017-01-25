using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin
{
    public class DoorTypeSet
    {
        private const string DPM_01_30k = "ДПМ 01/30к";
        private const string DM_100 = "ДМ 100";

        private List<doorType> doorTS;
        public List<doorType> DoorTS { 
            get { return doorTS; }
            set { doorTS = value; }
        }

        public DoorTypeSet()
        {
            doorTS = new List<doorType>
            {
                {new doorType
                {
                    GraphName = "ДМ-100-ЛУ",
                    PassportName = DM_100
                }},
                {new doorType
                {
                    GraphName = "ДМ-100-ПУ",
                    PassportName = DM_100
                }},
                {new doorType
                {
                    GraphName = "ДПМ-130-ЛУ",
                    PassportName = DPM_01_30k
                }},
                {new doorType
                {
                    GraphName = "ДПМ-130-ПУ",
                    PassportName = DPM_01_30k
                }}
            };
        }
    }
    
}
