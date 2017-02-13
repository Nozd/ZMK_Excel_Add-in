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
#region Одностворки
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
                {new doorType
                {
                    GraphName = "ДПМ-160-ЛУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_01_60k
                }},
                {new doorType
                {
                    GraphName = "ДПМ-160-ПУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_01_60k
                }},
    #region Одностворки с остеклением
                    {new doorType
                    {
                        GraphName = "ДПМО-130-ЛУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_01_30k
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-130-ПУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_01_30k
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-160-ЛУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_01_60k
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-160-ПУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_01_60k
                    }},
    #endregion
#endregion
#region Двухстворки
                {new doorType
                {
                    GraphName = "ДПМ-200-ЛУ",
                    PassportNameEnum = PassportNameSet.Enum.DM_200
                }},
                {new doorType
                {
                    GraphName = "ДПМ-200-ПУ",
                    PassportNameEnum = PassportNameSet.Enum.DM_200
                }},
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
                {new doorType
                {
                    GraphName = "ДПМ-260-ПУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_02_30k
                }},
                {new doorType
                {
                    GraphName = "ДПМ-260-ЛУ",
                    PassportNameEnum = PassportNameSet.Enum.DPM_02_30k
                }},
    #region Двухстворки с остеклением
                    {new doorType
                    {
                        GraphName = "ДПМО-200-ЛУ",
                        PassportNameEnum = PassportNameSet.Enum.DM_200
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-200-ПУ",
                        PassportNameEnum = PassportNameSet.Enum.DM_200
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-230-ЛУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_02_30k
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-230-ПУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_02_30k
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-260-ЛУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_02_60k
                    }},
                    {new doorType
                    {
                        GraphName = "ДПМО-260-ПУ",
                        PassportNameEnum = PassportNameSet.Enum.DPM_02_60k
                    }},
    #endregion
#endregion
#region Специальные двери 
                    {new doorType
                    {
                    GraphName = "ДМ-100-ЛСпец-(МД-5)",
                    PassportNameEnum = PassportNameSet.Enum.MD_5,
                    IsSpecial = true
                    }},
                    {new doorType
                    {
                    GraphName = "ДМ-100-ПСпец-(МД-5)",
                    PassportNameEnum = PassportNameSet.Enum.MD_5,
                    IsSpecial = true
                    }},
                    {new doorType
                    {
                    GraphName = "ДМ-100-ЛТ-(МД-7)",
                    PassportNameEnum = PassportNameSet.Enum.MD_7,
                    IsSpecial = true
                    }},
                    {new doorType
                    {
                    GraphName = "ДМ-100-ПТ-(МД-7)",
                    PassportNameEnum = PassportNameSet.Enum.MD_7,
                    IsSpecial = true
                    }},
#endregion
#region Other
                    {new doorType
                    {
                        GraphName = "Порог приставной к двери",
                        PassportNameEnum = PassportNameSet.Enum.ThreSholdAddl,
                        IsSpecial = true
                    }}
#endregion
            };
        }
    }
    
}
