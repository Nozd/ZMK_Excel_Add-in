using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin
{
    public static class DescriptionMainLeafSet
    {
        public static Dictionary<PassportNameSet.Enum, string> Dic = new Dictionary<PassportNameSet.Enum, string>()
        {
            //Одностворки
            {
                PassportNameSet.Enum.DM_100,
                "Дверь в сборе с установленным замком, ригелем, петлевыми подшипниками и ответной планкой"
            },
            {
                PassportNameSet.Enum.DPM_01_30k,
                "Дверь в сборе с установленным замком, ригелем, петлевыми подшипниками и ответной планкой"
            },
            
            //Двухстворки
            {
                PassportNameSet.Enum.DM_200,
                "Створка замковая с установленным замком и ригелем"
            },
            {
                PassportNameSet.Enum.DPM_02_30k,
                "Створка замковая с установленным замком, ригелем и термоблокиратором"
            },
        };
    }
}
