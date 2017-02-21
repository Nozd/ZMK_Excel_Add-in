using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pns = TheFirstAddin.PassportNameSet.Enum;

namespace TheFirstAddin
{
    public static class DescriptionMainLeafSet
    {
        public static Dictionary<PassportNameSet.Enum, string> Dic = new Dictionary<PassportNameSet.Enum, string>();
        private static readonly PassportNameSet.Enum[] Set1 = new PassportNameSet.Enum[]
        {
            pns.DM_100, 
            pns.DPM_01_30k,
            pns.DPM_01_60k
        };
        private static readonly PassportNameSet.Enum[] Set2 = new PassportNameSet.Enum[]
        {
            pns.DM_200
        };
        private static readonly PassportNameSet.Enum[] Set3 = new PassportNameSet.Enum[]
        {
            pns.DPM_02_30k,
            pns.DPM_02_60k
        };

        static DescriptionMainLeafSet()
        {
            foreach (var item in Set1)
            {
                Dic.Add(item, "Дверь в сборе с установленным замком, ригелем, петлевыми подшипниками и ответной планкой" );
            }
            foreach (var item in Set2)
            {
                Dic.Add(item, "Створка замковая с установленным замком и ригелем");
            }
            foreach (var item in Set3)
            {
                Dic.Add(item, "Створка замковая с установленным замком, ригелем и термоблокиратором");
            }
        }
    }
}
