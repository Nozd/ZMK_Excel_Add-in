using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using pns = TheFirstAddin.PassportNameSet.Enum;

namespace TheFirstAddin.Data
{
    public static class DescriptionSecondLeafSet
    {
        public static Dictionary<PassportNameSet.Enum, string> Dic = new Dictionary<PassportNameSet.Enum, string>();
        static readonly pns[] Set1 = new pns[]
        {
            pns.DM_200,
            pns.DPM_02_30k,
            pns.DPM_02_60k
        };

        static DescriptionSecondLeafSet()
        {
            foreach (var item in Set1)
            {
                Dic.Add(item, "Створка ответная с установленной ответной планкой, ригелем и торцевыми шпингалетами");
            }
        }
    }
}
