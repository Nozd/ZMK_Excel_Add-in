using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin.Data
{
    public static class DescriptionSecondLeafSet
    {
        public static Dictionary<PassportNameSet.Enum, string> Dic = new Dictionary<PassportNameSet.Enum, string>()
        {
            {
                PassportNameSet.Enum.DM_200,
                "Створка ответная с установленной ответной планкой, ригелем и торцевыми шпингалетами"
            },
            {
                PassportNameSet.Enum.DPM_02_30k,
                "Створка ответная с установленной ответной планкой, ригелем и торцевыми шпингалетами"
            },
        }; 
    }
}
