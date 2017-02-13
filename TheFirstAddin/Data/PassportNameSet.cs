using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin
{
    public static class PassportNameSet
    {
        /// <summary>
        /// Наименование дверей
        /// </summary>
        public enum Enum 
        {
            /// <summary>
            /// Одностворки
            /// </summary>
            DM_100,
            DPM_01_30k,
            DPM_01_60k,
            /// <summary>
            /// Двухстворки
            /// </summary>
            DM_200,
            DPM_02_30k,
            DPM_02_60k,
            /// <summary>
            /// Специальные двери
            /// </summary>
            MD_5,
            MD_7,
            /// <summary>
            /// Other
            /// </summary>
            ThreSholdAddl
            
        }
        public static Dictionary<Enum, string> Dic = new Dictionary<Enum, string>()
        {
            //Одностворки
            {Enum.DM_100,  "ДМ 100"},
            {Enum.DPM_01_30k, "ДПМ 01/30к"},
            {Enum.DPM_01_60k, "ДПМ 01/60к"},
            //Двухстворки
            {Enum.DM_200, "ДМ 200"},
            {Enum.DPM_02_30k, "ДПМ 02/30к"},
            {Enum.DPM_02_60k,"ДПМ 02/60к"},
            //Специальные двери
            {Enum.MD_5, "ДМ-100 (МД-5)"},
            {Enum.MD_7, "ДМ-100 (МД-7)"},
            //Other
            {Enum.ThreSholdAddl, "Порог приставной к двери"}
        };
        
        //private const string DPM_01_30k = "ДПМ 01/30к";
        //private const string DM_100 = "ДМ 100";
    }
}
