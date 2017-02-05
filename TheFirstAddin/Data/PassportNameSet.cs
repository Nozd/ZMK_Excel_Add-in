﻿using System;
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
            DPM_02_60k
            
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
            {Enum.DPM_02_60k,"ДПМ 02/60к"}
        };
        
        //private const string DPM_01_30k = "ДПМ 01/30к";
        //private const string DM_100 = "ДМ 100";
    }
}
