using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin.Data
{
    public static class ThresholdSet
    {
        public enum Enum
        {
            Hight,
            Low,
            Mounting,
            Default
        }
        public static Dictionary<Enum, string> Dic = new Dictionary<Enum, string>
        {
            {Enum.Hight, "высокий"},
            {Enum.Low, "низкий"},
            {Enum.Mounting, "монтажный"},
            {Enum.Default, ""}
        }; 
    }
}
