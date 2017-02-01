using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin.Data
{
    public static class UnitSet
    {
        public enum Enum
        {
            Thing,
            Kit
        }
        public static Dictionary<Enum, string> Dic = new Dictionary<Enum, string>()
        {
            {Enum.Thing, "шт."},
            {Enum.Kit, "к-кт"}
        };
    }
    
}
