using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin.Data
{
    public static class LockSet
    {
        public enum Enum
        {
            GBS_81,
            GBS_83,
            GBS_81_cylinder
        }
        public static Dictionary<Enum, string> Dic = new Dictionary<Enum, string>
        {
            {Enum.GBS_81, "GBS 81"},
            {Enum.GBS_83, "GBS 83"},
            {Enum.GBS_81_cylinder, "GBS 81+цилиндр ключ-вертушка"}
        }; 
    }
}
