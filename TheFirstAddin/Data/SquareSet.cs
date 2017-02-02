using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin.Data
{
    public static class SquareSet
    {
        public static Dictionary<LockSet.Enum, string> Dic = new Dictionary<LockSet.Enum, string>
        {
            {LockSet.Enum.GBS_81, "Квадрат"},
            {LockSet.Enum.GBS_83, "Квадрат разрезной"}
        };
    }
}
