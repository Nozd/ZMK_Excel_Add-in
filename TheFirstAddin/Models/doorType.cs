using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TheFirstAddin
{
    public class doorType
    {
        public doorType()
        {
            IsAngular = true;
        }
        public string GraphName { get; set; }
        public string PassportName { get; set; }
        public bool IsDouble { get; set; }//Является двухстворчатой
        public bool IsAngular { get; set; }//Является угловой
    }
}
