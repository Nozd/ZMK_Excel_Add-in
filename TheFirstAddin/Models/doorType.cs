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
        public string GraphName { get; set; }//Название в графике
        //public string PassportName { get; set; }//Название в паспорте
        public PassportNameSet.Enum PassportNameEnum { get; set; }
        public string DescriptionMainLeaf { get; set; }//Первая позиция в паспорте: описание основной или рабочей створки
        public string DescriptionSecondLeaf { get; set; }//Описание ответной створки
        public bool IsDouble { get; set; }//Является двухстворчатой
        public bool IsAngular { get; set; }//Является угловой
    }
}
