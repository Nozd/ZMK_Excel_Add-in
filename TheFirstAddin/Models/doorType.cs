using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TheFirstAddin.Data;

namespace TheFirstAddin
{
    public class doorType
    {
        public doorType()
        {
            IsAngular = true;
        }
        public string GraphName { get; set; }//Название в графике
        public PassportNameSet.Enum PassportNameEnum { get; set; }//Название в паспорте
        public ThresholdSet.Enum Threshold { get; set; }//Тип порога
        public LockSet.Enum Lock { get; set; }//Тип замка
        public string DescriptionMainLeaf { get; set; }//Первая позиция в паспорте: описание основной или рабочей створки
        public string DescriptionSecondLeaf { get; set; }//Описание ответной створки
        public bool IsDouble { get; set; }//Является двухстворчатой
        private bool _isAngular;

        public bool IsAngular
        {
            get
            {
                return string.Equals(GraphName.ToLower()[GraphName.Length - 1], 'у');
            }
            set { _isAngular = value; }
        }//Является угловой
        public bool IsGlazed { get; set; }//Есть ли остекление
    }
}
