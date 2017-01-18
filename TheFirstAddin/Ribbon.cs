using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace TheFirstAddin
{
    public partial class Ribbon
    {
        public event Action ButtonClicked;
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Passport_Click(object sender, RibbonControlEventArgs e)
        {
            if (ButtonClicked != null)
                ButtonClicked();
        }
    }
}
