using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace EyeCareColor
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
           
        }

        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Start();
        }

        private void btnStop_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Stop();
        }
    }
}
