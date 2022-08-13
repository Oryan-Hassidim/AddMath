using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddMath.WordAddIn
{
    public partial class AddMathRibbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.AddMathObject.
            //TODO: multi lingual
        }

        private void AddMathButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddMathObject.AddMath();
        }

        private void EditSuggestions_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void InitializeButton_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Templates.Count == 0)
            {
                MessageBox.Show("Must Be an active template!", "Add Math", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var vba = new VBA();
            vba.Show();
        }
    }
}
