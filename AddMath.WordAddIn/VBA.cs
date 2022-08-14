using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddMath.WordAddIn
{
    public partial class VBA : Form
    {
        public VBA()
        {
            InitializeComponent();
        }

        private void VBA_Load(object sender, EventArgs e)
        {
            VbaCodeTextBox.Rtf = Properties.Resources.VBA;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            app.CustomizationContext = app.ActiveDocument.get_AttachedTemplate();
            app.KeyBindings.Add(WdKeyCategory.wdKeyCategoryMacro, "AddMathAddIn", 220);
            app.KeyBindings.Add(WdKeyCategory.wdKeyCategoryMacro, "AddSlash", app.BuildKeyCode(WdKey.wdKeyAlt, 220));
        }

        private void OpenVbaEditor_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ShowVisualBasicEditor = true;
        }

        private void CopyButton_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(VbaCodeTextBox.Text);
        }
    }
}
