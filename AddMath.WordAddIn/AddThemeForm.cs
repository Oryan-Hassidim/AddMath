using System;
using System.Windows.Forms;

namespace AddMath.WordAddIn
{
    public partial class AddThemeForm : Form
    {
        public AddThemeForm()
        {
            InitializeComponent();
        }

        private void OkAddButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
