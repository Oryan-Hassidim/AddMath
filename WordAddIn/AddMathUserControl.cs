using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn
{
    public partial class AddMathUserControl : UserControl
    {
        public AddMathUserControl()
        {
            InitializeComponent();
        }

        private void AddMathUserControl_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
        }
    }
}
