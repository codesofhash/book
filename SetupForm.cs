using System;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    public partial class SetupForm : Form
    {
        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }

        public SetupForm()
        {
            InitializeComponent();
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            this.RowCount = (int)this.rowsNumericUpDown.Value;
            this.ColumnCount = (int)this.columnsNumericUpDown.Value;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
