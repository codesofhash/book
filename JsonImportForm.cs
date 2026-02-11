using System;
using System.Drawing;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    public partial class JsonImportForm : Form
    {
        public BookingOrder Result { get; private set; }

        public JsonImportForm()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            var result = JsonImportHelper.ValidateAndParse(txtJsonInput.Text);

            if (!result.Success)
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = result.ErrorMessage;
                return;
            }

            lblStatus.ForeColor = Color.Green;
            lblStatus.Text = "Valid JSON. Importing...";

            Result = result.BookingOrder;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
