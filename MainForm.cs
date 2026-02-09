using System;
using System.Drawing;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    public partial class MainForm : Form
    {
        public MainForm(int rowCount, int columnCount)
        {
            InitializeComponent();

            string[,] headers =
            {
                { "Programme Name", "100" },
                { "Time", "50" },
                { "Break", "50" },
                { "F/P", "50" },
                { "KWD", "50" },
                { "USD", "50" }

            };

            // Initialize the grid with some data
            customDataGridView.ColumnCount = columnCount;
            customDataGridView.RowCount = rowCount;
            for (int i = 0; i < customDataGridView.RowCount; i++)            {   

                for (int j = 0; j < customDataGridView.ColumnCount; j++)
                {
                    //if (j == 0)
                    //{
                        //customDataGridView.SetColumnName(j, headers[j,0]);
                        //customDataGridView.SetColumnWidth(j, int.Parse(headers[j, 1]));
                        //customDataGridView.Rows[i].Cells[j].Value = $"R{i} C{j}";
                    //}

                    if (j == 0 || j < 6)
                    {
                        customDataGridView.SetColumnName(j, headers[j, 0]);
                        customDataGridView.SetColumnWidth(j, int.Parse(headers[j, 1]));                        
                    }

                    customDataGridView.Rows[i].Cells[j].Value = $"R{i} C{j}";

                }
            }
        }

        private void copyButton_Click(object sender, EventArgs e)
        {
            customDataGridView.Copy();
        }

        private void cutButton_Click(object sender, EventArgs e)
        {
            customDataGridView.Cut();
        }

        private void pasteButton_Click(object sender, EventArgs e)
        {
            customDataGridView.Paste();
        }

        private void insertRowButton_Click(object sender, EventArgs e)
        {
            customDataGridView.InsertRow();
        }

        private void insertColumnButton_Click(object sender, EventArgs e)
        {
            customDataGridView.InsertColumn();
        }

        private void deleteRowButton_Click(object sender, EventArgs e)
        {
            customDataGridView.DeleteRow();
        }

        private void deleteColumnButton_Click(object sender, EventArgs e)
        {
            customDataGridView.DeleteColumn();
        }

        private void renameColumnButton_Click(object sender, EventArgs e)
        {
            if (customDataGridView.CurrentCell != null)
            {
                string currentName = customDataGridView.Columns[customDataGridView.CurrentCell.ColumnIndex].HeaderText;
                string newName = InputDialog.Show("Rename Column", "Enter new column name:", currentName);
                if (!string.IsNullOrEmpty(newName))
                {
                    customDataGridView.SetColumnName(customDataGridView.CurrentCell.ColumnIndex, newName);
                }
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.C))
            {
                customDataGridView.Copy();
                return true;
            }
            if (keyData == (Keys.Control | Keys.X))
            {
                customDataGridView.Cut();
                return true;
            }
            if (keyData == (Keys.Control | Keys.V))
            {
                customDataGridView.Paste();
                return true;
            }
            if (keyData == (Keys.Control | Keys.I))
            {
                customDataGridView.InsertRow();
                return true;
            }
            if (keyData == (Keys.Control | Keys.Shift | Keys.I))
            {
                customDataGridView.InsertColumn();
                return true;
            }
            if (keyData == (Keys.Control | Keys.D))
            {
                customDataGridView.DeleteRow();
                return true;
            }
            if (keyData == (Keys.Control | Keys.Shift | Keys.D))
            {
                customDataGridView.DeleteColumn();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void processBookingButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select a Booking Order File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var reader = new BookingOrderReader();
                        string jsonOutput = reader.ProcessBookingOrder(openFileDialog.FileName);

                        // Display the JSON in a new form
                        ShowJsonOutput(jsonOutput);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ShowJsonOutput(string json)
        {
            Form resultForm = new Form
            {
                Text = "Booking Order JSON Output",
                Size = new Size(600, 800),
                StartPosition = FormStartPosition.CenterParent
            };

            RichTextBox richTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Text = json,
                ReadOnly = true,
                Font = new Font("Consolas", 10)
            };

            resultForm.Controls.Add(richTextBox);
            resultForm.ShowDialog();
        }
    }

    public static class InputDialog
    {
        public static string Show(string title, string promptText, string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            return dialogResult == DialogResult.OK ? textBox.Text : "";
        }
    }
}