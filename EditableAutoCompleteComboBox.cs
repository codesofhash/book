using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    public class EditableAutoCompleteComboBox : ComboBox
    {
        private HashSet<string> autoCompleteItems;

        public EditableAutoCompleteComboBox()
        {
            autoCompleteItems = new HashSet<string>();

            // Configure autocomplete behavior
            this.DropDownStyle = ComboBoxStyle.DropDown;
            this.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.AutoCompleteSource = AutoCompleteSource.ListItems;

            // Handle events
            this.Leave += OnLeaveComboBox;
            this.TextChanged += OnTextChanged;
        }

        /// <summary>
        /// Loads autocomplete items from a list
        /// </summary>
        public void LoadAutoCompleteItems(IEnumerable<string> items)
        {
            autoCompleteItems.Clear();
            Items.Clear();

            foreach (string item in items.Where(i => !string.IsNullOrWhiteSpace(i)))
            {
                autoCompleteItems.Add(item);
                Items.Add(item);
            }
        }

        /// <summary>
        /// Adds a new item to autocomplete list
        /// </summary>
        public void AddAutoCompleteItem(string item)
        {
            if (!string.IsNullOrWhiteSpace(item) && !autoCompleteItems.Contains(item))
            {
                autoCompleteItems.Add(item);
                Items.Add(item);
            }
        }

        /// <summary>
        /// Gets all autocomplete items
        /// </summary>
        public List<string> GetAutoCompleteItems()
        {
            return autoCompleteItems.ToList();
        }

        /// <summary>
        /// Clears all items
        /// </summary>
        public void ClearAutoCompleteItems()
        {
            autoCompleteItems.Clear();
            Items.Clear();
            Text = string.Empty;
        }

        private void OnLeaveComboBox(object sender, EventArgs e)
        {
            // Add custom value to list if not already present
            string currentText = this.Text?.Trim();
            if (!string.IsNullOrEmpty(currentText))
            {
                AddAutoCompleteItem(currentText);
            }
        }

        private void OnTextChanged(object sender, EventArgs e)
        {
            // Custom text changed logic if needed
        }

        /// <summary>
        /// Positions this combo box over a DataGridView cell
        /// </summary>
        public void PositionOverCell(DataGridView grid, int rowIndex, int columnIndex)
        {
            if (grid == null || rowIndex < 0 || columnIndex < 0)
                return;

            var rect = grid.GetCellDisplayRectangle(columnIndex, rowIndex, false);
            this.Location = new System.Drawing.Point(rect.X + grid.Left, rect.Y + grid.Top);
            this.Size = new System.Drawing.Size(rect.Width, rect.Height);
            this.Parent = grid.Parent;
            this.BringToFront();
            this.Visible = true;
            this.Focus();
        }

        /// <summary>
        /// Hides the combo box
        /// </summary>
        public void HideComboBox()
        {
            this.Visible = false;
        }
    }
}
