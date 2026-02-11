using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace CSharpFlexGrid
{
    public partial class CalendarGridForm : Form
    {
        private BookingOrder bookingOrder;
        private string jsonFilePath;
        private DataTable calendarDataTable;
        private int year;
        private int month;
        private ContextMenuStrip contextMenu;
        private bool isUpdating = false; // Flag to prevent recursive recalculation during batch operations

        // AutoComplete control for Programme Name editing
        private AutoComplete programmeAutoComplete;
        private int autoCompleteEditingRowIndex = -1;

        public CalendarGridForm(BookingOrder order, string filePath)
        {
            InitializeComponent();
            bookingOrder = order;
            jsonFilePath = filePath;
            LoadCalendarData();
            UpdateHeaderLabels();
            SetupContextMenu();
            SetupKeyboardShortcuts();
            SetupEditingHandlers();
            SetupProgrammeAutoComplete();
        }

        private void SetupEditingHandlers()
        {
            dgvCalendar.CellBeginEdit += DgvCalendar_CellBeginEdit;
            dgvCalendar.CellEndEdit += DgvCalendar_CellEndEdit;
            dgvCalendar.CellFormatting += DgvCalendar_CellFormatting;
            dgvCalendar.EditingControlShowing += DgvCalendar_EditingControlShowing;
            dgvCalendar.DataError += DgvCalendar_DataError;
            dgvCalendar.CellValidating += DgvCalendar_CellValidating;
        }

        private void DgvCalendar_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            // For ComboBox columns, add new values to the list automatically
            var column = dgvCalendar.Columns[e.ColumnIndex];
            if (column is DataGridViewComboBoxColumn comboColumn)
            {
                string newValue = e.FormattedValue?.ToString();
                if (!string.IsNullOrEmpty(newValue) && !comboColumn.Items.Contains(newValue))
                {
                    // Add the new value to the ComboBox items
                    comboColumn.Items.Add(newValue);
                }
            }
        }

        private void DgvCalendar_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Handle ComboBox value not in list errors gracefully
            if (e.Exception is ArgumentException && e.Exception.Message.Contains("DataGridViewComboBoxCell"))
            {
                // Get the column and its value
                var column = dgvCalendar.Columns[e.ColumnIndex];
                var cell = dgvCalendar[e.ColumnIndex, e.RowIndex];

                if (column is DataGridViewComboBoxColumn comboColumn)
                {
                    // Add the current value to the ComboBox items if it's not there
                    var currentValue = cell.Value?.ToString();
                    if (!string.IsNullOrEmpty(currentValue) && !comboColumn.Items.Contains(currentValue))
                    {
                        comboColumn.Items.Add(currentValue);
                    }
                }

                e.ThrowException = false;
            }
        }

        private void UpdateHeaderLabels()
        {
            lblAgency.Text = $"Agency: {bookingOrder.Agency}";
            lblAdvertiser.Text = $"Advertiser: {bookingOrder.Advertiser}";
            lblProduct.Text = $"Product: {bookingOrder.Product}";
            lblCampaign.Text = $"Campaign: {bookingOrder.CampaignPeriod.StartDate} to {bookingOrder.CampaignPeriod.EndDate}";
            lblTotalSpots.Text = $"Total Spots: {bookingOrder.TotalSpots}";
        }

        private void LoadCalendarData()
        {
            // Determine the year and month from campaign dates
            var firstDate = DateTime.Parse(bookingOrder.CampaignPeriod.StartDate);
            year = firstDate.Year;
            month = firstDate.Month;
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // Create DataTable with calendar structure
            calendarDataTable = new DataTable();

            // Add metadata columns
            calendarDataTable.Columns.Add("Programs Name", typeof(string));
            calendarDataTable.Columns.Add("Time (KWT)", typeof(string));
            calendarDataTable.Columns.Add("Break In", typeof(string));
            calendarDataTable.Columns.Add("F/P", typeof(string));
            calendarDataTable.Columns.Add("Ratio", typeof(string));
            calendarDataTable.Columns.Add("Sponsor Type", typeof(string));
            calendarDataTable.Columns.Add("ORD", typeof(string));
            calendarDataTable.Columns.Add("Sponsor Price", typeof(string));
            calendarDataTable.Columns.Add("OID", typeof(string));
            calendarDataTable.Columns.Add("Unit Price KWD", typeof(string));
            calendarDataTable.Columns.Add("Price in US $", typeof(string));

            // Add hidden spot index column
            calendarDataTable.Columns.Add("Spot Index", typeof(int));

            // Add date columns with day-of-week headers
            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                string dayOfWeek = date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2);
                string columnName = $"{dayOfWeek}\n{day}";
                calendarDataTable.Columns.Add(columnName, typeof(int));
            }

            // Add Total Spots column
            calendarDataTable.Columns.Add("Total Spots", typeof(int));

            // Group spots by program name and time (preserve original Excel row order)
            var spotGroups = bookingOrder.Spots
                .Select((s, index) => new { Spot = s, Index = index })
                .GroupBy(x => new { x.Spot.ProgrammeName, x.Spot.ProgrammeStartTime })
                .OrderBy(g => g.Min(x => x.Index));

            int spotIndex = 0;
            foreach (var group in spotGroups)
            {
                DataRow row = calendarDataTable.NewRow();

                // Fill metadata columns
                row["Programs Name"] = group.Key.ProgrammeName ?? "";
                row["Time (KWT)"] = group.Key.ProgrammeStartTime ?? "";
                row["Break In"] = "WN"; // Default values - can be enhanced later
                row["F/P"] = "P";
                row["Ratio"] = "";
                row["Sponsor Type"] = "";
                row["ORD"] = "";
                row["Sponsor Price"] = "";
                row["OID"] = "";
                row["Unit Price KWD"] = "";
                row["Price in US $"] = "";
                row["Spot Index"] = spotIndex++;

                // Count spots per date for this program/time combination
                var allDatesForGroup = group.SelectMany(x => x.Spot.Dates).ToList();
                var dateCounts = allDatesForGroup
                    .GroupBy(d => d)
                    .ToDictionary(g => DateTime.Parse(g.Key), g => g.Count());

                int totalSpots = 0;
                for (int day = 1; day <= daysInMonth; day++)
                {
                    var date = new DateTime(year, month, day);
                    string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";

                    if (dateCounts.ContainsKey(date))
                    {
                        int count = dateCounts[date];
                        row[columnName] = count;
                        totalSpots += count;
                    }
                    else
                    {
                        row[columnName] = 0;
                    }
                }

                row["Total Spots"] = totalSpots;
                calendarDataTable.Rows.Add(row);
            }

            // Add Total row at the bottom
            DataRow totalRow = calendarDataTable.NewRow();
            totalRow["Programs Name"] = "Total";
            totalRow["Time (KWT)"] = "";
            totalRow["Break In"] = "";
            totalRow["F/P"] = "";
            totalRow["Ratio"] = "";
            totalRow["Sponsor Type"] = "";
            totalRow["ORD"] = "";
            totalRow["Sponsor Price"] = "";
            totalRow["OID"] = "";
            totalRow["Unit Price KWD"] = "";
            totalRow["Price in US $"] = "";
            totalRow["Spot Index"] = -1; // Special marker for total row

            int grandTotal = 0;
            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";

                int dayTotal = 0;
                foreach (DataRow dataRow in calendarDataTable.Rows)
                {
                    dayTotal += Convert.ToInt32(dataRow[columnName]);
                }
                totalRow[columnName] = dayTotal;
                grandTotal += dayTotal;
            }

            totalRow["Total Spots"] = grandTotal;
            calendarDataTable.Rows.Add(totalRow);

            // Bind to DataGridView
            dgvCalendar.DataSource = calendarDataTable;
            ConfigureGridAppearance();
        }

        private void SetupContextMenu()
        {
            contextMenu = new ContextMenuStrip();

            var addRowItem = new ToolStripMenuItem("Add Row");
            addRowItem.Click += AddRow_Click;
            contextMenu.Items.Add(addRowItem);

            var deleteRowItem = new ToolStripMenuItem("Delete Row");
            deleteRowItem.Click += DeleteRow_Click;
            contextMenu.Items.Add(deleteRowItem);

            contextMenu.Items.Add(new ToolStripSeparator());

            var copyItem = new ToolStripMenuItem("Copy (Ctrl+C)");
            copyItem.Click += Copy_Click;
            contextMenu.Items.Add(copyItem);

            var pasteItem = new ToolStripMenuItem("Paste (Ctrl+V)");
            pasteItem.Click += Paste_Click;
            contextMenu.Items.Add(pasteItem);

            dgvCalendar.ContextMenuStrip = contextMenu;
        }

        private void SetupKeyboardShortcuts()
        {
            dgvCalendar.KeyDown += DgvCalendar_KeyDown;
        }

        private void SetupProgrammeAutoComplete()
        {
            // Create the AutoComplete control as a child of the DataGridView
            programmeAutoComplete = new AutoComplete(
                dgvCalendar,
                new System.Drawing.Point(0, 0),
                150,
                dgvCalendar.Font,
                GetProgrammeNameList()
            );

            // Handle selection - commit value to cell
            programmeAutoComplete.ItemSelected += (selectedItem, commitKey) =>
            {
                if (autoCompleteEditingRowIndex >= 0)
                {
                    var programmeColumn = dgvCalendar.Columns["Programs Name"];
                    if (programmeColumn != null)
                    {
                        dgvCalendar.Rows[autoCompleteEditingRowIndex].Cells[programmeColumn.Index].Value = selectedItem;
                    }
                    programmeAutoComplete.Hide();
                    autoCompleteEditingRowIndex = -1;

                    // Move focus based on commit key
                    if (commitKey == Keys.Tab)
                    {
                        // Move to next cell
                        SendKeys.Send("{TAB}");
                    }
                    else if (commitKey == Keys.Enter)
                    {
                        // Move to next row, same column
                        dgvCalendar.Focus();
                    }
                }
            };

            // Handle cancel - hide without committing
            programmeAutoComplete.EditCancelled += () =>
            {
                programmeAutoComplete.Hide();
                autoCompleteEditingRowIndex = -1;
                dgvCalendar.Focus();
            };

            // Hide AutoComplete when scrolling
            dgvCalendar.Scroll += (s, e) =>
            {
                if (programmeAutoComplete.Visible)
                {
                    programmeAutoComplete.Hide();
                    autoCompleteEditingRowIndex = -1;
                }
            };

            // Use SelectionChanged to detect when user navigates to a Programme cell
            dgvCalendar.SelectionChanged += DgvCalendar_SelectionChanged_Programme;

            // Also use MouseDown for direct clicks
            dgvCalendar.MouseDown += DgvCalendar_MouseDown_Programme;
        }

        private void DgvCalendar_SelectionChanged_Programme(object sender, EventArgs e)
        {
            if (dgvCalendar.CurrentCell == null)
                return;

            int rowIndex = dgvCalendar.CurrentCell.RowIndex;
            int colIndex = dgvCalendar.CurrentCell.ColumnIndex;

            // Ignore invalid indices
            if (rowIndex < 0 || colIndex < 0)
                return;

            // Don't allow editing Total row
            if (rowIndex == calendarDataTable.Rows.Count - 1)
            {
                if (programmeAutoComplete.Visible)
                {
                    programmeAutoComplete.Hide();
                    autoCompleteEditingRowIndex = -1;
                }
                return;
            }

            var column = dgvCalendar.Columns[colIndex];
            if (column.Name != "Programs Name")
            {
                // Not a Programme cell - hide AutoComplete if visible
                if (programmeAutoComplete.Visible)
                {
                    programmeAutoComplete.Hide();
                    autoCompleteEditingRowIndex = -1;
                }
                return;
            }

            // Skip if already editing this row
            if (autoCompleteEditingRowIndex == rowIndex && programmeAutoComplete.Visible)
                return;

            // Show AutoComplete for this cell
            ShowAutoCompleteForCell(rowIndex, colIndex);
        }

        private void DgvCalendar_MouseDown_Programme(object sender, MouseEventArgs e)
        {
            // Get the cell at the click location
            var hitTest = dgvCalendar.HitTest(e.X, e.Y);

            if (hitTest.Type != DataGridViewHitTestType.Cell)
            {
                // Clicked outside cells - hide AutoComplete
                if (programmeAutoComplete.Visible)
                {
                    programmeAutoComplete.Hide();
                    autoCompleteEditingRowIndex = -1;
                }
                return;
            }

            int rowIndex = hitTest.RowIndex;
            int colIndex = hitTest.ColumnIndex;

            // Ignore invalid indices
            if (rowIndex < 0 || colIndex < 0)
                return;

            // Don't allow editing Total row
            if (rowIndex == calendarDataTable.Rows.Count - 1)
                return;

            var column = dgvCalendar.Columns[colIndex];
            if (column.Name != "Programs Name")
            {
                // Clicked on a different column - hide AutoComplete if visible
                if (programmeAutoComplete.Visible)
                {
                    programmeAutoComplete.Hide();
                    autoCompleteEditingRowIndex = -1;
                }
                return;
            }

            // Show AutoComplete for this cell
            ShowAutoCompleteForCell(rowIndex, colIndex);
        }

        private void ShowAutoCompleteForCell(int rowIndex, int colIndex)
        {
            // Get cell bounds relative to DataGridView
            var cellRect = dgvCalendar.GetCellDisplayRectangle(colIndex, rowIndex, false);

            // Debug: Show which cell we're editing in the title bar
            this.Text = $"DEBUG: Editing Row {rowIndex}, Col {colIndex}, Rect: X={cellRect.X}, Y={cellRect.Y}, W={cellRect.Width}";

            // Track which row we're editing
            autoCompleteEditingRowIndex = rowIndex;

            // Position the AutoComplete control over the cell
            programmeAutoComplete.SetPosition(
                new System.Drawing.Point(cellRect.X, cellRect.Y),
                cellRect.Width
            );

            // Set current cell value as initial text
            var currentValue = dgvCalendar.Rows[rowIndex].Cells[colIndex].Value?.ToString() ?? "";
            programmeAutoComplete.Clear();
            programmeAutoComplete.Text = currentValue;

            // Show and activate
            programmeAutoComplete.Activate();
        }

        // Static list methods (can be made dynamic later)
        private List<string> GetProgrammeNameList()
        {
            // Simple static dummy list for debugging
            return new List<string>
            {
                "Apple",
                "Apricot",
                "Avocado",
                "Banana",
                "Blueberry",
                "Cherry",
                "Date",
                "Grape",
                "Mango",
                "Orange",
                "Peach",
                "Pear"
            };
        }

        private List<string> GetBreakInList()
        {
            return new List<string>
            {
                "WN",
                "BB",
                "MB",
                "EB",
                "SB",
                "CB"
            };
        }

        private List<string> GetFPList()
        {
            return new List<string>
            {
                "P",
                "F",
                "S"
            };
        }

        private void ConvertToComboBoxColumn(string columnName, List<string> items)
        {
            int columnIndex = dgvCalendar.Columns[columnName]?.Index ?? -1;
            if (columnIndex == -1)
                return;

            // Get existing column position and data
            var oldColumn = dgvCalendar.Columns[columnIndex];
            int displayIndex = oldColumn.DisplayIndex;

            // Collect all unique existing values from the data
            var existingValues = new HashSet<string>();
            foreach (DataRow row in calendarDataTable.Rows)
            {
                var value = row[columnName]?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    existingValues.Add(value);
                }
            }

            // Create new ComboBox column
            var comboColumn = new DataGridViewComboBoxColumn
            {
                Name = columnName,
                HeaderText = columnName,
                DataPropertyName = columnName,
                Width = oldColumn.Width,
                DisplayIndex = displayIndex,
                FlatStyle = FlatStyle.Flat,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing, // Show as text until edited
                SortMode = DataGridViewColumnSortMode.Programmatic
            };

            // Add existing values first to prevent errors
            foreach (string value in existingValues.OrderBy(v => v))
            {
                if (!comboColumn.Items.Contains(value))
                {
                    comboColumn.Items.Add(value);
                }
            }

            // Add static list items (if not already added)
            foreach (string item in items)
            {
                if (!comboColumn.Items.Contains(item))
                {
                    comboColumn.Items.Add(item);
                }
            }

            // Remove old column and insert new one
            dgvCalendar.Columns.Remove(oldColumn);
            dgvCalendar.Columns.Insert(columnIndex, comboColumn);

            // Style the column
            comboColumn.ReadOnly = false;
            comboColumn.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(250, 250, 250);
        }

        private void DgvCalendar_KeyDown(object sender, KeyEventArgs e)
        {
            // Don't intercept copy/paste when editing a cell - let the editing control handle it
            if (dgvCalendar.IsCurrentCellInEditMode)
                return;

            if (e.Control && e.KeyCode == Keys.C)
            {
                dgvCalendar.Copy();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                PasteToSelectedCells();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Delete)
            {
                DeleteSelectedCells();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void AddRow_Click(object sender, EventArgs e)
        {
            // Get current row index, default to before Total row if nothing selected
            int insertIndex = calendarDataTable.Rows.Count - 1; // Default: before Total row

            if (dgvCalendar.CurrentCell != null)
            {
                int currentRowIndex = dgvCalendar.CurrentCell.RowIndex;
                // Don't insert after Total row
                if (currentRowIndex < calendarDataTable.Rows.Count - 1)
                {
                    insertIndex = currentRowIndex; // Insert at current row position
                }
            }

            DataRow newRow = calendarDataTable.NewRow();

            // Initialize with default values
            newRow["Programs Name"] = "";
            newRow["Time (KWT)"] = "";
            newRow["Break In"] = "WN";
            newRow["F/P"] = "P";
            newRow["Ratio"] = "";
            newRow["Sponsor Type"] = "";
            newRow["ORD"] = "";
            newRow["Sponsor Price"] = "";
            newRow["OID"] = "";
            newRow["Unit Price KWD"] = "";
            newRow["Price in US $"] = "";
            newRow["Spot Index"] = calendarDataTable.Rows.Count - 1;

            // Initialize date columns with 0
            int daysInMonth = DateTime.DaysInMonth(year, month);
            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";
                newRow[columnName] = 0;
            }

            newRow["Total Spots"] = 0;

            calendarDataTable.Rows.InsertAt(newRow, insertIndex);
            RecalculateColumnTotals();
            lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
        }

        private void DeleteRow_Click(object sender, EventArgs e)
        {
            // Get the row index from current cell (works for both cell and row selection)
            int rowIndex = -1;

            if (dgvCalendar.CurrentCell != null)
            {
                rowIndex = dgvCalendar.CurrentCell.RowIndex;
            }
            else if (dgvCalendar.SelectedRows.Count > 0)
            {
                rowIndex = dgvCalendar.SelectedRows[0].Index;
            }

            if (rowIndex < 0)
            {
                MessageBox.Show("Please select a row to delete.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Don't allow deleting the Total row
            if (rowIndex == calendarDataTable.Rows.Count - 1)
            {
                MessageBox.Show("Cannot delete the Total row.", "Invalid Operation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Delete this row?", "Confirm Delete",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                calendarDataTable.Rows.RemoveAt(rowIndex);
                RecalculateColumnTotals();
                lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
            }
        }

        private void Copy_Click(object sender, EventArgs e)
        {
            dgvCalendar.Copy();
        }

        private void Paste_Click(object sender, EventArgs e)
        {
            PasteToSelectedCells();
        }


        private void PasteToSelectedCells()
        {
            if (dgvCalendar.CurrentCell == null)
                return;

            // Cancel any active cell edit to prevent the current cell from overriding pasted values
            dgvCalendar.CancelEdit();

            string clipboardText = Clipboard.GetText();
            if (string.IsNullOrEmpty(clipboardText))
                return;

            int startRow = dgvCalendar.CurrentCell.RowIndex;
            int startCol = dgvCalendar.CurrentCell.ColumnIndex;

            // Don't paste into Total row
            if (startRow == calendarDataTable.Rows.Count - 1)
            {
                MessageBox.Show("Cannot paste into the Total row.", "Invalid Operation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Split on \r\n or \n (preserve empty rows so blanks paste as-is)
            string[] rows = clipboardText.TrimEnd('\r', '\n').Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            // Suspend events and layout during bulk update
            isUpdating = true;
            dgvCalendar.SuspendLayout();

            try
            {
                for (int i = 0; i < rows.Length; i++)
                {
                    int currentRow = startRow + i;
                    if (currentRow >= calendarDataTable.Rows.Count - 1) // Stop before Total row
                        break;

                    string[] cells = rows[i].Split('\t');
                    DataRow dataRow = calendarDataTable.Rows[currentRow];

                    for (int j = 0; j < cells.Length; j++)
                    {
                        int currentCol = startCol + j;
                        if (currentCol >= dgvCalendar.Columns.Count)
                            break;

                        var column = dgvCalendar.Columns[currentCol];
                        if (!column.ReadOnly && column.Visible)
                        {
                            try
                            {
                                // For ComboBox columns, ensure the value is in the items list before setting
                                if (column is DataGridViewComboBoxColumn comboCol)
                                {
                                    string val = cells[j];
                                    if (!string.IsNullOrEmpty(val) && !comboCol.Items.Contains(val))
                                    {
                                        comboCol.Items.Add(val);
                                    }
                                }

                                // Write directly to the DataTable to bypass DataGridView current-cell issues
                                string colName = column.DataPropertyName;
                                if (string.IsNullOrEmpty(colName)) colName = column.Name;
                                Type colType = calendarDataTable.Columns[colName].DataType;

                                if (colType == typeof(int))
                                {
                                    dataRow[colName] = int.TryParse(cells[j], out int intVal) ? intVal : 0;
                                }
                                else
                                {
                                    dataRow[colName] = cells[j];
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            finally
            {
                isUpdating = false;
                dgvCalendar.ResumeLayout();
            }

            dgvCalendar.Refresh();

            // Recalculate totals once at the end
            for (int i = 0; i < rows.Length; i++)
            {
                int currentRow = startRow + i;
                if (currentRow < calendarDataTable.Rows.Count - 1)
                {
                    RecalculateRowTotal(currentRow);
                }
            }
            RecalculateColumnTotals();
        }

        private void DeleteSelectedCells()
        {
            if (dgvCalendar.SelectedCells.Count == 0)
                return;

            // Suspend events during bulk delete
            isUpdating = true;
            dgvCalendar.SuspendLayout();

            try
            {
                foreach (DataGridViewCell cell in dgvCalendar.SelectedCells)
                {
                    // Don't delete cells in Total row or read-only columns
                    if (cell.RowIndex < calendarDataTable.Rows.Count - 1 && !cell.ReadOnly)
                    {
                        cell.Value = dgvCalendar.Columns[cell.ColumnIndex].ValueType == typeof(int) ? 0 : "";
                    }
                }
            }
            finally
            {
                isUpdating = false;
                dgvCalendar.ResumeLayout();
            }

            // Recalculate affected rows once at the end
            var affectedRows = dgvCalendar.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Distinct()
                .Where(r => r < calendarDataTable.Rows.Count - 1);

            foreach (int rowIndex in affectedRows)
            {
                RecalculateRowTotal(rowIndex);
            }
            RecalculateColumnTotals();
        }

        private void DgvCalendar_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Don't allow editing Total row
            if (e.RowIndex == calendarDataTable.Rows.Count - 1)
            {
                e.Cancel = true;
                return;
            }
        }

        private void DgvCalendar_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            string columnName = dgvCalendar.Columns[dgvCalendar.CurrentCell.ColumnIndex].Name;

            // Check if this is a ComboBox column (Programs Name now uses AutoComplete overlay)
            if (columnName == "Break In" || columnName == "F/P")
            {
                if (e.Control is DataGridViewComboBoxEditingControl comboBox)
                {
                    // Remove any previous event handlers to avoid duplicates
                    comboBox.Leave -= ComboBox_Leave;

                    comboBox.DropDownStyle = ComboBoxStyle.DropDown;
                    comboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    comboBox.AutoCompleteSource = AutoCompleteSource.ListItems;

                    // Add event handler to add custom values when user leaves the control
                    comboBox.Leave += ComboBox_Leave;
                }
            }
        }

        private void ComboBox_Leave(object sender, EventArgs e)
        {
            if (sender is DataGridViewComboBoxEditingControl comboBox)
            {
                string currentText = comboBox.Text;
                if (!string.IsNullOrEmpty(currentText) && !comboBox.Items.Contains(currentText))
                {
                    // Add the custom value to the ComboBox items
                    comboBox.Items.Add(currentText);

                    // Also add it to the column's items list
                    if (dgvCalendar.CurrentCell != null)
                    {
                        var column = dgvCalendar.Columns[dgvCalendar.CurrentCell.ColumnIndex];
                        if (column is DataGridViewComboBoxColumn comboColumn)
                        {
                            if (!comboColumn.Items.Contains(currentText))
                            {
                                comboColumn.Items.Add(currentText);
                            }
                        }
                    }
                }
            }
        }

        private void DgvCalendar_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            string columnName = dgvCalendar.Columns[e.ColumnIndex].Name;

            // Time formatting
            if (columnName == "Time (KWT)")
            {
                var cell = dgvCalendar[e.ColumnIndex, e.RowIndex];
                if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                {
                    string formattedTime = FormatTimeInput(cell.Value.ToString());
                    if (formattedTime != null)
                    {
                        cell.Value = formattedTime;
                    }
                }
            }
        }

        private string FormatTimeInput(string input)
        {
            // Remove any non-digit characters
            string digits = new string(input.Where(char.IsDigit).ToArray());

            if (string.IsNullOrEmpty(digits))
                return null;

            // Pad with leading zeros if needed
            if (digits.Length == 1)
                digits = "0" + digits + "00"; // "1" -> "0100" -> "01:00"
            else if (digits.Length == 2)
                digits = digits + "00"; // "12" -> "1200" -> "12:00"
            else if (digits.Length == 3)
                digits = "0" + digits; // "721" -> "0721" -> "07:21"
            else if (digits.Length > 4)
                digits = digits.Substring(0, 4); // Take first 4 digits

            // Parse hours and minutes
            int hours = int.Parse(digits.Substring(0, 2));
            int minutes = int.Parse(digits.Substring(2, 2));

            // Validate - Allow broadcast times up to 29:59 (24:00 = midnight next day, 25:00 = 1am next day, etc.)
            if (hours > 29 || minutes > 59)
                return null;

            return $"{hours:D2}:{minutes:D2}";
        }

        private void DgvCalendar_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            string columnName = dgvCalendar.Columns[e.ColumnIndex].Name;

            // Display 0 as blank in date columns
            if (columnName.Contains("\n") && char.IsDigit(columnName.Split('\n')[1][0]))
            {
                if (e.Value != null && e.Value is int intValue && intValue == 0)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }

                // Apply alternating colors to date columns
                int dayNumber = int.Parse(columnName.Split('\n')[1]);
                if (dayNumber % 2 == 0)
                {
                    e.CellStyle.BackColor = System.Drawing.Color.FromArgb(245, 250, 255); // Light blue
                }
                else
                {
                    e.CellStyle.BackColor = System.Drawing.Color.White;
                }
            }

            // Total row styling
            if (e.RowIndex == calendarDataTable.Rows.Count - 1)
            {
                e.CellStyle.BackColor = System.Drawing.Color.FromArgb(220, 220, 220); // Darker gray
                e.CellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
                e.CellStyle.ForeColor = System.Drawing.Color.FromArgb(50, 50, 50);
            }
        }

        private void ConfigureGridAppearance()
        {
            // General settings
            dgvCalendar.AllowUserToAddRows = false;
            dgvCalendar.AllowUserToDeleteRows = false;
            dgvCalendar.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgvCalendar.MultiSelect = true;
            dgvCalendar.ReadOnly = false;
            dgvCalendar.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            dgvCalendar.ColumnHeadersHeight = 40; // Increased height for two-line headers
            dgvCalendar.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dgvCalendar.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvCalendar.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Enable programmatic sorting - clicking column header triggers custom sort handler
            foreach (DataGridViewColumn column in dgvCalendar.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
            dgvCalendar.ColumnHeaderMouseClick -= DgvCalendar_ColumnHeaderMouseClick;
            dgvCalendar.ColumnHeaderMouseClick += DgvCalendar_ColumnHeaderMouseClick;

            // Hide Spot Index column
            if (dgvCalendar.Columns["Spot Index"] != null)
            {
                dgvCalendar.Columns["Spot Index"].Visible = false;
            }

            // Convert specific columns to ComboBox columns with autocomplete
            // Note: "Programs Name" now uses AutoComplete overlay instead of ComboBox
            // ConvertToComboBoxColumn("Programs Name", GetProgrammeNameList());
            ConvertToComboBoxColumn("Break In", GetBreakInList());
            ConvertToComboBoxColumn("F/P", GetFPList());

            // Make Programs Name column editable (ReadOnly = false is default, but setting explicitly)
            if (dgvCalendar.Columns["Programs Name"] != null)
            {
                dgvCalendar.Columns["Programs Name"].ReadOnly = true; // We handle editing via AutoComplete overlay
            }

            // Make other metadata columns editable but with light background to distinguish them
            string[] otherMetadataColumns = new string[]
            {
                "Time (KWT)", "Ratio", "Sponsor Type", "ORD", "Sponsor Price", "OID", "Unit Price KWD", "Price in US $"
            };

            foreach (string colName in otherMetadataColumns)
            {
                if (dgvCalendar.Columns[colName] != null)
                {
                    dgvCalendar.Columns[colName].ReadOnly = false;
                    dgvCalendar.Columns[colName].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(250, 250, 250);
                }
            }

            // Make Total Spots column read-only
            if (dgvCalendar.Columns["Total Spots"] != null)
            {
                dgvCalendar.Columns["Total Spots"].ReadOnly = true;
                dgvCalendar.Columns["Total Spots"].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                dgvCalendar.Columns["Total Spots"].DefaultCellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
            }

            // Date columns are editable
            int daysInMonth = DateTime.DaysInMonth(year, month);
            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";

                if (dgvCalendar.Columns[columnName] != null)
                {
                    dgvCalendar.Columns[columnName].Width = 35;
                    dgvCalendar.Columns[columnName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }

            // Style the Total row
            if (dgvCalendar.Rows.Count > 0)
            {
                var lastRow = dgvCalendar.Rows[dgvCalendar.Rows.Count - 1];
                lastRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                lastRow.DefaultCellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
                lastRow.ReadOnly = true;
            }

            lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}"; // -1 for total row
        }

        private void DgvCalendar_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (calendarDataTable == null || calendarDataTable.Rows.Count <= 1)
                return;

            var clickedColumn = dgvCalendar.Columns[e.ColumnIndex];

            // Skip if column doesn't support sorting
            if (clickedColumn.SortMode == DataGridViewColumnSortMode.NotSortable)
                return;

            // Get the column to sort by
            string columnName = clickedColumn.Name;

            // Determine sort direction (toggle)
            ListSortDirection direction = ListSortDirection.Ascending;
            if (clickedColumn.HeaderCell.SortGlyphDirection == SortOrder.Ascending)
            {
                direction = ListSortDirection.Descending;
            }

            // Step 1: Save and remove Total row (always the last row)
            int totalRowIndex = calendarDataTable.Rows.Count - 1;
            object[] totalRowData = calendarDataTable.Rows[totalRowIndex].ItemArray.Clone() as object[];
            calendarDataTable.Rows.RemoveAt(totalRowIndex);

            // Step 2: Sort the data rows using DataView
            DataView dv = calendarDataTable.DefaultView;
            dv.Sort = columnName + (direction == ListSortDirection.Descending ? " DESC" : " ASC");

            // Step 3: Rebuild table with sorted data
            DataTable sortedTable = dv.ToTable();
            calendarDataTable.Rows.Clear();
            foreach (DataRow row in sortedTable.Rows)
            {
                calendarDataTable.ImportRow(row);
            }

            // Step 4: Add Total row back at the bottom
            DataRow newTotalRow = calendarDataTable.NewRow();
            newTotalRow.ItemArray = totalRowData;
            calendarDataTable.Rows.Add(newTotalRow);

            // Update sort glyph (only for Programmatic columns)
            foreach (DataGridViewColumn col in dgvCalendar.Columns)
            {
                if (col.SortMode == DataGridViewColumnSortMode.Programmatic)
                {
                    col.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
            if (clickedColumn.SortMode == DataGridViewColumnSortMode.Programmatic)
            {
                clickedColumn.HeaderCell.SortGlyphDirection =
                    direction == ListSortDirection.Ascending ? SortOrder.Ascending : SortOrder.Descending;
            }

            // Reapply Total row styling (last row after sort)
            StyleTotalRow();
        }

        /// <summary>
        /// Applies consistent styling to the Total row (always the last row)
        /// </summary>
        private void StyleTotalRow()
        {
            if (dgvCalendar.Rows.Count > 0)
            {
                var lastRow = dgvCalendar.Rows[dgvCalendar.Rows.Count - 1];
                lastRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                lastRow.DefaultCellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
                lastRow.ReadOnly = true;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                // Reconstruct booking order from calendar grid
                UpdateBookingOrderFromCalendar();

                // Save to JSON file
                string jsonContent = JsonConvert.SerializeObject(bookingOrder, Formatting.Indented);
                File.WriteAllText(jsonFilePath, jsonContent);

                MessageBox.Show($"Changes saved to:\n{jsonFilePath}\n\nDatabase integration coming soon.",
                    "Save Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving changes: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void UpdateBookingOrderFromCalendar()
        {
            bookingOrder.Spots.Clear();
            int daysInMonth = DateTime.DaysInMonth(year, month);

            // Skip the last row (Total row)
            for (int i = 0; i < calendarDataTable.Rows.Count - 1; i++)
            {
                DataRow row = calendarDataTable.Rows[i];

                string programmeName = row["Programs Name"].ToString();
                string programmeTime = row["Time (KWT)"].ToString();

                // Collect dates based on spot counts in calendar cells
                List<string> dates = new List<string>();
                for (int day = 1; day <= daysInMonth; day++)
                {
                    var date = new DateTime(year, month, day);
                    string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";

                    int spotCount = Convert.ToInt32(row[columnName]);
                    for (int j = 0; j < spotCount; j++)
                    {
                        dates.Add(date.ToString("yyyy-MM-dd"));
                    }
                }

                if (dates.Count > 0)
                {
                    var spot = new Spot
                    {
                        ProgrammeName = programmeName,
                        ProgrammeStartTime = programmeTime,
                        Duration = "", // Can be enhanced later
                        Dates = dates,
                        TotalSpots = dates.Count
                    };
                    bookingOrder.Spots.Add(spot);
                }
            }

            bookingOrder.TotalSpots = bookingOrder.Spots.Sum(s => s.TotalSpots);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Any unsaved changes will be lost. Do you want to go back?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                MainMenuForm mainForm = new MainMenuForm();
                mainForm.Show();
                this.Close();
            }
        }

        private void dgvCalendar_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Skip if we're in batch update mode
            if (isUpdating || e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            // Check if it's a date column that changed
            string columnName = dgvCalendar.Columns[e.ColumnIndex].Name;
            if (columnName.Contains(" ") && char.IsDigit(columnName.Split(' ')[1][0]))
            {
                // Recalculate Total Spots for this row
                RecalculateRowTotal(e.RowIndex);

                // Recalculate column totals in the Total row
                RecalculateColumnTotals();
            }
        }

        private void RecalculateRowTotal(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= calendarDataTable.Rows.Count - 1)
                return;

            DataRow row = calendarDataTable.Rows[rowIndex];
            int total = 0;
            int daysInMonth = DateTime.DaysInMonth(year, month);

            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";

                try
                {
                    total += Convert.ToInt32(row[columnName]);
                }
                catch
                {
                    row[columnName] = 0;
                }
            }

            row["Total Spots"] = total;
        }

        private void RecalculateColumnTotals()
        {
            int daysInMonth = DateTime.DaysInMonth(year, month);
            DataRow totalRow = calendarDataTable.Rows[calendarDataTable.Rows.Count - 1];
            int grandTotal = 0;

            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
                string columnName = $"{date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2)} {day}";

                int dayTotal = 0;
                for (int i = 0; i < calendarDataTable.Rows.Count - 1; i++)
                {
                    try
                    {
                        dayTotal += Convert.ToInt32(calendarDataTable.Rows[i][columnName]);
                    }
                    catch { }
                }
                totalRow[columnName] = dayTotal;
                grandTotal += dayTotal;
            }

            totalRow["Total Spots"] = grandTotal;

            // Update header label
            bookingOrder.TotalSpots = grandTotal;
            lblTotalSpots.Text = $"Total Spots: {grandTotal}";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Cleanup AutoComplete control
            programmeAutoComplete?.Dispose();

            base.OnFormClosing(e);
            Application.Exit();
        }
    }
}
