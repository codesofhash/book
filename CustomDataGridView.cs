using System;
using System.Windows.Forms;

namespace CSharpFlexGrid
{
    public class CustomDataGridView : DataGridView
    {
        public CustomDataGridView()
        {
            this.AllowUserToAddRows = false;
            this.AllowUserToDeleteRows = false;
            this.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }

        public void Copy()
        {
            DataObject dataObj = this.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        public void Cut()
        {
            Copy();
            foreach (DataGridViewCell cell in this.SelectedCells)
            {
                cell.Value = "";
            }
        }

        public void Paste()
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text))
                return;

            string[] lines = text.TrimEnd('\r', '\n').Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            if (this.CurrentCell == null)
                return;

            int startRow = this.CurrentCell.RowIndex;
            int startCol = this.CurrentCell.ColumnIndex;

            for (int i = 0; i < lines.Length; i++)
            {
                if (startRow + i >= this.RowCount)
                    break;

                string[] cells = lines[i].Split('\t');
                for (int j = 0; j < cells.Length; j++)
                {
                    if (startCol + j >= this.ColumnCount)
                        break;

                    this.Rows[startRow + i].Cells[startCol + j].Value = cells[j];
                }
            }
        }

        public void InsertRow()
        {
            if (this.CurrentCell != null)
            {
                this.Rows.Insert(this.CurrentCell.RowIndex, 1);
            }
        }

        public void InsertColumn()
        {
            if (this.CurrentCell != null)
            {
                DataGridViewColumn newCol = new DataGridViewTextBoxColumn();
                newCol.HeaderText = "New Column";
                this.Columns.Insert(this.CurrentCell.ColumnIndex, newCol);
            }
        }

        public void DeleteRow()
        {
            foreach (DataGridViewRow row in this.SelectedRows)
            {
                if (!row.IsNewRow)
                {
                    this.Rows.Remove(row);
                    
                }
            }
        }

        public void DeleteColumn()
        {
            foreach (DataGridViewColumn column in this.SelectedColumns)
            {
                this.Columns.Remove(column);
            }
        }

        public void SetColumnName(int columnIndex, string name)
        {
            if (columnIndex >= 0 && columnIndex < this.ColumnCount)
            {
                this.Columns[columnIndex].HeaderText = name;
            }
        }

        public void SetColumnWidth(int columnIndex, int width)
        {
            if (columnIndex >= 0 && columnIndex < this.ColumnCount)
            {
                this.Columns[columnIndex].Width = width;
            }
        }

        public void SetRowHeight(int rowIndex, int height)
        {
            if (rowIndex >= 0 && rowIndex < this.RowCount)
            {
                this.Rows[rowIndex].Height = height;
            }
        }
    }
}