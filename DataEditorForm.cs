using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace CSharpFlexGrid
{
    public partial class DataEditorForm : Form
    {
        private BookingOrder bookingOrder;
        private string jsonFilePath;
        private DataTable dataTable;

        public DataEditorForm(BookingOrder order, string filePath)
        {
            InitializeComponent();
            bookingOrder = order;
            jsonFilePath = filePath;
            LoadBookingOrderData();
            UpdateHeaderLabels();
        }

        private void UpdateHeaderLabels()
        {
            lblAgency.Text = $"Agency: {bookingOrder.Agency}";
            lblAdvertiser.Text = $"Advertiser: {bookingOrder.Advertiser}";
            lblProduct.Text = $"Product: {bookingOrder.Product}";
            lblCampaign.Text = $"Campaign: {bookingOrder.CampaignPeriod.StartDate} to {bookingOrder.CampaignPeriod.EndDate}";
            lblTotalSpots.Text = $"Total Spots: {bookingOrder.TotalSpots}";
        }

        private void LoadBookingOrderData()
        {
            // Create DataTable with columns
            dataTable = new DataTable();
            dataTable.Columns.Add("Programme", typeof(string));
            dataTable.Columns.Add("Time", typeof(string));
            dataTable.Columns.Add("Duration", typeof(string));
            dataTable.Columns.Add("Date", typeof(string));
            dataTable.Columns.Add("Spot Index", typeof(int));

            // Flatten the booking order data into rows
            foreach (var spot in bookingOrder.Spots)
            {
                for (int i = 0; i < spot.Dates.Count; i++)
                {
                    DataRow row = dataTable.NewRow();
                    row["Programme"] = spot.ProgrammeName ?? "";
                    row["Time"] = spot.ProgrammeStartTime ?? "";
                    row["Duration"] = spot.Duration ?? "";
                    row["Date"] = spot.Dates[i];
                    row["Spot Index"] = i + 1;
                    dataTable.Rows.Add(row);
                }
            }

            // Bind to DataGridView
            dgvSpots.DataSource = dataTable;

            // Configure grid appearance
            dgvSpots.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvSpots.AllowUserToAddRows = false;
            dgvSpots.AllowUserToDeleteRows = false;
            dgvSpots.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvSpots.MultiSelect = false;
            dgvSpots.ReadOnly = false;

            // Make Spot Index column read-only
            dgvSpots.Columns["Spot Index"].ReadOnly = true;
            dgvSpots.Columns["Spot Index"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;

            lblRecordCount.Text = $"Total Records: {dataTable.Rows.Count}";
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // For now, just save back to JSON
                // Database integration will be added later
                Cursor = Cursors.WaitCursor;

                // Update the booking order object with edited data
                UpdateBookingOrderFromGrid();

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

        private void UpdateBookingOrderFromGrid()
        {
            // Reconstruct spots from the edited grid data
            var spotGroups = new Dictionary<string, List<DataRow>>();

            // Group rows by programme, time, and duration
            foreach (DataRow row in dataTable.Rows)
            {
                string key = $"{row["Programme"]}|{row["Time"]}|{row["Duration"]}";
                if (!spotGroups.ContainsKey(key))
                {
                    spotGroups[key] = new List<DataRow>();
                }
                spotGroups[key].Add(row);
            }

            // Rebuild spots list
            bookingOrder.Spots.Clear();
            int totalSpots = 0;

            foreach (var group in spotGroups)
            {
                var firstRow = group.Value.First();
                var spot = new Spot
                {
                    ProgrammeName = firstRow["Programme"].ToString(),
                    ProgrammeStartTime = firstRow["Time"].ToString(),
                    Duration = firstRow["Duration"].ToString(),
                    Dates = group.Value.Select(r => r["Date"].ToString()).ToList(),
                    TotalSpots = group.Value.Count
                };
                bookingOrder.Spots.Add(spot);
                totalSpots += spot.TotalSpots;
            }

            // Update total spots
            bookingOrder.TotalSpots = totalSpots;
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

        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "JSON Files|*.json|All Files|*.*";
                saveFileDialog.Title = "Export to JSON";
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(jsonFilePath) + "_edited.json";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        UpdateBookingOrderFromGrid();
                        string jsonContent = JsonConvert.SerializeObject(bookingOrder, Formatting.Indented);
                        File.WriteAllText(saveFileDialog.FileName, jsonContent);

                        MessageBox.Show($"Data exported successfully to:\n{saveFileDialog.FileName}",
                            "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting data: {ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                dataTable.DefaultView.RowFilter = string.Empty;
            }
            else
            {
                string searchText = txtSearch.Text.Replace("'", "''");
                dataTable.DefaultView.RowFilter = $"Programme LIKE '%{searchText}%' OR " +
                                                   $"Time LIKE '%{searchText}%' OR " +
                                                   $"Date LIKE '%{searchText}%'";
            }
            lblRecordCount.Text = $"Showing: {dgvSpots.Rows.Count} / {dataTable.Rows.Count} Records";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            // Exit the application when editor is closed
            Application.Exit();
        }
    }
}
