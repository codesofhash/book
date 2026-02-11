using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DotNetEnv;
using Newtonsoft.Json;

namespace CSharpFlexGrid
{
    public partial class BookingManagerForm : Form, IMessageFilter
    {
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private BookingOrder currentBookingOrder;
        private string currentJsonFilePath;
        private DataTable calendarDataTable;
        private DateTime campaignStartDate;
        private DateTime campaignEndDate;
        private bool isUpdating = false;
        private string jsonOutputPath;

        // Grid Mode controls
        private GroupBox grpGridMode;
        private RadioButton rbCampaignDates;
        private RadioButton rbSpecificDate;
        private DateTimePicker dtpSpecificDate;
        private Label lblSpecificDate;
        private string defaultGridMode = "CampaignDates";
        private int dealSearchDays = 15;
        private int currentDealNumber = 0;

        // Autocomplete suggestion list for Time column
        private ListBox suggestionListBox;
        private Control currentEditingControl;
        private string currentEditingColumn;

        public BookingManagerForm()
        {
            InitializeComponent();
            LoadSettings();  // Load settings first (includes JsonSavePath)
            LoadEnvironmentVariables();  // Then load .env for database credentials
            InitializeGridModeControls();
            InitializeSuggestionListBox();
            InitializeAutoCompleteData();
            SetupEventHandlers();

            // Register message filter to intercept keys at application level
            Application.AddMessageFilter(this);

            // Ensure form respects taskbar when maximized
            this.Load += (s, e) =>
            {
                var workingArea = Screen.PrimaryScreen.WorkingArea;
                this.MaximizedBounds = workingArea;
            };

            // Unregister filter when form closes
            this.FormClosed += (s, e) =>
            {
                Application.RemoveMessageFilter(this);
            };
        }

        // IMessageFilter implementation - intercepts keys before any control processes them
        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_KEYDOWN && suggestionListBox != null && suggestionListBox.Visible)
            {
                Keys keyCode = (Keys)m.WParam.ToInt32();

                if (keyCode == Keys.Down)
                {
                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (suggestionListBox.SelectedIndex < suggestionListBox.Items.Count - 1)
                            suggestionListBox.SelectedIndex++;
                        else
                            suggestionListBox.SelectedIndex = 0;
                    }
                    return true; // Message handled, don't process further
                }
                else if (keyCode == Keys.Up)
                {
                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (suggestionListBox.SelectedIndex > 0)
                            suggestionListBox.SelectedIndex--;
                        else
                            suggestionListBox.SelectedIndex = suggestionListBox.Items.Count - 1;
                    }
                    return true; // Message handled
                }
                else if (keyCode == Keys.Enter)
                {
                    if (suggestionListBox.SelectedIndex >= 0)
                    {
                        SelectSuggestion();
                    }
                    return true; // Message handled
                }
                else if (keyCode == Keys.Escape)
                {
                    HideSuggestionList();
                    return true; // Message handled
                }
                else if (keyCode == Keys.Tab)
                {
                    if (suggestionListBox.SelectedIndex >= 0)
                    {
                        SelectSuggestion();
                    }
                    return true; // Message handled
                }
            }

            return false; // Let message pass through
        }

        private void LoadSettings()
        {
            try
            {
                string settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");
                if (File.Exists(settingsPath))
                {
                    string json = File.ReadAllText(settingsPath);
                    var settings = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

                    if (settings.ContainsKey("DefaultGridMode"))
                        defaultGridMode = settings["DefaultGridMode"]?.ToString() ?? "CampaignDates";

                    if (settings.ContainsKey("DealSearchDays"))
                        int.TryParse(settings["DealSearchDays"]?.ToString(), out dealSearchDays);

                    // Load JSON save path from settings
                    if (settings.ContainsKey("JsonSavePath"))
                    {
                        string savePath = settings["JsonSavePath"]?.ToString();
                        if (!string.IsNullOrEmpty(savePath))
                        {
                            jsonOutputPath = savePath;
                            if (!Directory.Exists(jsonOutputPath))
                            {
                                Directory.CreateDirectory(jsonOutputPath);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                defaultGridMode = "CampaignDates";
                dealSearchDays = 15;
            }
        }

        private void InitializeSuggestionListBox()
        {
            suggestionListBox = new ListBox
            {
                Visible = false,
                Font = new System.Drawing.Font("Segoe UI", 9F),
                BorderStyle = BorderStyle.FixedSingle,
                SelectionMode = SelectionMode.One,
                IntegralHeight = false
            };

            suggestionListBox.Click += SuggestionListBox_Click;
            suggestionListBox.KeyDown += SuggestionListBox_KeyDown;
            suggestionListBox.Leave += SuggestionListBox_Leave;

            this.Controls.Add(suggestionListBox);
            suggestionListBox.BringToFront();
        }

        // Override ProcessCmdKey to intercept navigation keys before DataGridView handles them
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (suggestionListBox != null && suggestionListBox.Visible)
            {
                if (keyData == Keys.Down)
                {
                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (suggestionListBox.SelectedIndex < suggestionListBox.Items.Count - 1)
                            suggestionListBox.SelectedIndex++;
                        else
                            suggestionListBox.SelectedIndex = 0;
                    }
                    return true; // Key handled, don't process further
                }
                else if (keyData == Keys.Up)
                {
                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (suggestionListBox.SelectedIndex > 0)
                            suggestionListBox.SelectedIndex--;
                        else
                            suggestionListBox.SelectedIndex = suggestionListBox.Items.Count - 1;
                    }
                    return true; // Key handled, don't process further
                }
                else if (keyData == Keys.Enter)
                {
                    if (suggestionListBox.SelectedIndex >= 0)
                    {
                        SelectSuggestion();
                    }
                    return true; // Key handled, don't process further
                }
                else if (keyData == Keys.Escape)
                {
                    HideSuggestionList();
                    return true; // Key handled, don't process further
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void SuggestionListBox_Click(object sender, EventArgs e)
        {
            SelectSuggestion();
        }

        private void SuggestionListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SelectSuggestion();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                HideSuggestionList();
                e.Handled = true;
            }
        }

        private void SuggestionListBox_Leave(object sender, EventArgs e)
        {
            // Delay hiding to allow click selection
            BeginInvoke(new Action(() =>
            {
                if (!suggestionListBox.Focused && currentEditingControl != null && !currentEditingControl.Focused)
                {
                    HideSuggestionList();
                }
            }));
        }

        private void SelectSuggestion()
        {
            if (suggestionListBox.SelectedItem != null && currentEditingControl != null)
            {
                string selectedValue = suggestionListBox.SelectedItem.ToString();

                if (currentEditingControl is TextBox textBox)
                {
                    textBox.Text = selectedValue;
                    textBox.SelectionStart = textBox.Text.Length;
                }
                else if (currentEditingControl is DataGridViewComboBoxEditingControl comboBox)
                {
                    // Add to items if not exists
                    if (!comboBox.Items.Contains(selectedValue))
                    {
                        comboBox.Items.Add(selectedValue);
                    }
                    comboBox.Text = selectedValue;
                }

                HideSuggestionList();
                currentEditingControl.Focus();
            }
        }

        private void HideSuggestionList()
        {
            suggestionListBox.Visible = false;
            suggestionListBox.Items.Clear();
        }

        private void ShowSuggestionList(Control editingControl, string columnName, string searchText)
        {
            List<string> suggestions = new List<string>();

            if (columnName == "Time")
            {
                suggestions = GetTimeSuggestions(searchText);
            }
            else if (columnName == "Programme")
            {
                string timeValue = "";
                if (dgvCalendar.CurrentCell != null)
                {
                    timeValue = dgvCalendar["Time", dgvCalendar.CurrentCell.RowIndex].Value?.ToString() ?? "";
                }
                suggestions = GetProgrammeSuggestions(searchText, timeValue);
            }
            else if (columnName == "Sales Type")
            {
                suggestions = GetBreakInList()
                    .Where(s => s.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
            }

            if (suggestions.Count == 0)
            {
                HideSuggestionList();
                return;
            }

            // Populate the listbox
            suggestionListBox.Items.Clear();
            foreach (var suggestion in suggestions)
            {
                suggestionListBox.Items.Add(suggestion);
            }

            // Position the listbox below the cell
            PositionSuggestionList();

            // Set height based on item count (max 8 items visible)
            int itemHeight = suggestionListBox.ItemHeight;
            int visibleItems = Math.Min(suggestions.Count, 8);
            suggestionListBox.Height = (itemHeight * visibleItems) + 4;

            // Set initial selection to first item
            if (suggestionListBox.Items.Count > 0)
            {
                suggestionListBox.SelectedIndex = 0;
            }

            suggestionListBox.Visible = true;
            suggestionListBox.BringToFront();
        }

        private void PositionSuggestionList()
        {
            if (dgvCalendar.CurrentCell == null)
                return;

            // Get the cell rectangle relative to the DataGridView
            var cellRect = dgvCalendar.GetCellDisplayRectangle(
                dgvCalendar.CurrentCell.ColumnIndex,
                dgvCalendar.CurrentCell.RowIndex,
                false);

            // Convert to form coordinates
            var cellLocation = dgvCalendar.PointToScreen(new System.Drawing.Point(cellRect.Left, cellRect.Bottom));
            var formLocation = this.PointToClient(cellLocation);

            suggestionListBox.Left = formLocation.X;
            suggestionListBox.Top = formLocation.Y;
            suggestionListBox.Width = cellRect.Width;
        }

        private List<string> GetTimeSuggestions(string searchText)
        {
            var db = new DatabaseHelper();

            string startDate = campaignStartDate.ToString("yyyy-MM-dd");
            string endDate = campaignEndDate.ToString("yyyy-MM-dd");

            // Escape searchText for SQL
            string escapedSearch = searchText?.Replace("'", "''") ?? "";

            string query = $"SELECT DISTINCT start_at FROM u333577897_dbofhash.grid " +
                $"WHERE trans_dt >= '{startDate}' AND trans_dt <= '{endDate}' ";

            if (!string.IsNullOrWhiteSpace(escapedSearch))
            {
                query += $"AND start_at LIKE '{escapedSearch}%' ";
            }

            query += "ORDER BY start_at LIMIT 20";

            var times = db.GetComboBoxList(query);

            // Format times to HH:mm
            return times.Select(t =>
            {
                if (t.Length >= 5)
                    return t.Substring(0, 5);
                return t;
            }).Distinct().ToList();
        }

        private List<string> GetProgrammeSuggestions(string searchText, string timeValue)
        {
            var db = new DatabaseHelper();

            string startDate = campaignStartDate.ToString("yyyy-MM-dd");
            string endDate = campaignEndDate.ToString("yyyy-MM-dd");

            // Escape inputs for SQL
            string escapedSearch = searchText?.Replace("'", "''") ?? "";
            string escapedTime = timeValue?.Replace("'", "''") ?? "";

            // Format time to HH:mm:ss if needed
            if (!string.IsNullOrEmpty(escapedTime) && escapedTime.Length == 5)
            {
                escapedTime = escapedTime + ":00";
            }

            string query = $"SELECT DISTINCT(prog_details.name_en) FROM u333577897_dbofhash.grid " +
                $"LEFT JOIN prog_details ON grid.tagid = prog_details.idprog_details " +
                $"WHERE start_at = '{escapedTime}' " +
                $"AND (trans_dt >= '{startDate}' AND trans_dt <= '{endDate}') ";

            if (!string.IsNullOrWhiteSpace(escapedSearch))
            {
                query += $"AND prog_details.name_en LIKE '%{escapedSearch}%' ";
            }

            query += "ORDER BY prog_details.name_en";

            return db.GetComboBoxList(query);
        }

        private void EditingControl_TextChanged(object sender, EventArgs e)
        {
            if (currentEditingControl == null || string.IsNullOrEmpty(currentEditingColumn))
                return;

            string searchText = "";

            if (sender is TextBox textBox)
            {
                searchText = textBox.Text;
            }
            else if (sender is DataGridViewComboBoxEditingControl comboBox)
            {
                searchText = comboBox.Text;
            }

            // For Programme column, show all suggestions even when empty
            // For other columns, require at least 1 character
            if (currentEditingColumn == "Programme" || searchText.Length >= 1)
            {
                ShowSuggestionList(currentEditingControl, currentEditingColumn, searchText);
            }
            else
            {
                HideSuggestionList();
            }
        }

        private void EditingControl_TextUpdate(object sender, EventArgs e)
        {
            if (currentEditingControl == null || string.IsNullOrEmpty(currentEditingColumn))
                return;

            string searchText = "";

            if (sender is DataGridViewComboBoxEditingControl comboBox)
            {
                searchText = comboBox.Text;
            }

            // For Programme column, show all suggestions even when empty
            // For other columns, require at least 1 character
            if (currentEditingColumn == "Programme" || searchText.Length >= 1)
            {
                ShowSuggestionList(currentEditingControl, currentEditingColumn, searchText);
            }
            else
            {
                HideSuggestionList();
            }
        }

        private void EditingControl_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            // Mark arrow keys as input keys when suggestion list is visible
            // This prevents DataGridView from handling them (moving to next cell)
            if (suggestionListBox.Visible)
            {
                if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up ||
                    e.KeyCode == Keys.Enter || e.KeyCode == Keys.Escape)
                {
                    e.IsInputKey = true;
                }
            }
        }

        private void EditingControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (suggestionListBox.Visible)
            {
                // Close ComboBox dropdown if open to prevent interference
                if (sender is DataGridViewComboBoxEditingControl comboBox && comboBox.DroppedDown)
                {
                    comboBox.DroppedDown = false;
                }

                if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;

                    // Give focus to the suggestion list so it can handle navigation
                    suggestionListBox.Focus();

                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (e.KeyCode == Keys.Down)
                        {
                            if (suggestionListBox.SelectedIndex < suggestionListBox.Items.Count - 1)
                                suggestionListBox.SelectedIndex++;
                            else
                                suggestionListBox.SelectedIndex = 0;
                        }
                        else // Up
                        {
                            if (suggestionListBox.SelectedIndex > 0)
                                suggestionListBox.SelectedIndex--;
                            else
                                suggestionListBox.SelectedIndex = suggestionListBox.Items.Count - 1;
                        }
                    }
                    return;
                }
                else if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
                {
                    if (suggestionListBox.SelectedIndex >= 0)
                    {
                        SelectSuggestion();
                    }
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    HideSuggestionList();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void EditingControl_LostFocus(object sender, EventArgs e)
        {
            // Delay to allow click on suggestion list
            BeginInvoke(new Action(() =>
            {
                if (!suggestionListBox.Focused)
                {
                    HideSuggestionList();
                }
            }));
        }

        private void InitializeGridModeControls()
        {
            // Create GroupBox for Grid Mode
            grpGridMode = new GroupBox
            {
                Text = "Grid Mode",
                Location = new System.Drawing.Point(10, 5),
                Size = new System.Drawing.Size(450, 45),
                Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
            };

            // Radio Button 1 - Campaign Dates
            rbCampaignDates = new RadioButton
            {
                Text = "Grid Based on Campaign Dates",
                Location = new System.Drawing.Point(15, 18),
                AutoSize = true,
                Checked = (defaultGridMode == "CampaignDates"),
                Font = new System.Drawing.Font("Segoe UI", 9F)
            };
            rbCampaignDates.CheckedChanged += RbGridMode_CheckedChanged;

            // Radio Button 2 - Specific Date
            rbSpecificDate = new RadioButton
            {
                Text = "Grid Based on Specific Date",
                Location = new System.Drawing.Point(240, 18),
                AutoSize = true,
                Checked = (defaultGridMode == "SpecificDate"),
                Font = new System.Drawing.Font("Segoe UI", 9F)
            };
            rbSpecificDate.CheckedChanged += RbGridMode_CheckedChanged;

            // Label for Specific Date
            lblSpecificDate = new Label
            {
                Text = "Select Date:",
                Location = new System.Drawing.Point(480, 18),
                AutoSize = true,
                Visible = (defaultGridMode == "SpecificDate"),
                Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
            };

            // DateTimePicker for Specific Date
            dtpSpecificDate = new DateTimePicker
            {
                Location = new System.Drawing.Point(570, 14),
                Size = new System.Drawing.Size(150, 25),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today,
                Visible = (defaultGridMode == "SpecificDate")
            };

            grpGridMode.Controls.Add(rbCampaignDates);
            grpGridMode.Controls.Add(rbSpecificDate);

            // Add controls to panelGridMode
            panelGridMode.Controls.Add(grpGridMode);
            panelGridMode.Controls.Add(lblSpecificDate);
            panelGridMode.Controls.Add(dtpSpecificDate);
        }

        private void RbGridMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbSpecificDate.Checked)
            {
                lblSpecificDate.Visible = true;
                dtpSpecificDate.Visible = true;
                SetOIDColumnEnabled(true);
            }
            else
            {
                lblSpecificDate.Visible = false;
                dtpSpecificDate.Visible = false;
                SetOIDColumnEnabled(false);
                ClearOIDCells();
            }
        }

        private void ClearOIDCells()
        {
            if (dgvCalendar == null || dgvCalendar.Columns["OID"] == null)
                return;

            int oidColumnIndex = dgvCalendar.Columns["OID"].Index;
            for (int i = 0; i < dgvCalendar.Rows.Count; i++)
            {
                dgvCalendar.Rows[i].Cells[oidColumnIndex].Value = "";
            }
        }

        private void SetOIDColumnEnabled(bool enabled)
        {
            if (dgvCalendar != null && dgvCalendar.Columns["OID"] != null)
            {
                dgvCalendar.Columns["OID"].ReadOnly = !enabled;
                if (enabled)
                {
                    dgvCalendar.Columns["OID"].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 255, 200);
                }
                else
                {
                    dgvCalendar.Columns["OID"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                }
            }
        }

        private string GetPaymentRatio(string programme, string time, DateTime startDate, DateTime endDate)
        {
            var db = new DatabaseHelper();

            // Format time to HH:mm:ss if needed
            string formattedTime = time;
            if (!string.IsNullOrEmpty(time) && time.Length == 5)
            {
                formattedTime = time + ":00";
            }

            // Escape apostrophes in programme name for SQL
            string escapedProgramme = programme?.Replace("'", "''") ?? "";

            // Format dates for SQL
            string startDateStr = startDate.ToString("yyyy-MM-dd");
            string endDateStr = endDate.ToString("yyyy-MM-dd");

            var ratioList = db.GetComboBoxList(
                $"SELECT DISTINCT payment_ratio FROM u333577897_dbofhash.grid " +
                $"LEFT JOIN prog_details ON grid.tagid = prog_details.idprog_details " +
                $"WHERE name_en = '{escapedProgramme}' AND start_at = '{formattedTime}' " +
                $"AND trans_dt >= '{startDateStr}' AND trans_dt <= '{endDateStr}' " +
               // $"AND payment_ratio IS NOT NULL AND payment_ratio != '' " +
                $"ORDER BY payment_ratio DESC LIMIT 1");

            return ratioList.Count > 0 ? ratioList[0] : "";
        }

        private (string Time, string Programme, string FP) GetProgrammeDetailsByOID(string oid, DateTime selectedDate)
        {
            var db = new DatabaseHelper();
            string formattedDate = selectedDate.ToString("yyyy-MM-dd");

            var timeList = db.GetComboBoxList(
                $"SELECT start_at FROM u333577897_dbofhash.grid " +
                $"LEFT JOIN prog_details ON grid.tagid = prog_details.idprog_details " +
                $"WHERE ordinday = '{oid}' AND trans_dt = '{formattedDate}'");

            var programmeList = db.GetComboBoxList(
                $"SELECT name_en FROM u333577897_dbofhash.grid " +
                $"LEFT JOIN prog_details ON grid.tagid = prog_details.idprog_details " +
                $"WHERE ordinday = '{oid}' AND trans_dt = '{formattedDate}'");

            var fpList = db.GetComboBoxList(
                $"SELECT payment_ratio FROM u333577897_dbofhash.grid " +
                $"LEFT JOIN prog_details ON grid.tagid = prog_details.idprog_details " +
                $"WHERE ordinday = '{oid}' AND trans_dt = '{formattedDate}'");

            string time = timeList.Count > 0 ? timeList[0] : "";
            string programme = programmeList.Count > 0 ? programmeList[0] : "";
            string fp = fpList.Count > 0 ? fpList[0] : "";

            return (time, programme, fp);
        }

        private string GetRatioByDur(int dur)
        {
            // Validate dur is between 0 and 215
            if (dur < 0 || dur > 215)
                return "";

            var db = new DatabaseHelper();
            var ratioList = db.GetComboBoxList(
                $"SELECT ratio FROM u333577897_dbofhash.rate_card WHERE dur = '{dur}'");

            return ratioList.Count > 0 ? ratioList[0] : "";
        }

        private void LoadEnvironmentVariables()
        {
            try
            {
                string envPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".env");
                if (File.Exists(envPath))
                {
                    Env.Load(envPath);
                    // Only use .env path if not already set from settings.json
                    if (string.IsNullOrEmpty(jsonOutputPath))
                    {
                        jsonOutputPath = Environment.GetEnvironmentVariable("JSON_OUTPUT_PATH");

                        if (!string.IsNullOrEmpty(jsonOutputPath) && !Directory.Exists(jsonOutputPath))
                        {
                            Directory.CreateDirectory(jsonOutputPath);
                        }
                    }
                }

                // Fallback to default if still not set
                if (string.IsNullOrEmpty(jsonOutputPath))
                {
                    jsonOutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProcessedOrders");
                    Directory.CreateDirectory(jsonOutputPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading configuration: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeAutoCompleteData()
        {
            // Load data from database for autocomplete
            var db = new DatabaseHelper();

            // Load Agency list
            var agencies = db.GetComboBoxList(
                "SELECT DISTINCT(agency) FROM u333577897_dbofhash.bookingdetails ORDER BY agency");
            cmbAgency.LoadAutoCompleteItems(agencies.Count > 0 ? agencies : new List<string> { "Media Agency A", "Creative Agency B", "Digital Agency C" });

            // Load Advertiser list
            var advertisers = db.GetComboBoxList(
                // "SELECT DISTINCT(Advertiser) FROM u333577897_dbofhash.s4sales ORDER BY Advertiser ");
                "SELECT DISTINCT(advertiser) FROM u333577897_dbofhash.bookingdetails ORDER BY advertiser ");
            cmbAdvertiser.LoadAutoCompleteItems(advertisers.Count > 0 ? advertisers : new List<string> { "Advertiser 1", "Advertiser 2", "Advertiser 3" });

            // Load Product list
            var products = db.GetComboBoxList(
                "SELECT DISTINCT(product) FROM u333577897_dbofhash.bookingdetails2 ORDER BY product");
            cmbProduct.LoadAutoCompleteItems(products.Count > 0 ? products : new List<string> { "Product A", "Product B", "Product C" });

            // Set default dates
            dtpCampaignStart.Value = DateTime.Today;
            dtpCampaignEnd.Value = DateTime.Today.AddMonths(1);
        }

        private void SetupEventHandlers()
        {
            dgvCalendar.CellValueChanged += DgvCalendar_CellValueChanged;
            dgvCalendar.CellBeginEdit += DgvCalendar_CellBeginEdit;
            dgvCalendar.CellEndEdit += DgvCalendar_CellEndEdit;
            dgvCalendar.CellFormatting += DgvCalendar_CellFormatting;
            dgvCalendar.EditingControlShowing += DgvCalendar_EditingControlShowing;
            dgvCalendar.DataError += DgvCalendar_DataError;
            dgvCalendar.CellValidating += DgvCalendar_CellValidating;
            dgvCalendar.PreviewKeyDown += DgvCalendar_PreviewKeyDown;
            dgvCalendar.KeyDown += DgvCalendar_KeyDown;

            // Campaign date change handlers
            dtpCampaignStart.ValueChanged += DtpCampaignStart_ValueChanged;
            dtpCampaignEnd.ValueChanged += DtpCampaignEnd_ValueChanged;

            // Package Cost cleanup and calculation trigger
            txtPackageCost.Leave += TxtPackageCost_Leave;
        }

        private void DgvCalendar_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            // Mark arrow keys as input keys when suggestion list is visible
            // This prevents DataGridView from moving to next cell
            if (suggestionListBox != null && suggestionListBox.Visible)
            {
                if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up ||
                    e.KeyCode == Keys.Enter || e.KeyCode == Keys.Escape)
                {
                    e.IsInputKey = true;
                }
            }
        }

        private string GetDateColumnName(DateTime date)
        {
            // Format: "DayOfWeek\nDay\nMonth" e.g., "Mo\n15\nJan"
            string dayOfWeek = date.ToString("ddd", CultureInfo.InvariantCulture).Substring(0, 2);
            string monthAbbr = date.ToString("MMM", CultureInfo.InvariantCulture);
            return $"{dayOfWeek}\n{date.Day}\n{monthAbbr}";
        }

        private DateTime? GetDateFromColumnName(string columnName)
        {
            // Parse column name format: "DayOfWeek\nDay\nMonth"
            if (!columnName.Contains("\n"))
                return null;

            string[] parts = columnName.Split('\n');
            if (parts.Length < 3)
                return null;

            if (!int.TryParse(parts[1], out int day))
                return null;

            // Find the date in the campaign range that matches
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                if (GetDateColumnName(date) == columnName)
                    return date;
            }

            return null;
        }

        private void TxtPackageCost_Leave(object sender, EventArgs e)
        {
            FormatPackageCost();
            RecalculatePricing();
        }

        private void FormatPackageCost()
        {
            // Clean up Package Cost: remove all non-numeric characters except '.'
            string input = txtPackageCost.Text;
            if (string.IsNullOrWhiteSpace(input))
                return;

            // Keep only digits and decimal point
            string cleaned = new string(input.Where(c => char.IsDigit(c) || c == '.').ToArray());

            // Handle multiple decimal points - keep only the first one
            int firstDot = cleaned.IndexOf('.');
            if (firstDot >= 0)
            {
                string beforeDot = cleaned.Substring(0, firstDot);
                string afterDot = cleaned.Substring(firstDot + 1).Replace(".", "");
                cleaned = beforeDot + "." + afterDot;
            }

            // Parse and format to 3 decimal places
            if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal value))
            {
                // Format with 3 decimal places (e.g., 5000 -> 5000.000)
                txtPackageCost.Text = value.ToString("0.000");
            }
            else
            {
                txtPackageCost.Text = "";
            }
        }

        private void RecalculatePricing()
        {
            try
            {
                if (dgvCalendar == null || calendarDataTable == null || calendarDataTable.Rows.Count <= 1)
                    return;

                // Check if required columns exist
                if (!dgvCalendar.Columns.Contains("F/P") ||
                    !dgvCalendar.Columns.Contains("Ratio") ||
                    !dgvCalendar.Columns.Contains("Dur") ||
                    !dgvCalendar.Columns.Contains("Total Spots") ||
                    !dgvCalendar.Columns.Contains("Unit Price KWD"))
                    return;

                /* DEBUG: Uncomment to activate debug output
                var debug = new StringBuilder();
                debug.AppendLine("=== PRICING DEBUG ===\n");
                */

                // Get Package Cost
                decimal packageCost = 0;
                if (!string.IsNullOrWhiteSpace(txtPackageCost.Text))
                {
                    string cleaned = txtPackageCost.Text.Replace(",", "");
                    decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out packageCost);
                }
                // debug.AppendLine($"Package Cost: {packageCost}\n");

                // Calculate Total Paid Space (sum of all rows' Paid Space)
                // Paid Space per row = F/P * Ratio * Dur * Total Spots
                decimal totalPaidSpace = 0;

                // debug.AppendLine("--- PAID SPACE PER ROW ---");
                for (int i = 0; i < dgvCalendar.Rows.Count - 1; i++) // Exclude Total row
                {
                    decimal fp = 0, ratio = 0, dur = 0, totalSpots = 0;

                    var fpValue = dgvCalendar["F/P", i].Value?.ToString() ?? "";
                    var ratioValue = dgvCalendar["Ratio", i].Value?.ToString() ?? "";
                    var durValue = dgvCalendar["Dur", i].Value?.ToString() ?? "";
                    var spotsValue = dgvCalendar["Total Spots", i].Value?.ToString() ?? "";

                    decimal.TryParse(fpValue, NumberStyles.Any, CultureInfo.InvariantCulture, out fp);
                    decimal.TryParse(ratioValue, NumberStyles.Any, CultureInfo.InvariantCulture, out ratio);
                    decimal.TryParse(durValue, NumberStyles.Any, CultureInfo.InvariantCulture, out dur);
                    decimal.TryParse(spotsValue, NumberStyles.Any, CultureInfo.InvariantCulture, out totalSpots);

                    decimal paidSpace = fp * ratio * dur * totalSpots;
                    totalPaidSpace += paidSpace;

                    // debug.AppendLine($"Row {i}: F/P={fp}, Ratio={ratio}, Dur={dur}, TotalSpots={totalSpots}");
                    // debug.AppendLine($"        PaidSpace = {fp} × {ratio} × {dur} × {totalSpots} = {paidSpace}");
                }

                // debug.AppendLine($"\nTotal Paid Space: {totalPaidSpace}");

                // Calculate Average per second
                decimal avgPerSecond = 0;
                if (totalPaidSpace > 0)
                {
                    avgPerSecond = packageCost / totalPaidSpace;
                }
                // debug.AppendLine($"Average per second: {packageCost} / {totalPaidSpace} = {avgPerSecond}\n");

                // Calculate Unit Price KWD for each row
                // Unit Price KWD = F/P * Ratio * Dur * Average per second
                // debug.AppendLine("--- UNIT PRICE KWD PER ROW ---");
                for (int i = 0; i < dgvCalendar.Rows.Count - 1; i++) // Exclude Total row
                {
                    decimal fp = 0, ratio = 0, dur = 0;

                    var fpValue = dgvCalendar["F/P", i].Value?.ToString() ?? "";
                    var ratioValue = dgvCalendar["Ratio", i].Value?.ToString() ?? "";
                    var durValue = dgvCalendar["Dur", i].Value?.ToString() ?? "";

                    decimal.TryParse(fpValue, NumberStyles.Any, CultureInfo.InvariantCulture, out fp);
                    decimal.TryParse(ratioValue, NumberStyles.Any, CultureInfo.InvariantCulture, out ratio);
                    decimal.TryParse(durValue, NumberStyles.Any, CultureInfo.InvariantCulture, out dur);

                    decimal unitPriceKWD = fp * ratio * dur * avgPerSecond;

                    // debug.AppendLine($"Row {i}: UnitPriceKWD = {fp} × {ratio} × {dur} × {avgPerSecond:F6} = {unitPriceKWD}");

                    // Update the Unit Price KWD column
                    dgvCalendar["Unit Price KWD", i].Value = unitPriceKWD > 0 ? unitPriceKWD.ToString("N3") : "";
                }

                // DEBUG: Uncomment to show debug message
                // MessageBox.Show(debug.ToString(), "Pricing Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception)
            {
                // Silently ignore pricing calculation errors
            }
        }

        private void DtpCampaignStart_ValueChanged(object sender, EventArgs e)
        {
            if (currentBookingOrder == null || calendarDataTable == null)
                return;

            var newStartDate = dtpCampaignStart.Value.Date;
            var oldStartDate = campaignStartDate;

            // Calculate the offset (how many days the start date shifted)
            int dayOffset = (int)(newStartDate - oldStartDate).TotalDays;

            // Calculate new end date to maintain same campaign duration
            int campaignDuration = (int)(campaignEndDate - campaignStartDate).TotalDays;
            var newEndDate = newStartDate.AddDays(campaignDuration);

            // Update end date picker without triggering its event
            dtpCampaignEnd.ValueChanged -= DtpCampaignEnd_ValueChanged;
            dtpCampaignEnd.Value = newEndDate;
            dtpCampaignEnd.ValueChanged += DtpCampaignEnd_ValueChanged;

            // Update booking order
            currentBookingOrder.CampaignPeriod.StartDate = newStartDate.ToString("yyyy-MM-dd");
            currentBookingOrder.CampaignPeriod.EndDate = newEndDate.ToString("yyyy-MM-dd");

            // Rebuild grid with shifted spots
            RebuildGridWithShiftedSpots(newStartDate, newEndDate, dayOffset);
        }

        private void DtpCampaignEnd_ValueChanged(object sender, EventArgs e)
        {
            if (currentBookingOrder == null || calendarDataTable == null)
                return;

            var newStartDate = dtpCampaignStart.Value.Date;
            var newEndDate = dtpCampaignEnd.Value.Date;

            // Ensure end date is not before start date
            if (newEndDate < newStartDate)
            {
                newEndDate = newStartDate;
                dtpCampaignEnd.ValueChanged -= DtpCampaignEnd_ValueChanged;
                dtpCampaignEnd.Value = newEndDate;
                dtpCampaignEnd.ValueChanged += DtpCampaignEnd_ValueChanged;
            }

            // Update booking order
            currentBookingOrder.CampaignPeriod.StartDate = newStartDate.ToString("yyyy-MM-dd");
            currentBookingOrder.CampaignPeriod.EndDate = newEndDate.ToString("yyyy-MM-dd");

            // Rebuild grid for new date range (no shifting, just extend/shrink)
            RebuildGridForDateRange(newStartDate, newEndDate);
        }

        private void RebuildGridWithShiftedSpots(DateTime newStartDate, DateTime newEndDate, int dayOffset)
        {
            if (calendarDataTable == null)
                return;

            isUpdating = true;

            // Save current row data with spot values indexed by day offset from campaign start
            var savedRows = new List<Dictionary<string, object>>();

            for (int i = 0; i < calendarDataTable.Rows.Count - 1; i++) // Skip Total row
            {
                var rowData = new Dictionary<string, object>
                {
                    ["OID"] = calendarDataTable.Rows[i]["OID"],
                    ["Time"] = calendarDataTable.Rows[i]["Time"],
                    ["Programme"] = calendarDataTable.Rows[i]["Programme"],
                    ["F/P"] = calendarDataTable.Rows[i]["F/P"],
                    ["Dur"] = calendarDataTable.Rows[i]["Dur"],
                    ["Ratio"] = calendarDataTable.Rows[i]["Ratio"],
                    ["Sales Type"] = calendarDataTable.Rows[i]["Sales Type"],
                    ["ORD"] = calendarDataTable.Rows[i]["ORD"],
                    ["Sponsor Type"] = calendarDataTable.Rows[i]["Sponsor Type"],
                    ["Unit Price KWD"] = calendarDataTable.Rows[i]["Unit Price KWD"],
                    ["Price in US $"] = calendarDataTable.Rows[i]["Price in US $"]
                };

                // Save spot values by day index (0 = first day of campaign, 1 = second day, etc.)
                int dayIndex = 0;
                for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                {
                    string columnName = GetDateColumnName(date);
                    try
                    {
                        if (calendarDataTable.Columns.Contains(columnName))
                        {
                            rowData[$"DayIndex_{dayIndex}"] = Convert.ToInt32(calendarDataTable.Rows[i][columnName]);
                        }
                    }
                    catch
                    {
                        rowData[$"DayIndex_{dayIndex}"] = 0;
                    }
                    dayIndex++;
                }

                savedRows.Add(rowData);
            }

            // Update campaign dates
            campaignStartDate = newStartDate;
            campaignEndDate = newEndDate;

            // Rebuild the DataTable with new date columns
            calendarDataTable = new DataTable();

            // Add metadata columns
            calendarDataTable.Columns.Add("OID", typeof(string));
            calendarDataTable.Columns.Add("Time", typeof(string));
            calendarDataTable.Columns.Add("Programme", typeof(string));
            calendarDataTable.Columns.Add("F/P", typeof(string));
            calendarDataTable.Columns.Add("Dur", typeof(string));
            calendarDataTable.Columns.Add("Ratio", typeof(string));
            calendarDataTable.Columns.Add("Sales Type", typeof(string));
            calendarDataTable.Columns.Add("ORD", typeof(string));
            calendarDataTable.Columns.Add("Sponsor Type", typeof(string));
            calendarDataTable.Columns.Add("Unit Price KWD", typeof(string));
            calendarDataTable.Columns.Add("Price in US $", typeof(string));

            // Add new date columns for the campaign range
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);
                calendarDataTable.Columns.Add(columnName, typeof(int));
            }

            calendarDataTable.Columns.Add("Total Spots", typeof(int));

            // Restore saved rows with spots at same relative positions
            foreach (var rowData in savedRows)
            {
                DataRow newRow = calendarDataTable.NewRow();

                newRow["OID"] = rowData["OID"];
                newRow["Time"] = rowData["Time"];
                newRow["Programme"] = rowData["Programme"];
                newRow["F/P"] = rowData["F/P"];
                newRow["Dur"] = rowData["Dur"];
                newRow["Ratio"] = rowData["Ratio"];
                newRow["Sales Type"] = rowData["Sales Type"];
                newRow["ORD"] = rowData["ORD"];
                newRow["Sponsor Type"] = rowData["Sponsor Type"];
                newRow["Unit Price KWD"] = rowData["Unit Price KWD"];
                newRow["Price in US $"] = rowData["Price in US $"];

                int rowTotal = 0;
                int dayIndex = 0;
                for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                {
                    string columnName = GetDateColumnName(date);

                    // Restore spot value from same day index (relative position)
                    int spotValue = 0;
                    if (rowData.ContainsKey($"DayIndex_{dayIndex}"))
                    {
                        spotValue = Convert.ToInt32(rowData[$"DayIndex_{dayIndex}"]);
                    }
                    newRow[columnName] = spotValue;
                    rowTotal += spotValue;
                    dayIndex++;
                }

                newRow["Total Spots"] = rowTotal;
                calendarDataTable.Rows.Add(newRow);
            }

            // Add Total row
            DataRow totalRow = calendarDataTable.NewRow();
            totalRow["OID"] = "";
            totalRow["Time"] = "";
            totalRow["Programme"] = "Total";
            totalRow["F/P"] = "";
            totalRow["Dur"] = "";
            totalRow["Ratio"] = "";
            totalRow["Sales Type"] = "";
            totalRow["ORD"] = "";
            totalRow["Sponsor Type"] = "";
            totalRow["Unit Price KWD"] = "";
            totalRow["Price in US $"] = "";

            int grandTotal = 0;
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);

                int dayTotal = 0;
                for (int i = 0; i < calendarDataTable.Rows.Count; i++)
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
            calendarDataTable.Rows.Add(totalRow);

            // Rebind to grid
            dgvCalendar.DataSource = null;
            dgvCalendar.DataSource = calendarDataTable;

            // Reapply column formatting
            ApplyGridFormatting();

            isUpdating = false;

            // Update labels
            lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
            if (currentBookingOrder != null)
            {
                currentBookingOrder.TotalSpots = grandTotal;
                lblTotalSpots.Text = $"Total Spots: {grandTotal}";
            }
        }

        private void RebuildGridForDateRange(DateTime newStartDate, DateTime newEndDate)
        {
            if (calendarDataTable == null)
                return;

            isUpdating = true;

            // Save current row data (metadata and spot values by date)
            var savedRows = new List<Dictionary<string, object>>();

            for (int i = 0; i < calendarDataTable.Rows.Count - 1; i++) // Skip Total row
            {
                var rowData = new Dictionary<string, object>
                {
                    ["OID"] = calendarDataTable.Rows[i]["OID"],
                    ["Time"] = calendarDataTable.Rows[i]["Time"],
                    ["Programme"] = calendarDataTable.Rows[i]["Programme"],
                    ["F/P"] = calendarDataTable.Rows[i]["F/P"],
                    ["Dur"] = calendarDataTable.Rows[i]["Dur"],
                    ["Ratio"] = calendarDataTable.Rows[i]["Ratio"],
                    ["Sales Type"] = calendarDataTable.Rows[i]["Sales Type"],
                    ["ORD"] = calendarDataTable.Rows[i]["ORD"],
                    ["Sponsor Type"] = calendarDataTable.Rows[i]["Sponsor Type"],
                    ["Unit Price KWD"] = calendarDataTable.Rows[i]["Unit Price KWD"],
                    ["Price in US $"] = calendarDataTable.Rows[i]["Price in US $"]
                };

                // Save spot values by actual date
                for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                {
                    string columnName = GetDateColumnName(date);
                    try
                    {
                        if (calendarDataTable.Columns.Contains(columnName))
                        {
                            rowData[date.ToString("yyyy-MM-dd")] = Convert.ToInt32(calendarDataTable.Rows[i][columnName]);
                        }
                    }
                    catch
                    {
                        rowData[date.ToString("yyyy-MM-dd")] = 0;
                    }
                }

                savedRows.Add(rowData);
            }

            // Update campaign dates
            campaignStartDate = newStartDate;
            campaignEndDate = newEndDate;

            // Rebuild the DataTable with new date columns
            calendarDataTable = new DataTable();

            // Add metadata columns
            calendarDataTable.Columns.Add("OID", typeof(string));
            calendarDataTable.Columns.Add("Time", typeof(string));
            calendarDataTable.Columns.Add("Programme", typeof(string));
            calendarDataTable.Columns.Add("F/P", typeof(string));
            calendarDataTable.Columns.Add("Dur", typeof(string));
            calendarDataTable.Columns.Add("Ratio", typeof(string));
            calendarDataTable.Columns.Add("Sales Type", typeof(string));
            calendarDataTable.Columns.Add("ORD", typeof(string));
            calendarDataTable.Columns.Add("Sponsor Type", typeof(string));
            calendarDataTable.Columns.Add("Unit Price KWD", typeof(string));
            calendarDataTable.Columns.Add("Price in US $", typeof(string));

            // Add new date columns for the campaign range
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);
                calendarDataTable.Columns.Add(columnName, typeof(int));
            }

            calendarDataTable.Columns.Add("Total Spots", typeof(int));

            // Restore saved rows
            foreach (var rowData in savedRows)
            {
                DataRow newRow = calendarDataTable.NewRow();

                newRow["OID"] = rowData["OID"];
                newRow["Time"] = rowData["Time"];
                newRow["Programme"] = rowData["Programme"];
                newRow["F/P"] = rowData["F/P"];
                newRow["Dur"] = rowData["Dur"];
                newRow["Ratio"] = rowData["Ratio"];
                newRow["Sales Type"] = rowData["Sales Type"];
                newRow["ORD"] = rowData["ORD"];
                newRow["Sponsor Type"] = rowData["Sponsor Type"];
                newRow["Unit Price KWD"] = rowData["Unit Price KWD"];
                newRow["Price in US $"] = rowData["Price in US $"];

                int rowTotal = 0;
                for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                {
                    string columnName = GetDateColumnName(date);
                    string dateKey = date.ToString("yyyy-MM-dd");

                    // Restore spot value if date existed in old range
                    int spotValue = 0;
                    if (rowData.ContainsKey(dateKey))
                    {
                        spotValue = Convert.ToInt32(rowData[dateKey]);
                    }
                    newRow[columnName] = spotValue;
                    rowTotal += spotValue;
                }

                newRow["Total Spots"] = rowTotal;
                calendarDataTable.Rows.Add(newRow);
            }

            // Add Total row
            DataRow totalRow = calendarDataTable.NewRow();
            totalRow["OID"] = "";
            totalRow["Time"] = "";
            totalRow["Programme"] = "Total";
            totalRow["F/P"] = "";
            totalRow["Dur"] = "";
            totalRow["Ratio"] = "";
            totalRow["Sales Type"] = "";
            totalRow["ORD"] = "";
            totalRow["Sponsor Type"] = "";
            totalRow["Unit Price KWD"] = "";
            totalRow["Price in US $"] = "";

            int grandTotal = 0;
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);

                int dayTotal = 0;
                for (int i = 0; i < calendarDataTable.Rows.Count; i++)
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
            calendarDataTable.Rows.Add(totalRow);

            // Rebind to grid
            dgvCalendar.DataSource = null;
            dgvCalendar.DataSource = calendarDataTable;

            // Reapply column formatting
            ApplyGridFormatting();

            isUpdating = false;

            // Update labels
            lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
            if (currentBookingOrder != null)
            {
                currentBookingOrder.TotalSpots = grandTotal;
                lblTotalSpots.Text = $"Total Spots: {grandTotal}";
            }
        }

        private void ApplyGridFormatting()
        {
            // Set all columns to Programmatic sort mode for custom sorting
            foreach (DataGridViewColumn column in dgvCalendar.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }
            dgvCalendar.ColumnHeaderMouseClick -= DgvCalendar_ColumnHeaderMouseClick;
            dgvCalendar.ColumnHeaderMouseClick += DgvCalendar_ColumnHeaderMouseClick;

            // Set fixed columns widths
            if (dgvCalendar.Columns.Contains("OID"))
                dgvCalendar.Columns["OID"].Width = 80;
            if (dgvCalendar.Columns.Contains("Time"))
                dgvCalendar.Columns["Time"].Width = 60;
            if (dgvCalendar.Columns.Contains("Programme"))
                dgvCalendar.Columns["Programme"].Width = 150;
            if (dgvCalendar.Columns.Contains("F/P"))
                dgvCalendar.Columns["F/P"].Width = 40;
            if (dgvCalendar.Columns.Contains("Dur"))
                dgvCalendar.Columns["Dur"].Width = 40;
            if (dgvCalendar.Columns.Contains("Ratio"))
                dgvCalendar.Columns["Ratio"].Width = 50;
            if (dgvCalendar.Columns.Contains("Sales Type"))
                dgvCalendar.Columns["Sales Type"].Width = 70;
            if (dgvCalendar.Columns.Contains("ORD"))
                dgvCalendar.Columns["ORD"].Width = 40;
            if (dgvCalendar.Columns.Contains("Sponsor Type"))
                dgvCalendar.Columns["Sponsor Type"].Width = 90;
            if (dgvCalendar.Columns.Contains("Unit Price KWD"))
                dgvCalendar.Columns["Unit Price KWD"].Width = 100;
            if (dgvCalendar.Columns.Contains("Price in US $"))
                dgvCalendar.Columns["Price in US $"].Width = 90;
            if (dgvCalendar.Columns.Contains("Total Spots"))
                dgvCalendar.Columns["Total Spots"].Width = 80;

            // Convert Programme and Sales Type columns to ComboBox columns
            //ConvertToComboBoxColumn("Programme", GetProgrammeNameList());
            ConvertToComboBoxColumn("Sales Type", GetBreakInList());

            // Set date column widths
            foreach (DataGridViewColumn col in dgvCalendar.Columns)
            {
                if (col.Name.Contains("\n"))
                {
                    col.Width = 40;
                }
            }

            // Make Total row read-only
            if (calendarDataTable != null && calendarDataTable.Rows.Count > 0)
            {
                int totalRowIndex = calendarDataTable.Rows.Count - 1;
                dgvCalendar.Rows[totalRowIndex].ReadOnly = true;
                dgvCalendar.Rows[totalRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                dgvCalendar.Rows[totalRowIndex].DefaultCellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
            }

            // Configure OID column based on Grid Mode selection
            if (rbCampaignDates != null && rbCampaignDates.Checked)
            {
                SetOIDColumnEnabled(false);
            }
            else if (rbSpecificDate != null && rbSpecificDate.Checked)
            {
                SetOIDColumnEnabled(true);
            }
        }

        /// <summary>
        /// Handle column header click for sorting.
        /// Removes Total row before sorting, sorts data rows, then adds Total row back.
        /// This ensures Total row always stays at the bottom.
        /// </summary>
        private void DgvCalendar_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (calendarDataTable == null || calendarDataTable.Rows.Count <= 1)
                return;

            var clickedColumn = dgvCalendar.Columns[e.ColumnIndex];

            // Skip hidden or non-sortable columns
            if (!clickedColumn.Visible || clickedColumn.SortMode == DataGridViewColumnSortMode.NotSortable)
                return;

            string columnName = clickedColumn.Name;

            // Determine sort direction (toggle)
            ListSortDirection direction = ListSortDirection.Ascending;
            if (clickedColumn.HeaderCell.SortGlyphDirection == SortOrder.Ascending)
            {
                direction = ListSortDirection.Descending;
            }

            // Step 1: Save and remove Total row (last row)
            int totalRowIndex = calendarDataTable.Rows.Count - 1;
            object[] totalRowData = calendarDataTable.Rows[totalRowIndex].ItemArray.Clone() as object[];
            calendarDataTable.Rows.RemoveAt(totalRowIndex);

            // Step 2: Sort the data rows
            DataView dv = calendarDataTable.DefaultView;
            dv.Sort = columnName + (direction == ListSortDirection.Descending ? " DESC" : " ASC");

            // Step 3: Rebuild table with sorted data
            DataTable sortedTable = dv.ToTable();

            // Clear the sort to prevent auto-resorting on data changes
            dv.Sort = "";

            calendarDataTable.Rows.Clear();
            foreach (DataRow row in sortedTable.Rows)
            {
                calendarDataTable.ImportRow(row);
            }

            // Step 4: Add Total row back at the bottom
            DataRow newTotalRow = calendarDataTable.NewRow();
            newTotalRow.ItemArray = totalRowData;
            calendarDataTable.Rows.Add(newTotalRow);

            // Update sort glyph
            foreach (DataGridViewColumn col in dgvCalendar.Columns)
            {
                if (col.Visible && col.SortMode == DataGridViewColumnSortMode.Programmatic)
                {
                    col.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
            clickedColumn.HeaderCell.SortGlyphDirection =
                direction == ListSortDirection.Ascending ? SortOrder.Ascending : SortOrder.Descending;

            // Reapply Total row styling
            StyleTotalRow();
        }

        /// <summary>
        /// Apply styling to the Total row (last row): gray background, bold font, read-only.
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

        private void SetupContextMenu()
        {
            ContextMenuStrip contextMenu = new ContextMenuStrip();

            var addRowItem = new ToolStripMenuItem("Add Row");
            addRowItem.Click += AddRow_Click;
            contextMenu.Items.Add(addRowItem);

            var duplicateRowItem = new ToolStripMenuItem("Duplicate Row(s)");
            duplicateRowItem.Click += DuplicateRows_Click;
            contextMenu.Items.Add(duplicateRowItem);

            var insertRowsItem = new ToolStripMenuItem("Insert Rows...");
            insertRowsItem.Click += InsertRows_Click;
            contextMenu.Items.Add(insertRowsItem);

            var deleteRowItem = new ToolStripMenuItem("Delete Rows");
            deleteRowItem.Click += DeleteRows_Click;
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

        private void DgvCalendar_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            var column = dgvCalendar.Columns[e.ColumnIndex];
            if (column is DataGridViewComboBoxColumn comboColumn)
            {
                string newValue = e.FormattedValue?.ToString();
                if (!string.IsNullOrEmpty(newValue) && !comboColumn.Items.Contains(newValue))
                {
                    comboColumn.Items.Add(newValue);
                }
            }
        }

        private void DgvCalendar_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception is ArgumentException && e.Exception.Message.Contains("DataGridViewComboBoxCell"))
            {
                var column = dgvCalendar.Columns[e.ColumnIndex];
                var cell = dgvCalendar[e.ColumnIndex, e.RowIndex];

                if (column is DataGridViewComboBoxColumn comboColumn)
                {
                    var currentValue = cell.Value?.ToString();
                    if (!string.IsNullOrEmpty(currentValue) && !comboColumn.Items.Contains(currentValue))
                    {
                        comboColumn.Items.Add(currentValue);
                    }
                }

                e.ThrowException = false;
            }
        }

        private void DgvCalendar_KeyDown(object sender, KeyEventArgs e)
        {
            // Handle suggestion list navigation first
            if (suggestionListBox != null && suggestionListBox.Visible)
            {
                if (e.KeyCode == Keys.Down)
                {
                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (suggestionListBox.SelectedIndex < suggestionListBox.Items.Count - 1)
                            suggestionListBox.SelectedIndex++;
                        else
                            suggestionListBox.SelectedIndex = 0;
                    }
                    e.Handled = true;
                    return;
                }
                else if (e.KeyCode == Keys.Up)
                {
                    if (suggestionListBox.Items.Count > 0)
                    {
                        if (suggestionListBox.SelectedIndex > 0)
                            suggestionListBox.SelectedIndex--;
                        else
                            suggestionListBox.SelectedIndex = suggestionListBox.Items.Count - 1;
                    }
                    e.Handled = true;
                    return;
                }
                else if (e.KeyCode == Keys.Enter)
                {
                    if (suggestionListBox.SelectedIndex >= 0)
                    {
                        SelectSuggestion();
                    }
                    e.Handled = true;
                    return;
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    HideSuggestionList();
                    e.Handled = true;
                    return;
                }
            }

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
            if (calendarDataTable == null)
                return;

            int insertIndex = calendarDataTable.Rows.Count - 1;

            if (dgvCalendar.CurrentCell != null)
            {
                int currentRowIndex = dgvCalendar.CurrentCell.RowIndex;
                if (currentRowIndex < calendarDataTable.Rows.Count - 1)
                {
                    insertIndex = currentRowIndex;
                }
            }

            DataRow newRow = calendarDataTable.NewRow();

            newRow["OID"] = "";
            newRow["Time"] = "";
            newRow["Programme"] = "";
            newRow["F/P"] = "P";
            newRow["Dur"] = "";
            newRow["Ratio"] = "";
            newRow["Sales Type"] = "WN";
            newRow["ORD"] = "";
            newRow["Sponsor Type"] = "";
            newRow["Unit Price KWD"] = "";
            newRow["Price in US $"] = "";

            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);
                newRow[columnName] = 0;
            }

            newRow["Total Spots"] = 0;

            calendarDataTable.Rows.InsertAt(newRow, insertIndex);
            RecalculateColumnTotals();
            lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
        }

        private void DuplicateRows_Click(object sender, EventArgs e)
        {
            if (calendarDataTable == null || dgvCalendar.SelectedCells.Count == 0)
                return;

            // Get unique selected row indices, sorted ascending
            var selectedRowIndices = dgvCalendar.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Distinct()
                .OrderBy(i => i)
                .ToList();

            // Filter out the Total row (last row)
            int totalRowIndex = calendarDataTable.Rows.Count - 1;
            selectedRowIndices = selectedRowIndices.Where(i => i < totalRowIndex).ToList();

            if (selectedRowIndices.Count == 0)
            {
                MessageBox.Show("Please select data rows to duplicate (not the Total row).",
                    "No Valid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // First, collect all row data BEFORE inserting (to avoid index shifting issues)
            var rowDataList = new List<object[]>();
            foreach (int rowIndex in selectedRowIndices)
            {
                DataRow sourceRow = calendarDataTable.Rows[rowIndex];
                object[] rowData = new object[calendarDataTable.Columns.Count];
                for (int col = 0; col < calendarDataTable.Columns.Count; col++)
                {
                    rowData[col] = sourceRow[col];
                }
                rowDataList.Add(rowData);
            }

            // Insert duplicates above the first selected row
            int insertIndex = selectedRowIndices[0];

            isUpdating = true;
            try
            {
                // Insert duplicates in same order as selected
                for (int i = 0; i < rowDataList.Count; i++)
                {
                    DataRow newRow = calendarDataTable.NewRow();

                    // Copy all column values from the saved data
                    for (int col = 0; col < calendarDataTable.Columns.Count; col++)
                    {
                        newRow[col] = rowDataList[i][col];
                    }

                    calendarDataTable.Rows.InsertAt(newRow, insertIndex + i);
                }

                RecalculateColumnTotals();
                lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
            }
            finally
            {
                isUpdating = false;
            }
        }

        private void InsertRows_Click(object sender, EventArgs e)
        {
            if (calendarDataTable == null)
                return;

            // Get the insertion point (above current selection or before Total row)
            int insertIndex = calendarDataTable.Rows.Count - 1; // Default: before Total row

            if (dgvCalendar.CurrentCell != null)
            {
                int currentRowIndex = dgvCalendar.CurrentCell.RowIndex;
                // Don't allow insertion at or after Total row
                if (currentRowIndex < calendarDataTable.Rows.Count - 1)
                {
                    insertIndex = currentRowIndex;
                }
            }

            // Show input dialog to get number of rows
            using (Form inputForm = new Form())
            {
                inputForm.Text = "Insert Rows";
                inputForm.Width = 300;
                inputForm.Height = 150;
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputForm.StartPosition = FormStartPosition.CenterParent;
                inputForm.MaximizeBox = false;
                inputForm.MinimizeBox = false;

                Label label = new Label()
                {
                    Text = "Number of rows to insert:",
                    Left = 20,
                    Top = 20,
                    Width = 150
                };

                NumericUpDown numericInput = new NumericUpDown()
                {
                    Left = 170,
                    Top = 18,
                    Width = 80,
                    Minimum = 1,
                    Maximum = 100,
                    Value = 1
                };

                Button okButton = new Button()
                {
                    Text = "OK",
                    Left = 100,
                    Top = 70,
                    Width = 80,
                    DialogResult = DialogResult.OK
                };

                Button cancelButton = new Button()
                {
                    Text = "Cancel",
                    Left = 190,
                    Top = 70,
                    Width = 80,
                    DialogResult = DialogResult.Cancel
                };

                inputForm.Controls.Add(label);
                inputForm.Controls.Add(numericInput);
                inputForm.Controls.Add(okButton);
                inputForm.Controls.Add(cancelButton);
                inputForm.AcceptButton = okButton;
                inputForm.CancelButton = cancelButton;

                if (inputForm.ShowDialog(this) == DialogResult.OK)
                {
                    int rowCount = (int)numericInput.Value;
                    InsertBlankRows(insertIndex, rowCount);
                }
            }
        }

        private void InsertBlankRows(int insertIndex, int count)
        {
            if (calendarDataTable == null || count <= 0)
                return;

            isUpdating = true;
            try
            {
                for (int i = 0; i < count; i++)
                {
                    DataRow newRow = calendarDataTable.NewRow();

                    newRow["OID"] = "";
                    newRow["Time"] = "";
                    newRow["Programme"] = "";
                    newRow["F/P"] = "P";
                    newRow["Dur"] = "";
                    newRow["Ratio"] = "";
                    newRow["Sales Type"] = "WN";
                    newRow["ORD"] = "";
                    newRow["Sponsor Type"] = "";
                    newRow["Unit Price KWD"] = "";
                    newRow["Price in US $"] = "";

                    for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                    {
                        string columnName = GetDateColumnName(date);
                        newRow[columnName] = 0;
                    }

                    newRow["Total Spots"] = 0;

                    calendarDataTable.Rows.InsertAt(newRow, insertIndex);
                }

                RecalculateColumnTotals();
                lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
            }
            finally
            {
                isUpdating = false;
            }
        }

        private void DeleteRows_Click(object sender, EventArgs e)
        {
            if (calendarDataTable == null || dgvCalendar.SelectedCells.Count == 0)
            {
                MessageBox.Show("Please select row(s) to delete.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Get unique selected row indices
            var selectedRowIndices = dgvCalendar.SelectedCells
                .Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Distinct()
                .ToList();

            // Filter out the Total row (last row) and sort descending for safe deletion
            int totalRowIndex = calendarDataTable.Rows.Count - 1;
            selectedRowIndices = selectedRowIndices
                .Where(i => i >= 0 && i < totalRowIndex)
                .OrderByDescending(i => i)
                .ToList();

            if (selectedRowIndices.Count == 0)
            {
                MessageBox.Show("Cannot delete the Total row. Please select data rows to delete.",
                    "Invalid Operation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string confirmMessage = selectedRowIndices.Count == 1
                ? "Delete this row?"
                : $"Delete {selectedRowIndices.Count} selected rows?";

            if (MessageBox.Show(confirmMessage, "Confirm Delete",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                isUpdating = true;
                try
                {
                    // Collect DataRow references first (before any deletion)
                    var rowsToDelete = new List<DataRow>();
                    foreach (int rowIndex in selectedRowIndices)
                    {
                        rowsToDelete.Add(calendarDataTable.Rows[rowIndex]);
                    }

                    // Delete the collected rows
                    foreach (DataRow row in rowsToDelete)
                    {
                        calendarDataTable.Rows.Remove(row);
                    }

                    RecalculateColumnTotals();
                    lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";
                }
                finally
                {
                    isUpdating = false;
                }
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
            if (dgvCalendar.CurrentCell == null || calendarDataTable == null)
                return;

            // Cancel any active cell edit to prevent the current cell from overriding pasted values
            dgvCalendar.CancelEdit();

            string clipboardText = Clipboard.GetText();
            if (string.IsNullOrEmpty(clipboardText))
                return;

            int startRow = dgvCalendar.CurrentCell.RowIndex;
            int startCol = dgvCalendar.CurrentCell.ColumnIndex;

            if (startRow == calendarDataTable.Rows.Count - 1)
            {
                MessageBox.Show("Cannot paste into the Total row.", "Invalid Operation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Split on \r\n or \n (preserve empty rows so blanks paste as-is)
            string[] rows = clipboardText.TrimEnd('\r', '\n').Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            isUpdating = true;
            dgvCalendar.SuspendLayout();

            try
            {
                for (int i = 0; i < rows.Length; i++)
                {
                    int currentRow = startRow + i;
                    if (currentRow >= calendarDataTable.Rows.Count - 1)
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
            if (dgvCalendar.SelectedCells.Count == 0 || calendarDataTable == null)
                return;

            isUpdating = true;
            dgvCalendar.SuspendLayout();

            try
            {
                foreach (DataGridViewCell cell in dgvCalendar.SelectedCells)
                {
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
            // Don't allow editing the Total row
            if (calendarDataTable != null && e.RowIndex == calendarDataTable.Rows.Count - 1)
            {
                e.Cancel = true;
                return;
            }
        }

        /// <summary>
        /// Handle EditingControlShowing to set up autocomplete suggestion lists
        /// for Time, Programme, and Sales Type columns.
        /// </summary>
        private void DgvCalendar_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            string columnName = dgvCalendar.Columns[dgvCalendar.CurrentCell.ColumnIndex].Name;

            // Remove previous event handlers to avoid duplicates
            if (e.Control is TextBox textBox)
            {
                textBox.TextChanged -= EditingControl_TextChanged;
                textBox.PreviewKeyDown -= EditingControl_PreviewKeyDown;
                textBox.KeyDown -= EditingControl_KeyDown;
                textBox.LostFocus -= EditingControl_LostFocus;
            }
            if (e.Control is DataGridViewComboBoxEditingControl comboBox)
            {
                comboBox.TextUpdate -= EditingControl_TextUpdate;
                comboBox.PreviewKeyDown -= EditingControl_PreviewKeyDown;
                comboBox.KeyDown -= EditingControl_KeyDown;
                comboBox.LostFocus -= EditingControl_LostFocus;
                comboBox.Leave -= ComboBox_Leave;
            }

            // Setup autocomplete suggestion for Time, Programme, and Sales Type columns
            if (columnName == "Time" || columnName == "Programme")
            {
                // Handle as TextBox (Programme was changed from ComboBox to TextBox)
                if (e.Control is TextBox editTextBox)
                {
                    currentEditingControl = editTextBox;
                    currentEditingColumn = columnName;

                    editTextBox.TextChanged += EditingControl_TextChanged;
                    editTextBox.PreviewKeyDown += EditingControl_PreviewKeyDown;
                    editTextBox.KeyDown += EditingControl_KeyDown;
                    editTextBox.LostFocus += EditingControl_LostFocus;

                    // Show suggestion list immediately on focus for Programme column
                    if (columnName == "Programme")
                    {
                        BeginInvoke(new Action(() =>
                        {
                            ShowSuggestionList(editTextBox, columnName, editTextBox.Text);
                        }));
                    }
                }
                // Also handle ComboBox case for backwards compatibility
                else if (e.Control is DataGridViewComboBoxEditingControl editComboBox)
                {
                    currentEditingControl = editComboBox;
                    currentEditingColumn = columnName;

                    editComboBox.DropDownStyle = ComboBoxStyle.DropDown;
                    editComboBox.TextUpdate += EditingControl_TextUpdate;
                    editComboBox.PreviewKeyDown += EditingControl_PreviewKeyDown;
                    editComboBox.KeyDown += EditingControl_KeyDown;
                    editComboBox.LostFocus += EditingControl_LostFocus;
                    editComboBox.Leave += ComboBox_Leave;

                    // Show suggestion list immediately on focus for Programme column
                    if (columnName == "Programme")
                    {
                        BeginInvoke(new Action(() =>
                        {
                            ShowSuggestionList(editComboBox, columnName, editComboBox.Text);
                        }));
                    }
                }
            }
            else if (columnName == "Sales Type")
            {
                if (e.Control is DataGridViewComboBoxEditingControl editComboBox)
                {
                    currentEditingControl = editComboBox;
                    currentEditingColumn = columnName;

                    editComboBox.DropDownStyle = ComboBoxStyle.DropDown;
                    editComboBox.TextUpdate += EditingControl_TextUpdate;
                    editComboBox.PreviewKeyDown += EditingControl_PreviewKeyDown;
                    editComboBox.KeyDown += EditingControl_KeyDown;
                    editComboBox.LostFocus += EditingControl_LostFocus;
                    editComboBox.Leave += ComboBox_Leave;
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
                    comboBox.Items.Add(currentText);

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

            if (columnName == "Time")
            {
                var cell = dgvCalendar[e.ColumnIndex, e.RowIndex];
                if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                {
                    string formattedTime = FormatTimeInput(cell.Value.ToString());
                    if (formattedTime != null)
                    {
                        cell.Value = formattedTime;

                        // Auto-fill Programme Name only in Campaign Dates mode
                        if (rbCampaignDates != null && rbCampaignDates.Checked)
                        {
                            var programmes = GetProgrammeNameListByTime(formattedTime);
                            if (programmes.Count == 1)
                            {
                                var programNameCell = dgvCalendar["Programme", e.RowIndex];

                                // Add to combobox items if not exists
                                var column = dgvCalendar.Columns["Programme"] as DataGridViewComboBoxColumn;
                                if (column != null && !column.Items.Contains(programmes[0]))
                                {
                                    column.Items.Add(programmes[0]);
                                }

                                programNameCell.Value = programmes[0];

                                // Auto-fill F/P and recalculate pricing
                                AutoFillFP(e.RowIndex, programmes[0], formattedTime);
                                RecalculatePricing();
                            }
                            // If multiple programmes, leave blank for user to pick from list
                        }
                    }
                }
            }
            else if (columnName == "Programme")
            {
                // Auto-fill F/P when Programme is set
                string programme = dgvCalendar["Programme", e.RowIndex].Value?.ToString() ?? "";
                string time = dgvCalendar["Time", e.RowIndex].Value?.ToString() ?? "";

                if (!string.IsNullOrWhiteSpace(programme) && !string.IsNullOrWhiteSpace(time))
                {
                    AutoFillFP(e.RowIndex, programme, time);
                    RecalculatePricing();
                }
            }
            else if (columnName == "OID" && rbSpecificDate != null && rbSpecificDate.Checked)
            {
                // Auto-fill Time, Programme, F/P when OID changes (Specific Date mode)
                string oidStr = dgvCalendar["OID", e.RowIndex].Value?.ToString() ?? "";

                // Validate OID is an integer
                if (!string.IsNullOrWhiteSpace(oidStr) && int.TryParse(oidStr, out int oid))
                {
                    DateTime selectedDate = dtpSpecificDate.Value;
                    var details = GetProgrammeDetailsByOID(oidStr, selectedDate);

                    if (!string.IsNullOrWhiteSpace(details.Time))
                    {
                        // Format time as HH:mm (remove seconds if present)
                        string formattedTime = details.Time;
                        if (details.Time.Length >= 5)
                        {
                            formattedTime = details.Time.Substring(0, 5);
                        }
                        dgvCalendar["Time", e.RowIndex].Value = formattedTime;
                    }

                    if (!string.IsNullOrWhiteSpace(details.Programme))
                    {
                        // Add to combobox items if not exists
                        var column = dgvCalendar.Columns["Programme"] as DataGridViewComboBoxColumn;
                        if (column != null && !column.Items.Contains(details.Programme))
                        {
                            column.Items.Add(details.Programme);
                        }
                        dgvCalendar["Programme", e.RowIndex].Value = details.Programme;
                    }

                    if (!string.IsNullOrWhiteSpace(details.FP))
                    {
                        dgvCalendar["F/P", e.RowIndex].Value = details.FP;
                    }

                    // Recalculate pricing after OID auto-fill
                    RecalculatePricing();
                }
                else if (!string.IsNullOrWhiteSpace(oidStr))
                {
                    // Clear invalid OID value
                    dgvCalendar["OID", e.RowIndex].Value = "";
                    MessageBox.Show("OID must be an integer value.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else if (columnName == "Dur")
            {
                // Auto-fill Ratio when Dur is set
                string durStr = dgvCalendar["Dur", e.RowIndex].Value?.ToString() ?? "";

                if (!string.IsNullOrWhiteSpace(durStr) && int.TryParse(durStr, out int dur))
                {
                    // Validate dur is between 0 and 215
                    if (dur >= 0 && dur <= 215)
                    {
                        string ratio = GetRatioByDur(dur);
                        if (!string.IsNullOrWhiteSpace(ratio))
                        {
                            dgvCalendar["Ratio", e.RowIndex].Value = ratio;
                        }
                        // Recalculate pricing after Dur/Ratio change
                        RecalculatePricing();
                    }
                    else
                    {
                        // Clear invalid Dur value (outside 0-215 range)
                        dgvCalendar["Dur", e.RowIndex].Value = "";
                    }
                }
                else if (!string.IsNullOrWhiteSpace(durStr))
                {
                    // Clear invalid Dur value (not an integer)
                    dgvCalendar["Dur", e.RowIndex].Value = "";
                }
            }
        }

        private void AutoFillFP(int rowIndex, string programme, string time)
        {
            string paymentRatio = GetPaymentRatio(programme, time, campaignStartDate, campaignEndDate);

            if (!string.IsNullOrWhiteSpace(paymentRatio))
            {
                dgvCalendar["F/P", rowIndex].Value = paymentRatio;
            }
        }

        /// <summary>
        /// Auto-fills Ratio and F/P for all rows after grid is loaded from Excel/JSON.
        /// This ensures auto-fill logic runs even when data is loaded programmatically
        /// (not through manual cell editing which triggers CellEndEdit).
        /// </summary>
        private void AutoFillGridAfterLoad()
        {
            if (dgvCalendar == null || dgvCalendar.Rows.Count == 0)
                return;

            isUpdating = true;
            try
            {
                for (int rowIndex = 0; rowIndex < dgvCalendar.Rows.Count; rowIndex++)
                {
                    // Skip the Total row (last row where Programme = "Total")
                    string programmeValue = dgvCalendar["Programme", rowIndex].Value?.ToString() ?? "";
                    if (programmeValue == "Total")
                        continue;

                    string time = dgvCalendar["Time", rowIndex].Value?.ToString() ?? "";
                    string programme = programmeValue;
                    string durStr = dgvCalendar["Dur", rowIndex].Value?.ToString() ?? "";

                    // Auto-fill Ratio based on Duration (if valid 0-215)
                    if (!string.IsNullOrWhiteSpace(durStr) && int.TryParse(durStr, out int dur))
                    {
                        if (dur >= 0 && dur <= 215)
                        {
                            string ratio = GetRatioByDur(dur);
                            if (!string.IsNullOrWhiteSpace(ratio))
                            {
                                dgvCalendar["Ratio", rowIndex].Value = ratio;
                            }
                        }
                    }

                    // Auto-fill F/P based on Programme and Time (if both are present)
                    if (!string.IsNullOrWhiteSpace(programme) && !string.IsNullOrWhiteSpace(time))
                    {
                        string paymentRatio = GetPaymentRatio(programme, time, campaignStartDate, campaignEndDate);
                        if (!string.IsNullOrWhiteSpace(paymentRatio))
                        {
                            dgvCalendar["F/P", rowIndex].Value = paymentRatio;
                        }
                    }
                }

                // Recalculate pricing after auto-fill
                RecalculatePricing();
            }
            finally
            {
                isUpdating = false;
            }
        }

        private string FormatTimeInput(string input)
        {
            string digits = new string(input.Where(char.IsDigit).ToArray());

            if (string.IsNullOrEmpty(digits))
                return null;

            if (digits.Length == 1)
                digits = "0" + digits + "00";
            else if (digits.Length == 2)
                digits = digits + "00";
            else if (digits.Length == 3)
                digits = "0" + digits;
            else if (digits.Length > 4)
                digits = digits.Substring(0, 4);

            int hours = int.Parse(digits.Substring(0, 2));
            int minutes = int.Parse(digits.Substring(2, 2));

            if (hours > 29 || minutes > 59)
                return null;

            return $"{hours:D2}:{minutes:D2}";
        }

        private void DgvCalendar_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || calendarDataTable == null)
                return;

            string columnName = dgvCalendar.Columns[e.ColumnIndex].Name;

            if (columnName.Contains("\n") && char.IsDigit(columnName.Split('\n')[1][0]))
            {
                if (e.Value != null && e.Value is int intValue && intValue == 0)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }

                int dayNumber = int.Parse(columnName.Split('\n')[1]);
                if (dayNumber % 2 == 0)
                {
                    e.CellStyle.BackColor = System.Drawing.Color.FromArgb(245, 250, 255);
                }
                else
                {
                    e.CellStyle.BackColor = System.Drawing.Color.White;
                }
            }

            if (e.RowIndex == calendarDataTable.Rows.Count - 1)
            {
                e.CellStyle.BackColor = System.Drawing.Color.FromArgb(220, 220, 220);
                e.CellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
                e.CellStyle.ForeColor = System.Drawing.Color.FromArgb(50, 50, 50);
            }
        }

        
        private List<string> GetProgrammeNameList()
        {
            
            //var db = new DatabaseHelper();
            //var programmes = db.GetComboBoxList(
                //"SELECT DISTINCT(prog_details.name_en) FROM u333577897_dbofhash.grid LEFT JOIN prog_details ON grid.tagid=prog_details.idprog_details ORDER BY prog_details.name_en");

            //if (programmes.Count > 0)
                //return programmes;

            // Fallback to defaults if database returns empty
            return new List<string>
            {
                "Masrah Al Hayat",
                "Talk Show 'With Bu Shoail'",
                "Warahom Warahom",
                "News Bulletin",
                "Morning Show"
            };
            
        }
        

        private List<string> GetProgrammeNameListByTime(string time)
        {
            var db = new DatabaseHelper();

            // Format time to HH:mm:ss if needed
            string formattedTime = time;
            if (time.Length == 5) // HH:mm format
            {
                formattedTime = time + ":00";
            }

            // Get campaign period dates
            string startDate = dtpCampaignStart.Value.ToString("yyyy-MM-dd");
            string endDate = dtpCampaignEnd.Value.ToString("yyyy-MM-dd");

            var programmes = db.GetComboBoxList(
                $"SELECT DISTINCT(prog_details.name_en) FROM u333577897_dbofhash.grid " +
                $"LEFT JOIN prog_details ON grid.tagid = prog_details.idprog_details " +
                $"WHERE start_at = '{formattedTime}' " +
                $"AND (trans_dt >= '{startDate}' AND trans_dt <= '{endDate}') " +
                $"ORDER BY prog_details.name_en");

            return programmes;
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

            var oldColumn = dgvCalendar.Columns[columnIndex];
            int displayIndex = oldColumn.DisplayIndex;

            var existingValues = new HashSet<string>();
            foreach (DataRow row in calendarDataTable.Rows)
            {
                var value = row[columnName]?.ToString();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    existingValues.Add(value);
                }
            }

            var comboColumn = new DataGridViewComboBoxColumn
            {
                Name = columnName,
                HeaderText = columnName,
                DataPropertyName = columnName,
                Width = oldColumn.Width,
                DisplayIndex = displayIndex,
                FlatStyle = FlatStyle.Flat,
                DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };

            foreach (string value in existingValues.OrderBy(v => v))
            {
                if (!comboColumn.Items.Contains(value))
                {
                    comboColumn.Items.Add(value);
                }
            }

            foreach (string item in items)
            {
                if (!comboColumn.Items.Contains(item))
                {
                    comboColumn.Items.Add(item);
                }
            }

            dgvCalendar.Columns.Remove(oldColumn);
            dgvCalendar.Columns.Insert(columnIndex, comboColumn);

            comboColumn.ReadOnly = false;
            comboColumn.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(250, 250, 250);
        }
        


        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls|JSON Files|*.json|All Files|*.*";
                openFileDialog.Title = "Select Booking Order File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string extension = Path.GetExtension(openFileDialog.FileName).ToLower();

                    if (extension == ".json")
                    {
                        LoadJsonFile(openFileDialog.FileName);
                    }
                    else
                    {
                        ProcessExcelFile(openFileDialog.FileName);
                    }
                }
            }
        }

        private void LoadJsonFile(string filePath)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string jsonContent = File.ReadAllText(filePath);
                currentBookingOrder = JsonConvert.DeserializeObject<BookingOrder>(jsonContent);

                if (currentBookingOrder == null)
                {
                    MessageBox.Show("Failed to load the JSON file.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Save to configured output path, not the source file location
                string fileName = Path.GetFileName(filePath);
                currentJsonFilePath = Path.Combine(jsonOutputPath, fileName);
                LoadBookingOrderToForm();
                LoadCalendarData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading JSON file: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void btnImportJSON_Click(object sender, EventArgs e)
        {
            using (var importForm = new JsonImportForm())
            {
                if (importForm.ShowDialog() == DialogResult.OK)
                {
                    currentBookingOrder = importForm.Result;
                    string fileName = $"{currentBookingOrder.Agency}_{currentBookingOrder.Product}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
                    currentJsonFilePath = Path.Combine(jsonOutputPath, fileName);
                    LoadBookingOrderToForm();
                    LoadCalendarData();
                }
            }
        }

        private void ProcessExcelFile(string filePath)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                BookingOrderReader reader = new BookingOrderReader();
                string jsonResult = reader.ProcessBookingOrder(filePath);

                if (jsonResult.Contains("\"error\""))
                {
                    MessageBox.Show("Failed to process the booking order file.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                currentBookingOrder = JsonConvert.DeserializeObject<BookingOrder>(jsonResult);

                if (currentBookingOrder == null)
                {
                    MessageBox.Show("Failed to deserialize the booking order data.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Set the file path for saving later (but don't save automatically)
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string jsonFileName = $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
                currentJsonFilePath = Path.Combine(jsonOutputPath, jsonFileName);

                LoadBookingOrderToForm();
                LoadCalendarData();

                MessageBox.Show($"File processed successfully!\nUse Save button to save changes.",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing file: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void LoadBookingOrderToForm()
        {
            if (currentBookingOrder == null)
                return;

            // Load header data
            cmbAgency.Text = currentBookingOrder.Agency ?? "";
            cmbAdvertiser.Text = currentBookingOrder.Advertiser ?? "";
            cmbProduct.Text = currentBookingOrder.Product ?? "";

            if (DateTime.TryParse(currentBookingOrder.CampaignPeriod?.StartDate, out DateTime startDate))
                dtpCampaignStart.Value = startDate;

            if (DateTime.TryParse(currentBookingOrder.CampaignPeriod?.EndDate, out DateTime endDate))
                dtpCampaignEnd.Value = endDate;

            // Load Package Cost (GrossCost)
            txtPackageCost.Text = currentBookingOrder.GrossCost > 0
                ? currentBookingOrder.GrossCost.ToString("N2")
                : "";
        }

        private void btnCreateNew_Click(object sender, EventArgs e)
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(cmbAgency.Text))
            {
                MessageBox.Show("Please enter an Agency name.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbAgency.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(cmbAdvertiser.Text))
            {
                MessageBox.Show("Please enter an Advertiser name.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbAdvertiser.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(cmbProduct.Text))
            {
                MessageBox.Show("Please enter a Product name.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbProduct.Focus();
                return;
            }

            if (dtpCampaignEnd.Value < dtpCampaignStart.Value)
            {
                MessageBox.Show("Campaign End date must be on or after Campaign Start date.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Parse Package Cost
            decimal packageCost = 0;
            if (!string.IsNullOrWhiteSpace(txtPackageCost.Text))
            {
                decimal.TryParse(txtPackageCost.Text.Replace(",", ""), out packageCost);
            }

            // Create new booking order
            currentBookingOrder = new BookingOrder
            {
                Agency = cmbAgency.Text.Trim(),
                Advertiser = cmbAdvertiser.Text.Trim(),
                Product = cmbProduct.Text.Trim(),
                CampaignPeriod = new CampaignPeriod
                {
                    StartDate = dtpCampaignStart.Value.ToString("yyyy-MM-dd"),
                    EndDate = dtpCampaignEnd.Value.ToString("yyyy-MM-dd")
                },
                GrossCost = packageCost,
                Spots = new List<Spot>(),
                TotalSpots = 0
            };

            // Set the file path for saving later (but don't save automatically)
            string fileName = $"{currentBookingOrder.Agency}_{currentBookingOrder.Product}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
            currentJsonFilePath = Path.Combine(jsonOutputPath, fileName);

            // Load calendar
            LoadCalendarData();

            MessageBox.Show($"New booking created successfully!\nUse Save button to save changes.",
                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void LoadCalendarData()
        {
            if (currentBookingOrder == null)
                return;

            // Get campaign date range
            campaignStartDate = DateTime.Parse(currentBookingOrder.CampaignPeriod.StartDate);
            campaignEndDate = DateTime.Parse(currentBookingOrder.CampaignPeriod.EndDate);

            // Create DataTable with calendar structure
            calendarDataTable = new DataTable();

            // Add hidden sort order column (0 = data row, 1 = total row - ensures Total always at bottom)

            // Add metadata columns (in display order)
            calendarDataTable.Columns.Add("OID", typeof(string));
            calendarDataTable.Columns.Add("Time", typeof(string));
            calendarDataTable.Columns.Add("Programme", typeof(string));
            calendarDataTable.Columns.Add("F/P", typeof(string));
            calendarDataTable.Columns.Add("Dur", typeof(string));
            calendarDataTable.Columns.Add("Ratio", typeof(string));
            calendarDataTable.Columns.Add("Sales Type", typeof(string));
            calendarDataTable.Columns.Add("ORD", typeof(string));
            calendarDataTable.Columns.Add("Sponsor Type", typeof(string));
            calendarDataTable.Columns.Add("Unit Price KWD", typeof(string));
            calendarDataTable.Columns.Add("Price in US $", typeof(string));

            // Add date columns for the campaign range
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);
                calendarDataTable.Columns.Add(columnName, typeof(int));
            }

            // Add Total Spots column
            calendarDataTable.Columns.Add("Total Spots", typeof(int));

            // Group spots by program name and time
            if (currentBookingOrder.Spots != null && currentBookingOrder.Spots.Count > 0)
            {
                // Preserve original Excel row order (group duplicates but keep first-occurrence order)
                var spotGroups = currentBookingOrder.Spots
                    .Select((s, index) => new { Spot = s, Index = index })
                    .GroupBy(x => new { x.Spot.ProgrammeName, x.Spot.ProgrammeStartTime })
                    .OrderBy(g => g.Min(x => x.Index));

                foreach (var group in spotGroups)
                {
                    DataRow row = calendarDataTable.NewRow();

                    // Get duration from the first spot in the group
                    var firstSpot = group.First().Spot;

                    row["OID"] = "";
                    row["Time"] = group.Key.ProgrammeStartTime ?? "";
                    row["Programme"] = group.Key.ProgrammeName ?? "";
                    row["F/P"] = "P";
                    // Extract only digits from Duration (handles "30 sec", "30s", etc.)
                    string rawDuration = firstSpot.Duration ?? "";
                    string cleanDuration = new string(rawDuration.Where(char.IsDigit).ToArray());
                    row["Dur"] = cleanDuration;
                    row["Ratio"] = "";
                    row["Sales Type"] = "WN";
                    row["ORD"] = "";
                    row["Sponsor Type"] = "";
                    row["Unit Price KWD"] = "";
                    row["Price in US $"] = "";

                    var allDatesForGroup = group.SelectMany(x => x.Spot.Dates).ToList();
                    var dateCounts = allDatesForGroup
                        .GroupBy(d => d)
                        .ToDictionary(g => DateTime.Parse(g.Key), g => g.Count());

                    int totalSpots = 0;
                    for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                    {
                        string columnName = GetDateColumnName(date);

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
            }

            // Add Total row
            DataRow totalRow = calendarDataTable.NewRow();
            totalRow["OID"] = "";
            totalRow["Time"] = "";
            totalRow["Programme"] = "Total";
            totalRow["F/P"] = "";
            totalRow["Dur"] = "";
            totalRow["Ratio"] = "";
            totalRow["Sales Type"] = "";
            totalRow["ORD"] = "";
            totalRow["Sponsor Type"] = "";
            totalRow["Unit Price KWD"] = "";
            totalRow["Price in US $"] = "";

            int grandTotal = 0;
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);

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

            // Bind to grid
            dgvCalendar.DataSource = calendarDataTable;
            ConfigureGridAppearance();

            // Auto-fill Ratio and F/P for all rows based on loaded data
            AutoFillGridAfterLoad();

            // Show grid and controls
            dgvCalendar.Visible = true;
            btnSave.Visible = true;
            btnCalculate.Visible = true;
            lblRecordCount.Visible = true;
            lblRecordCount.Text = $"Total Programs: {calendarDataTable.Rows.Count - 1}";

            // Update Total Spots label
            lblTotalSpots.Text = $"Total Spots: {grandTotal}";
            if (currentBookingOrder != null)
            {
                currentBookingOrder.TotalSpots = grandTotal;
            }
        }

        private void ConfigureGridAppearance()
        {
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

            // ============================================
            // COLUMN WIDTHS CONFIGURATION - Edit values here
            // ============================================
            var columnWidths = new Dictionary<string, int>
            {
                { "OID", 60 },
                { "Time", 70 },
                { "Programme", 280 },
                { "F/P", 40 },
                { "Dur", 50 },
                { "Ratio", 50 },
                { "Sales Type", 80 },
                { "ORD", 50 },
                { "Sponsor Type", 90 },
                { "Unit Price KWD", 100 },
                { "Price in US $", 90 },
                { "Total Spots", 80 }
            };
            int dateColumnWidth = 35;  // Width for all date columns (Su\n1, Mo\n2, etc.)
            // ============================================

            // Apply column widths
            foreach (var kvp in columnWidths)
            {
                if (dgvCalendar.Columns[kvp.Key] != null)
                {
                    dgvCalendar.Columns[kvp.Key].Width = kvp.Value;
                }
            }

            // Convert Programme and Sales Type columns to ComboBox columns
            //ConvertToComboBoxColumn("Programme", GetProgrammeNameList());
            ConvertToComboBoxColumn("Sales Type", GetBreakInList());

            // Make other metadata columns editable (including F/P as normal cell)
            string[] otherMetadataColumns = new string[]
            {
                "OID", "Time", "Dur", "Ratio", "F/P", "ORD", "Sponsor Type", "Unit Price KWD", "Price in US $"
            };

            foreach (string colName in otherMetadataColumns)
            {
                if (dgvCalendar.Columns[colName] != null)
                {
                    dgvCalendar.Columns[colName].ReadOnly = false;
                    dgvCalendar.Columns[colName].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(250, 250, 250);
                }
            }

            // Style date columns
            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);

                if (dgvCalendar.Columns[columnName] != null)
                {
                    dgvCalendar.Columns[columnName].Width = dateColumnWidth;
                    dgvCalendar.Columns[columnName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }

            // Make Total Spots column read-only
            if (dgvCalendar.Columns["Total Spots"] != null)
            {
                dgvCalendar.Columns["Total Spots"].ReadOnly = true;
                dgvCalendar.Columns["Total Spots"].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                dgvCalendar.Columns["Total Spots"].DefaultCellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
            }

            // Style the Total row
            if (dgvCalendar.Rows.Count > 0)
            {
                var lastRow = dgvCalendar.Rows[dgvCalendar.Rows.Count - 1];
                lastRow.DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                lastRow.DefaultCellStyle.Font = new System.Drawing.Font(dgvCalendar.Font, System.Drawing.FontStyle.Bold);
                lastRow.ReadOnly = true;
            }

            // Configure OID column based on Grid Mode selection
            if (rbCampaignDates != null && rbCampaignDates.Checked)
            {
                SetOIDColumnEnabled(false);
            }
            else if (rbSpecificDate != null && rbSpecificDate.Checked)
            {
                SetOIDColumnEnabled(true);
            }

            // Setup context menu
            SetupContextMenu();
        }

        private void btnUploadCSV_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This feature will be implemented later.",
                "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnEditBooking_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This feature will be implemented later.",
                "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (currentBookingOrder == null || string.IsNullOrEmpty(currentJsonFilePath))
            {
                MessageBox.Show("No booking order to save.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                // Update booking order from form
                currentBookingOrder.Agency = cmbAgency.Text.Trim();
                currentBookingOrder.Advertiser = cmbAdvertiser.Text.Trim();
                currentBookingOrder.Product = cmbProduct.Text.Trim();
                currentBookingOrder.CampaignPeriod.StartDate = dtpCampaignStart.Value.ToString("yyyy-MM-dd");
                currentBookingOrder.CampaignPeriod.EndDate = dtpCampaignEnd.Value.ToString("yyyy-MM-dd");

                // Update Package Cost (GrossCost)
                if (decimal.TryParse(txtPackageCost.Text.Replace(",", ""), out decimal packageCost))
                {
                    currentBookingOrder.GrossCost = packageCost;
                }

                // Update from calendar grid
                UpdateBookingOrderFromCalendar();

                // Ensure directory exists
                string directory = Path.GetDirectoryName(currentJsonFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save to JSON
                string jsonContent = JsonConvert.SerializeObject(currentBookingOrder, Formatting.Indented);
                File.WriteAllText(currentJsonFilePath, jsonContent);

                MessageBox.Show($"Changes saved successfully to:\n{currentJsonFilePath}",
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
            if (calendarDataTable == null)
                return;

            currentBookingOrder.Spots.Clear();

            for (int i = 0; i < calendarDataTable.Rows.Count - 1; i++) // Skip Total row
            {
                DataRow row = calendarDataTable.Rows[i];

                string programmeName = row["Programme"].ToString();
                string programmeTime = row["Time"].ToString();

                List<string> dates = new List<string>();
                for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                {
                    string columnName = GetDateColumnName(date);

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
                        Duration = "",
                        Dates = dates,
                        TotalSpots = dates.Count
                    };
                    currentBookingOrder.Spots.Add(spot);
                }
            }

            currentBookingOrder.TotalSpots = currentBookingOrder.Spots.Sum(s => s.TotalSpots);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            MainMenuForm mainForm = new MainMenuForm();
            mainForm.Show();
            this.Close();
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            // Step 1: Validate that no Dur cells are empty
            List<int> emptyDurRows = new List<int>();
            for (int i = 0; i < dgvCalendar.Rows.Count - 1; i++) // Exclude Total row
            {
                var durValue = dgvCalendar["Dur", i].Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(durValue) || !int.TryParse(durValue, out _))
                {
                    emptyDurRows.Add(i + 1); // 1-based row number for display
                }
            }

            if (emptyDurRows.Count > 0)
            {
                string rowNumbers = string.Join(", ", emptyDurRows);
                MessageBox.Show($"Dur column must contain integer values.\n\nRows with empty or invalid Dur: {rowNumbers}",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Step 2: Get form field values
            string agency = cmbAgency.Text;
            string advertiser = cmbAdvertiser.Text;

            if (string.IsNullOrWhiteSpace(advertiser))
            {
                MessageBox.Show("Please enter an Advertiser name.",
                    "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DateTime campaignStart = dtpCampaignStart.Value;
            DateTime campaignEnd = dtpCampaignEnd.Value;

            // Step 3: Get campaign periods
            var periods = GetCampaignPeriods(campaignStart, campaignEnd);

            // Step 4: Group bookings by duration and period
            var groupedBookings = GroupBookingsByDurationAndPeriod(periods);

            // Show grouped bookings table
            ShowGroupedBookingsTable(groupedBookings, periods);

            // Step 5: Check for existing deals (show dialog last)
            var existingDeals = CheckForExistingDeals(advertiser, campaignStart, campaignEnd);

            if (existingDeals.Count > 0)
            {
                // Build deal info message
                string dealInfo = "Existing deals found for this advertiser:\n\n";
                foreach (var deal in existingDeals)
                {
                    dealInfo += $"Deal #{deal.Id}: {deal.Agency}\n";
                    dealInfo += $"  Campaign: {deal.CampaignStart:yyyy-MM-dd} to {deal.CampaignEnd:yyyy-MM-dd}\n\n";
                }
                dealInfo += "What would you like to do?";

                // Show custom dialog with three options
                var (result, selectedDealId) = ShowDealOptionsDialog(dealInfo, existingDeals);

                if (result == DealOption.Decline)
                {
                    // Just close dialog and focus on grid
                    dgvCalendar.Focus();
                    return;
                }
                else if (result == DealOption.AddToExisting)
                {
                    // Get the deal number for later use
                    currentDealNumber = selectedDealId;
                }
                else if (result == DealOption.Proceed)
                {
                    // Create a new deal
                    currentDealNumber = CreateNewDeal(agency, advertiser, campaignStart, campaignEnd);
                    if (currentDealNumber == 0)
                    {
                        MessageBox.Show("Failed to create new deal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                // No existing deals - create a new one
                currentDealNumber = CreateNewDeal(agency, advertiser, campaignStart, campaignEnd);
                if (currentDealNumber == 0)
                {
                    MessageBox.Show("Failed to create new deal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // Step 6: Process bookings for the deal (check existing, compare, insert/replace)
            string product = cmbProduct.Text;
            ProcessBookingsForDeal(groupedBookings, currentDealNumber, product, agency, advertiser, periods);

            // Step 7: Show report
            ShowBookingReport(groupedBookings, periods);
        }

        private List<DealInfo> CheckForExistingDeals(string advertiser, DateTime campaignStart, DateTime campaignEnd)
        {
            var deals = new List<DealInfo>();
            var db = new DatabaseHelper();

            // Expand search range by dealSearchDays
            DateTime searchStart = campaignStart.AddDays(-dealSearchDays);
            DateTime searchEnd = campaignEnd.AddDays(dealSearchDays);

            string startDate = searchStart.ToString("yyyy-MM-dd");
            string endDate = searchEnd.ToString("yyyy-MM-dd");

            // Escape apostrophes in advertiser name for SQL
            string escapedAdvertiser = advertiser?.Replace("'", "''") ?? "";

            try
            {
                // Query to find overlapping deals
                string query = $"SELECT id, agency, advertiser, c_start, c_end, file, created " +
                    $"FROM u333577897_dbofhash.bookingdetails " +
                    $"WHERE advertiser = '{escapedAdvertiser}' " +
                    $"AND ((c_start <= '{endDate}' AND c_end >= '{startDate}'))";

                var idList = db.GetComboBoxList(
                    $"SELECT id FROM u333577897_dbofhash.bookingdetails " +
                    $"WHERE advertiser = '{escapedAdvertiser}' " +
                    $"AND ((c_start <= '{endDate}' AND c_end >= '{startDate}'))");

                var agencyList = db.GetComboBoxList(
                    $"SELECT agency FROM u333577897_dbofhash.bookingdetails " +
                    $"WHERE advertiser = '{escapedAdvertiser}' " +
                    $"AND ((c_start <= '{endDate}' AND c_end >= '{startDate}'))");

                var cStartList = db.GetComboBoxList(
                    $"SELECT c_start FROM u333577897_dbofhash.bookingdetails " +
                    $"WHERE advertiser = '{escapedAdvertiser}' " +
                    $"AND ((c_start <= '{endDate}' AND c_end >= '{startDate}'))");

                var cEndList = db.GetComboBoxList(
                    $"SELECT c_end FROM u333577897_dbofhash.bookingdetails " +
                    $"WHERE advertiser = '{escapedAdvertiser}' " +
                    $"AND ((c_start <= '{endDate}' AND c_end >= '{startDate}'))");

                for (int i = 0; i < idList.Count; i++)
                {
                    deals.Add(new DealInfo
                    {
                        Id = int.Parse(idList[i]),
                        Agency = i < agencyList.Count ? agencyList[i] : "",
                        Advertiser = advertiser,
                        CampaignStart = i < cStartList.Count ? DateTime.Parse(cStartList[i]) : DateTime.MinValue,
                        CampaignEnd = i < cEndList.Count ? DateTime.Parse(cEndList[i]) : DateTime.MinValue
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking for existing deals: {ex.Message}",
                    "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            return deals;
        }

        private enum DealOption
        {
            Proceed,
            Decline,
            AddToExisting
        }

        private class DealInfo
        {
            public int Id { get; set; }
            public string Agency { get; set; }
            public string Advertiser { get; set; }
            public DateTime CampaignStart { get; set; }
            public DateTime CampaignEnd { get; set; }
        }

        private (DealOption, int) ShowDealOptionsDialog(string message, List<DealInfo> deals)
        {
            using (Form dialogForm = new Form())
            {
                dialogForm.Text = "Existing Deal Found";
                dialogForm.Size = new System.Drawing.Size(480, 320);
                dialogForm.StartPosition = FormStartPosition.CenterParent;
                dialogForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialogForm.MaximizeBox = false;
                dialogForm.MinimizeBox = false;

                Label lblMessage = new Label
                {
                    Text = message,
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(430, 120),
                    AutoSize = false
                };

                Label lblSelectDeal = new Label
                {
                    Text = "Select deal to add to:",
                    Location = new System.Drawing.Point(20, 145),
                    AutoSize = true,
                    Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
                };

                ComboBox cmbDeals = new ComboBox
                {
                    Location = new System.Drawing.Point(20, 165),
                    Size = new System.Drawing.Size(430, 25),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };

                // Populate deals dropdown
                foreach (var deal in deals)
                {
                    cmbDeals.Items.Add($"Deal #{deal.Id}: {deal.Agency} ({deal.CampaignStart:yyyy-MM-dd} to {deal.CampaignEnd:yyyy-MM-dd})");
                }
                if (cmbDeals.Items.Count > 0)
                    cmbDeals.SelectedIndex = 0;

                Button btnProceed = new Button
                {
                    Text = "Create New Deal",
                    Location = new System.Drawing.Point(20, 220),
                    Size = new System.Drawing.Size(130, 35),
                    BackColor = System.Drawing.Color.FromArgb(0, 150, 136),
                    ForeColor = System.Drawing.Color.White,
                    FlatStyle = FlatStyle.Flat
                };

                Button btnDecline = new Button
                {
                    Text = "Decline",
                    Location = new System.Drawing.Point(170, 220),
                    Size = new System.Drawing.Size(130, 35),
                    BackColor = System.Drawing.Color.FromArgb(244, 67, 54),
                    ForeColor = System.Drawing.Color.White,
                    FlatStyle = FlatStyle.Flat
                };

                Button btnAddToExisting = new Button
                {
                    Text = "Add to Existing",
                    Location = new System.Drawing.Point(320, 220),
                    Size = new System.Drawing.Size(130, 35),
                    BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                    ForeColor = System.Drawing.Color.White,
                    FlatStyle = FlatStyle.Flat
                };

                DealOption result = DealOption.Decline;
                int selectedDealId = 0;

                btnProceed.Click += (s, ev) => { result = DealOption.Proceed; dialogForm.Close(); };
                btnDecline.Click += (s, ev) => { result = DealOption.Decline; dialogForm.Close(); };
                btnAddToExisting.Click += (s, ev) =>
                {
                    result = DealOption.AddToExisting;
                    if (cmbDeals.SelectedIndex >= 0 && cmbDeals.SelectedIndex < deals.Count)
                    {
                        selectedDealId = deals[cmbDeals.SelectedIndex].Id;
                    }
                    dialogForm.Close();
                };

                dialogForm.Controls.Add(lblMessage);
                dialogForm.Controls.Add(lblSelectDeal);
                dialogForm.Controls.Add(cmbDeals);
                dialogForm.Controls.Add(btnProceed);
                dialogForm.Controls.Add(btnDecline);
                dialogForm.Controls.Add(btnAddToExisting);

                dialogForm.ShowDialog();

                return (result, selectedDealId);
            }
        }

        private int CreateNewDeal(string agency, string advertiser, DateTime campaignStart, DateTime campaignEnd)
        {
            var db = new DatabaseHelper();

            string startDate = campaignStart.ToString("yyyy-MM-dd");
            string endDate = campaignEnd.ToString("yyyy-MM-dd");
            string created = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Escape apostrophes for SQL
            string escapedAgency = agency?.Replace("'", "''") ?? "";
            string escapedAdvertiser = advertiser?.Replace("'", "''") ?? "";

            // Insert new deal and get the ID in single connection
            string insertQuery = $"INSERT INTO u333577897_dbofhash.bookingdetails " +
                $"(agency, advertiser, c_start, c_end, created) VALUES " +
                $"('{escapedAgency}', '{escapedAdvertiser}', '{startDate}', '{endDate}', '{created}')";

            return db.ExecuteInsertAndGetId(insertQuery);
        }

        private class PeriodInfo
        {
            public int Id { get; set; }
            public string PeriodName { get; set; }
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
            // Effective dates within campaign range
            public DateTime EffectiveStart { get; set; }
            public DateTime EffectiveEnd { get; set; }
        }

        private List<PeriodInfo> GetCampaignPeriods(DateTime campaignStart, DateTime campaignEnd)
        {
            var periods = new List<PeriodInfo>();
            var db = new DatabaseHelper();

            try
            {
                string cStart = campaignStart.ToString("yyyy-MM-dd");
                string cEnd = campaignEnd.ToString("yyyy-MM-dd");

                // Find all periods that overlap with the campaign
                var idList = db.GetComboBoxList(
                    $"SELECT id FROM u333577897_dbofhash.currency_rate " +
                    $"WHERE start_date <= '{cEnd}' AND end_date >= '{cStart}' " +
                    $"ORDER BY start_date");

                var periodList = db.GetComboBoxList(
                    $"SELECT period FROM u333577897_dbofhash.currency_rate " +
                    $"WHERE start_date <= '{cEnd}' AND end_date >= '{cStart}' " +
                    $"ORDER BY start_date");

                var startDateList = db.GetComboBoxList(
                    $"SELECT start_date FROM u333577897_dbofhash.currency_rate " +
                    $"WHERE start_date <= '{cEnd}' AND end_date >= '{cStart}' " +
                    $"ORDER BY start_date");

                var endDateList = db.GetComboBoxList(
                    $"SELECT end_date FROM u333577897_dbofhash.currency_rate " +
                    $"WHERE start_date <= '{cEnd}' AND end_date >= '{cStart}' " +
                    $"ORDER BY start_date");

                for (int i = 0; i < idList.Count; i++)
                {
                    DateTime periodStart = i < startDateList.Count ? DateTime.Parse(startDateList[i]) : DateTime.MinValue;
                    DateTime periodEnd = i < endDateList.Count ? DateTime.Parse(endDateList[i]) : DateTime.MinValue;

                    // Calculate effective dates (intersection of period and campaign)
                    DateTime effectiveStart = periodStart > campaignStart ? periodStart : campaignStart;
                    DateTime effectiveEnd = periodEnd < campaignEnd ? periodEnd : campaignEnd;

                    periods.Add(new PeriodInfo
                    {
                        Id = int.TryParse(idList[i], out int id) ? id : 0,
                        PeriodName = i < periodList.Count ? periodList[i] : "",
                        StartDate = periodStart,
                        EndDate = periodEnd,
                        EffectiveStart = effectiveStart,
                        EffectiveEnd = effectiveEnd
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting campaign periods: {ex.Message}",
                    "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return periods;
        }

        private class BookingGroup
        {
            public int Duration { get; set; }
            public string PeriodName { get; set; }
            public DateTime EffectiveStart { get; set; }
            public DateTime EffectiveEnd { get; set; }
            public DateTime ActualStartDate { get; set; } = DateTime.MaxValue;
            public DateTime ActualEndDate { get; set; } = DateTime.MinValue;
            public int TotalSpots { get; set; }
            public decimal TotalAmount { get; set; }
            public int TotalSpace { get; set; } // Dur × Spots
        }

        private List<BookingGroup> GroupBookingsByDurationAndPeriod(List<PeriodInfo> periods)
        {
            var groups = new List<BookingGroup>();

            // Iterate through each row in the grid (excluding Total row)
            for (int rowIndex = 0; rowIndex < dgvCalendar.Rows.Count - 1; rowIndex++)
            {
                int duration = int.TryParse(dgvCalendar["Dur", rowIndex].Value?.ToString(), out int d) ? d : 0;

                // Get Unit Price KWD for this row
                decimal unitPriceKWD = 0;
                if (dgvCalendar.Columns.Contains("Unit Price KWD"))
                {
                    var priceValue = dgvCalendar["Unit Price KWD", rowIndex].Value?.ToString() ?? "";
                    decimal.TryParse(priceValue.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out unitPriceKWD);
                }

                // Iterate through date columns in the campaign range
                for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
                {
                    string columnName = GetDateColumnName(date);

                    int spots = 0;
                    var cellValue = dgvCalendar[columnName, rowIndex].Value;
                    if (cellValue != null && int.TryParse(cellValue.ToString(), out int s))
                    {
                        spots = s;
                    }

                    if (spots > 0)
                    {
                        // Find which period this date belongs to
                        foreach (var period in periods)
                        {
                            if (date >= period.EffectiveStart && date <= period.EffectiveEnd)
                            {
                                // Find or create group
                                var group = groups.FirstOrDefault(g =>
                                    g.Duration == duration &&
                                    g.PeriodName == period.PeriodName);

                                if (group == null)
                                {
                                    group = new BookingGroup
                                    {
                                        Duration = duration,
                                        PeriodName = period.PeriodName,
                                        EffectiveStart = period.EffectiveStart,
                                        EffectiveEnd = period.EffectiveEnd,
                                        ActualStartDate = date,
                                        ActualEndDate = date,
                                        TotalSpots = 0,
                                        TotalAmount = 0,
                                        TotalSpace = 0
                                    };
                                    groups.Add(group);
                                }

                                // Track actual date range of spots in this group
                                if (date < group.ActualStartDate)
                                    group.ActualStartDate = date;
                                if (date > group.ActualEndDate)
                                    group.ActualEndDate = date;

                                group.TotalSpots += spots;
                                group.TotalAmount += unitPriceKWD * spots;
                                group.TotalSpace += duration * spots;
                                break; // Date can only belong to one period
                            }
                        }
                    }
                }
            }

            // Sort groups by duration then by actual start date
            return groups.OrderBy(g => g.Duration).ThenBy(g => g.ActualStartDate).ToList();
        }

        private void ShowGroupedBookingsTable(List<BookingGroup> groups, List<PeriodInfo> periods)
        {
            Form tableForm = new Form
            {
                Text = "Grouped Bookings by Duration and Period",
                Size = new System.Drawing.Size(800, 450),
                StartPosition = FormStartPosition.CenterParent,
                MinimumSize = new System.Drawing.Size(600, 350)
            };

            // Create DataGridView
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            // Add columns
            dgv.Columns.Add("Duration", "Duration");
            dgv.Columns.Add("Period", "Period");
            dgv.Columns.Add("PeriodRange", "Period Range");
            dgv.Columns.Add("Spots", "Spots");
            dgv.Columns.Add("Amount", "Amount");

            // Set column widths
            dgv.Columns["Duration"].Width = 80;
            dgv.Columns["Period"].Width = 120;
            dgv.Columns["PeriodRange"].Width = 180;
            dgv.Columns["Spots"].Width = 70;
            dgv.Columns["Amount"].Width = 100;

            // Right-align numeric columns
            dgv.Columns["Spots"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns["Amount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            // Add data rows
            foreach (var group in groups)
            {
                dgv.Rows.Add(
                    $"{group.Duration} sec",
                    group.PeriodName,
                    $"{group.ActualStartDate:yyyy-MM-dd} to {group.ActualEndDate:yyyy-MM-dd}",
                    group.TotalSpots,
                    group.TotalAmount.ToString("N3")
                );
            }

            // Add total row
            int grandTotalSpots = groups.Sum(g => g.TotalSpots);
            decimal grandTotalAmount = groups.Sum(g => g.TotalAmount);
            int totalRowIndex = dgv.Rows.Add("", "", "GRAND TOTAL", grandTotalSpots, grandTotalAmount.ToString("N3"));
            dgv.Rows[totalRowIndex].DefaultCellStyle.Font = new System.Drawing.Font(dgv.Font, System.Drawing.FontStyle.Bold);
            dgv.Rows[totalRowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;

            // Style header
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgv.EnableHeadersVisualStyles = false;

            // Close button
            Button btnClose = new Button
            {
                Text = "Close",
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = System.Drawing.Color.FromArgb(100, 100, 100),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.Click += (s, ev) => tableForm.Close();

            tableForm.Controls.Add(dgv);
            tableForm.Controls.Add(btnClose);

            tableForm.ShowDialog();
        }

        private void ShowBookingReport(List<BookingGroup> groups, List<PeriodInfo> periods)
        {
            Form reportForm = new Form
            {
                Text = $"Booking Report - Deal #{currentDealNumber}",
                Size = new System.Drawing.Size(800, 500),
                StartPosition = FormStartPosition.CenterParent,
                MinimumSize = new System.Drawing.Size(600, 400)
            };

            // Create report text
            StringBuilder report = new StringBuilder();
            report.AppendLine($"BOOKING REPORT");
            report.AppendLine($"==============");
            report.AppendLine($"Deal Number: {currentDealNumber}");
            report.AppendLine($"Agency: {cmbAgency.Text}");
            report.AppendLine($"Advertiser: {cmbAdvertiser.Text}");
            report.AppendLine($"Product: {cmbProduct.Text}");
            report.AppendLine($"Campaign: {dtpCampaignStart.Value:yyyy-MM-dd} to {dtpCampaignEnd.Value:yyyy-MM-dd}");
            report.AppendLine();

            report.AppendLine($"PERIODS COVERED:");
            report.AppendLine($"----------------");
            foreach (var period in periods)
            {
                report.AppendLine($"  {period.PeriodName}: {period.EffectiveStart:yyyy-MM-dd} to {period.EffectiveEnd:yyyy-MM-dd}");
            }
            report.AppendLine();

            report.AppendLine($"BOOKINGS GROUPED BY DURATION AND PERIOD:");
            report.AppendLine($"-----------------------------------------");
            report.AppendLine();
            report.AppendLine($"{"Duration",-12} {"Period",-20} {"Date Range",-25} {"Spots",-10} {"Amount",-15}");
            report.AppendLine($"{new string('-', 85)}");

            foreach (var group in groups)
            {
                string dateRange = $"{group.ActualStartDate:yyyy-MM-dd} to {group.ActualEndDate:yyyy-MM-dd}";
                report.AppendLine($"{group.Duration + " sec",-12} {group.PeriodName,-20} {dateRange,-25} {group.TotalSpots,-10} {group.TotalAmount.ToString("N3"),-15}");
            }
            report.AppendLine();

            // Summary by Duration
            report.AppendLine($"SUMMARY BY DURATION:");
            report.AppendLine($"--------------------");
            var summaryByDuration = groups.GroupBy(g => g.Duration)
                .Select(g => new { Duration = g.Key, TotalSpots = g.Sum(x => x.TotalSpots), TotalAmount = g.Sum(x => x.TotalAmount) });

            foreach (var summary in summaryByDuration.OrderBy(s => s.Duration))
            {
                report.AppendLine($"  Duration {summary.Duration} sec: {summary.TotalSpots} spots, Amount: {summary.TotalAmount:N3}");
            }
            report.AppendLine();

            // Summary by Period
            report.AppendLine($"SUMMARY BY PERIOD:");
            report.AppendLine($"------------------");
            var summaryByPeriod = groups.GroupBy(g => g.PeriodName)
                .Select(g => new { Period = g.Key, TotalSpots = g.Sum(x => x.TotalSpots), TotalAmount = g.Sum(x => x.TotalAmount) });

            foreach (var summary in summaryByPeriod)
            {
                report.AppendLine($"  {summary.Period}: {summary.TotalSpots} spots, Amount: {summary.TotalAmount:N3}");
            }
            report.AppendLine();
            report.AppendLine($"  GRAND TOTAL: {groups.Sum(g => g.TotalSpots)} spots, Amount: {groups.Sum(g => g.TotalAmount):N3}");

            // TextBox to display report
            TextBox txtReport = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Consolas", 10F),
                Text = report.ToString(),
                WordWrap = false
            };

            // Close button
            Button btnClose = new Button
            {
                Text = "Close",
                Dock = DockStyle.Bottom,
                Height = 40,
                BackColor = System.Drawing.Color.FromArgb(100, 100, 100),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClose.Click += (s, ev) => reportForm.Close();

            reportForm.Controls.Add(txtReport);
            reportForm.Controls.Add(btnClose);

            reportForm.Show();
        }

        // Class to hold existing booking info from bookingdetails2
        private class ExistingBookingInfo
        {
            public int Id { get; set; }
            public int Ord { get; set; }
            public int Schedule { get; set; }
            public int Dur { get; set; }
            public string SalesPeriod { get; set; }
            public DateTime DStart { get; set; }
            public DateTime DEnd { get; set; }
            public int Spots { get; set; }
            public int Space { get; set; }
            public decimal NetAmount { get; set; }
        }

        private List<ExistingBookingInfo> GetExistingBookings(int dealNumber)
        {
            var bookings = new List<ExistingBookingInfo>();
            var db = new DatabaseHelper();

            try
            {
                var idList = db.GetComboBoxList(
                    $"SELECT id FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var ordList = db.GetComboBoxList(
                    $"SELECT ord FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var scheduleList = db.GetComboBoxList(
                    $"SELECT schedule FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var durList = db.GetComboBoxList(
                    $"SELECT dur FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var periodList = db.GetComboBoxList(
                    $"SELECT salesPeriod FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var dStartList = db.GetComboBoxList(
                    $"SELECT d_start FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var dEndList = db.GetComboBoxList(
                    $"SELECT d_end FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var spotsList = db.GetComboBoxList(
                    $"SELECT spots FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var spaceList = db.GetComboBoxList(
                    $"SELECT space FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
                var amountList = db.GetComboBoxList(
                    $"SELECT netAmount FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");

                for (int i = 0; i < idList.Count; i++)
                {
                    bookings.Add(new ExistingBookingInfo
                    {
                        Id = int.TryParse(idList[i], out int id) ? id : 0,
                        Ord = int.TryParse(ordList.Count > i ? ordList[i] : "0", out int ord) ? ord : 0,
                        Schedule = int.TryParse(scheduleList.Count > i ? scheduleList[i] : "0", out int sch) ? sch : 0,
                        Dur = int.TryParse(durList[i], out int dur) ? dur : 0,
                        SalesPeriod = periodList.Count > i ? periodList[i] : "",
                        DStart = DateTime.TryParse(dStartList.Count > i ? dStartList[i] : "", out DateTime ds) ? ds : DateTime.MinValue,
                        DEnd = DateTime.TryParse(dEndList.Count > i ? dEndList[i] : "", out DateTime de) ? de : DateTime.MinValue,
                        Spots = int.TryParse(spotsList.Count > i ? spotsList[i] : "0", out int sp) ? sp : 0,
                        Space = int.TryParse(spaceList.Count > i ? spaceList[i] : "0", out int spc) ? spc : 0,
                        NetAmount = decimal.TryParse(amountList.Count > i ? amountList[i] : "0", NumberStyles.Any, CultureInfo.InvariantCulture, out decimal amt) ? amt : 0
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fetching existing bookings: {ex.Message}",
                    "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return bookings;
        }

        private bool ShowBookingComparisonDialog(List<ExistingBookingInfo> existingBookings, List<BookingGroup> newBookings)
        {
            Form compareForm = new Form
            {
                Text = $"Compare Bookings - Deal #{currentDealNumber}",
                Size = new System.Drawing.Size(900, 600),
                StartPosition = FormStartPosition.CenterParent,
                MinimumSize = new System.Drawing.Size(700, 400)
            };

            // Panel for existing bookings
            Label lblExisting = new Label
            {
                Text = "Existing Bookings in Database:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold)
            };

            DataGridView dgvExisting = new DataGridView
            {
                Dock = DockStyle.Top,
                Height = 200,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.FromArgb(45, 45, 45),
                ForeColor = System.Drawing.Color.White
            };

            dgvExisting.Columns.Add("Dur", "Dur");
            dgvExisting.Columns.Add("Period", "Period");
            dgvExisting.Columns.Add("DateRange", "Date Range");
            dgvExisting.Columns.Add("Spots", "Spots");
            dgvExisting.Columns.Add("Space", "Space");
            dgvExisting.Columns.Add("Amount", "Amount");

            foreach (var booking in existingBookings)
            {
                dgvExisting.Rows.Add(
                    booking.Dur,
                    booking.SalesPeriod,
                    $"{booking.DStart:yyyy-MM-dd} to {booking.DEnd:yyyy-MM-dd}",
                    booking.Spots,
                    booking.Space,
                    booking.NetAmount.ToString("N3")
                );
            }

            // Panel for new bookings
            Label lblNew = new Label
            {
                Text = "New Bookings to Insert:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold)
            };

            DataGridView dgvNew = new DataGridView
            {
                Dock = DockStyle.Top,
                Height = 200,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.FromArgb(45, 45, 45),
                ForeColor = System.Drawing.Color.White
            };

            dgvNew.Columns.Add("Dur", "Dur");
            dgvNew.Columns.Add("Period", "Period");
            dgvNew.Columns.Add("DateRange", "Date Range");
            dgvNew.Columns.Add("Spots", "Spots");
            dgvNew.Columns.Add("Space", "Space");
            dgvNew.Columns.Add("Amount", "Amount");

            foreach (var booking in newBookings)
            {
                dgvNew.Rows.Add(
                    booking.Duration,
                    booking.PeriodName,
                    $"{booking.ActualStartDate:yyyy-MM-dd} to {booking.ActualEndDate:yyyy-MM-dd}",
                    booking.TotalSpots,
                    booking.TotalSpace,
                    booking.TotalAmount.ToString("N3")
                );
            }

            // Buttons panel
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            Button btnReplace = new Button
            {
                Text = "Replace Existing",
                Width = 150,
                Height = 35,
                Left = 200,
                Top = 8,
                BackColor = System.Drawing.Color.FromArgb(200, 80, 80),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };

            Button btnCancel = new Button
            {
                Text = "Cancel",
                Width = 150,
                Height = 35,
                Left = 370,
                Top = 8,
                BackColor = System.Drawing.Color.FromArgb(100, 100, 100),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };

            bool replaceConfirmed = false;

            btnReplace.Click += (s, ev) =>
            {
                replaceConfirmed = true;
                compareForm.Close();
            };

            btnCancel.Click += (s, ev) =>
            {
                replaceConfirmed = false;
                compareForm.Close();
            };

            buttonPanel.Controls.Add(btnReplace);
            buttonPanel.Controls.Add(btnCancel);

            // Add controls in correct order (bottom to top for docking)
            compareForm.Controls.Add(buttonPanel);
            compareForm.Controls.Add(dgvNew);
            compareForm.Controls.Add(lblNew);
            compareForm.Controls.Add(dgvExisting);
            compareForm.Controls.Add(lblExisting);

            compareForm.ShowDialog();

            return replaceConfirmed;
        }

        private (int ord, int schedule) GetNextOrdAndSchedule()
        {
            var db = new DatabaseHelper();

            var ordList = db.GetComboBoxList("SELECT IFNULL(MAX(ord), 0) + 1 FROM u333577897_dbofhash.bookingdetails2");
            var scheduleList = db.GetComboBoxList("SELECT IFNULL(MAX(schedule), 0) + 1 FROM u333577897_dbofhash.bookingdetails2");

            int ord = ordList.Count > 0 && int.TryParse(ordList[0], out int o) ? o : 1;
            int schedule = scheduleList.Count > 0 && int.TryParse(scheduleList[0], out int s) ? s : 1;

            return (ord, schedule);
        }

        private void DeleteExistingBookings(int dealNumber, List<ExistingBookingInfo> existingBookings)
        {
            var db = new DatabaseHelper();

            // Delete spots from spot_book by ord/schedule
            foreach (var booking in existingBookings)
            {
                db.ExecuteNonQuery($"DELETE FROM u333577897_dbofhash.spot_book WHERE ord = {booking.Ord} AND schedule = {booking.Schedule}");
            }

            // Delete from bookingdetails2
            db.ExecuteNonQuery($"DELETE FROM u333577897_dbofhash.bookingdetails2 WHERE deal = {dealNumber}");
        }

        private void InsertBookings(List<BookingGroup> bookings, int dealNumber, string product, string agency, string advertiser, List<PeriodInfo> periods)
        {
            var db = new DatabaseHelper();
            var (ord, schedule) = GetNextOrdAndSchedule();

            foreach (var booking in bookings)
            {
                string dStart = booking.ActualStartDate.ToString("yyyy-MM-dd");
                string dEnd = booking.ActualEndDate.ToString("yyyy-MM-dd");
                string year = booking.ActualStartDate.ToString("yyyy");
                string month = booking.ActualStartDate.ToString("MMM");

                string insertQuery = $@"INSERT INTO u333577897_dbofhash.bookingdetails2
                    (deal, vehicle, ord, schedule, agencyorder, agencyschedule, market, type,
                     d_start, d_end, year, month, salesPeriod, s_ref, product, ptype, status,
                     spots, dur, space, currency, paymentType, netGross, discount, netAmount, avr, file, remark)
                    VALUES
                    ({dealNumber}, 'Alrai TV', {ord}, {schedule}, '{ord}', '{schedule}', 'KUWAIT', 'COMMERCIAL',
                     '{dStart}', '{dEnd}', '{year}', '{month}', '{booking.PeriodName}', 'SALES TEAM', '{product.Replace("'", "''")}', 'NA', 'OK',
                     {booking.TotalSpots}, {booking.Duration}, {booking.TotalSpace}, 'KWD', 'CREDIT',
                     '{booking.TotalAmount.ToString("0.000", CultureInfo.InvariantCulture)}', '0',
                     '{booking.TotalAmount.ToString("0.000", CultureInfo.InvariantCulture)}', 1, '', '')";

                db.ExecuteNonQuery(insertQuery);

                // Insert individual spots into spot_book
                InsertSpotsForBooking(booking, ord, schedule, product, agency, advertiser, periods);

                ord++;
                schedule++;
            }
        }

        private void InsertSpotsForBooking(BookingGroup booking, int ord, int schedule, string product, string agency, string advertiser, List<PeriodInfo> periods)
        {
            var db = new DatabaseHelper();

            // Find the period that matches this booking
            var period = periods.FirstOrDefault(p => p.PeriodName == booking.PeriodName);
            if (period == null) return;

            // Loop through grid rows to find rows with matching duration
            for (int rowIndex = 0; rowIndex < dgvCalendar.Rows.Count - 1; rowIndex++)
            {
                int rowDuration = int.TryParse(dgvCalendar["Dur", rowIndex].Value?.ToString(), out int d) ? d : 0;

                // Skip if duration doesn't match
                if (rowDuration != booking.Duration) continue;

                // Get row data
                string programme = dgvCalendar["Programme", rowIndex].Value?.ToString() ?? "";
                string time = dgvCalendar["Time", rowIndex].Value?.ToString() ?? "";
                string salesType = dgvCalendar.Columns.Contains("Sales Type")
                    ? (dgvCalendar["Sales Type", rowIndex].Value?.ToString() ?? "NA")
                    : "NA";

                // Get Unit Price KWD for this row
                decimal unitPrice = 0;
                if (dgvCalendar.Columns.Contains("Unit Price KWD"))
                {
                    var priceValue = dgvCalendar["Unit Price KWD", rowIndex].Value?.ToString() ?? "";
                    decimal.TryParse(priceValue.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out unitPrice);
                }

                // Loop through dates within the booking's actual date range
                for (DateTime date = booking.ActualStartDate; date <= booking.ActualEndDate; date = date.AddDays(1))
                {
                    // Check if this date is within the period's effective range
                    if (date < period.EffectiveStart || date > period.EffectiveEnd) continue;

                    string columnName = GetDateColumnName(date);
                    if (!dgvCalendar.Columns.Contains(columnName)) continue;

                    // Get spot count for this cell
                    int spots = 0;
                    var cellValue = dgvCalendar[columnName, rowIndex].Value;
                    if (cellValue != null && int.TryParse(cellValue.ToString(), out int s))
                    {
                        spots = s;
                    }

                    // Insert one row per spot
                    for (int i = 0; i < spots; i++)
                    {
                        string transDate = date.ToString("yyyy-MM-dd");
                        string dayName = date.ToString("ddd");

                        string spotQuery = $@"INSERT INTO u333577897_dbofhash.spot_book
                            (Rai, Booking_Status, Control, BR_No, PGRA, PrgNm, Payment, Price, PrgStartTime, SpotTime,
                             BR_Type, REV_Type, BreakNo, OrderinBreak, SponsorOrder, NoofSpots, TransDate, Day,
                             ProductType, Product, Version, Agency, Advertiser, AdvDur, ClipID, ord, schedule)
                            VALUES
                            (0, 'OK', 'NA', 0, 'NA', '{programme.Replace("'", "''")}', 'NA',
                             {unitPrice.ToString("0.000", CultureInfo.InvariantCulture)}, '{time}', 'NA',
                             '{salesType.Replace("'", "''")}', 'NA', 0, 0, 'NA', 1, '{transDate}', '{dayName}',
                             'NA', '{product.Replace("'", "''")}', '', '{agency.Replace("'", "''")}',
                             '{advertiser.Replace("'", "''")}', {rowDuration}, 'NA', {ord}, {schedule})";

                        db.ExecuteNonQuery(spotQuery);
                    }
                }
            }
        }

        private void ProcessBookingsForDeal(List<BookingGroup> groupedBookings, int dealNumber, string product, string agency, string advertiser, List<PeriodInfo> periods)
        {
            // Check for existing bookings
            var existingBookings = GetExistingBookings(dealNumber);

            if (existingBookings.Count > 0)
            {
                // Show comparison dialog
                bool replace = ShowBookingComparisonDialog(existingBookings, groupedBookings);

                if (replace)
                {
                    // Delete existing and insert new
                    DeleteExistingBookings(dealNumber, existingBookings);
                    InsertBookings(groupedBookings, dealNumber, product, agency, advertiser, periods);
                    MessageBox.Show($"Replaced {existingBookings.Count} existing bookings with {groupedBookings.Count} new bookings.",
                        "Bookings Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Operation cancelled. No changes made to bookings.",
                        "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                // No existing bookings - insert new ones
                InsertBookings(groupedBookings, dealNumber, product, agency, advertiser, periods);
                MessageBox.Show($"Inserted {groupedBookings.Count} new bookings for Deal #{dealNumber}.",
                    "Bookings Inserted", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void DgvCalendar_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (isUpdating || e.RowIndex < 0 || e.ColumnIndex < 0 || calendarDataTable == null)
                return;

            string columnName = dgvCalendar.Columns[e.ColumnIndex].Name;

            // Recalculate totals when date cell values change
            if (columnName.Contains("\n") && char.IsDigit(columnName.Split('\n')[1][0]))
            {
                RecalculateRowTotal(e.RowIndex);
                RecalculateColumnTotals();
                RecalculatePricing();
            }
            // Trigger pricing recalculation when F/P, Dur, or Ratio changes
            else if (columnName == "F/P" || columnName == "Dur" || columnName == "Ratio")
            {
                RecalculatePricing();
            }
        }

        private void RecalculateRowTotal(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= calendarDataTable.Rows.Count - 1)
                return;

            DataRow row = calendarDataTable.Rows[rowIndex];
            int total = 0;

            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);

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
            if (calendarDataTable == null)
                return;

            DataRow totalRow = calendarDataTable.Rows[calendarDataTable.Rows.Count - 1];
            int grandTotal = 0;

            for (DateTime date = campaignStartDate; date <= campaignEndDate; date = date.AddDays(1))
            {
                string columnName = GetDateColumnName(date);

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

            // Update the Total Spots label
            if (currentBookingOrder != null)
            {
                currentBookingOrder.TotalSpots = grandTotal;
                lblTotalSpots.Text = $"Total Spots: {grandTotal}";
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            Application.Exit();
        }
    }
}
