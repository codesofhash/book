using System;
using System.IO;
using System.Windows.Forms;
using DotNetEnv;
using Newtonsoft.Json;

namespace CSharpFlexGrid
{
    public partial class MainMenuForm : Form
    {
        private string jsonOutputPath;

        public MainMenuForm()
        {
            InitializeComponent();
            LoadEnvironmentVariables();
        }

        private void LoadEnvironmentVariables()
        {
            try
            {
                // Load .env file from the application directory
                string envPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".env");
                if (File.Exists(envPath))
                {
                    Env.Load(envPath);
                    jsonOutputPath = Environment.GetEnvironmentVariable("JSON_OUTPUT_PATH");

                    // Create output directory if it doesn't exist
                    if (!string.IsNullOrEmpty(jsonOutputPath) && !Directory.Exists(jsonOutputPath))
                    {
                        Directory.CreateDirectory(jsonOutputPath);
                    }
                }
                else
                {
                    MessageBox.Show("Warning: .env file not found. Using default output path.",
                        "Configuration", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            // Open the BookingManagerForm
            BookingManagerForm bookingForm = new BookingManagerForm();
            bookingForm.Show();
            this.Hide();
        }

        private void ProcessExcelFile(string filePath)
        {
            try
            {
                // Show processing message
                Cursor = Cursors.WaitCursor;
                lblStatus.Text = "Processing file...";
                lblStatus.Visible = true;
                Application.DoEvents();

                // Process the Excel file
                BookingOrderReader reader = new BookingOrderReader();
                string jsonResult = reader.ProcessBookingOrder(filePath);

                // Check if result contains an error
                if (jsonResult.Contains("\"error\""))
                {
                    MessageBox.Show("Failed to process the booking order file.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lblStatus.Text = "Processing failed.";
                    return;
                }

                // Deserialize JSON to BookingOrder object
                var bookingOrder = JsonConvert.DeserializeObject<BookingOrder>(jsonResult);

                if (bookingOrder == null)
                {
                    MessageBox.Show("Failed to deserialize the booking order data.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Save JSON to configured path
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string jsonFileName = $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
                string jsonFilePath = Path.Combine(jsonOutputPath, jsonFileName);

                File.WriteAllText(jsonFilePath, jsonResult);

                lblStatus.Text = $"Saved: {jsonFileName}";

                // Open the calendar grid form
                CalendarGridForm calendarForm = new CalendarGridForm(bookingOrder, jsonFilePath);
                calendarForm.Show();
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing file: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Processing failed.";
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void btnPlaceholder_Click(object sender, EventArgs e)
        {
            // Placeholder for future functionality
            MessageBox.Show("This feature is not yet implemented.",
                "Coming Soon", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            // Exit the application when main menu is closed
            Application.Exit();
        }
    }
}
