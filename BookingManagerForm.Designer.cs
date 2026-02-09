namespace CSharpFlexGrid
{
    partial class BookingManagerForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Panel panelInfo;
        private System.Windows.Forms.Panel panelActions;
        private System.Windows.Forms.Panel panelGrid;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblAgencyLabel;
        private System.Windows.Forms.Label lblAdvertiserLabel;
        private System.Windows.Forms.Label lblProductLabel;
        private System.Windows.Forms.Label lblCampaignStartLabel;
        private System.Windows.Forms.Label lblCampaignEndLabel;
        private System.Windows.Forms.Label lblTotalSpots;
        private System.Windows.Forms.Label lblPackageCostLabel;
        private System.Windows.Forms.TextBox txtPackageCost;
        private EditableAutoCompleteComboBox cmbAgency;
        private EditableAutoCompleteComboBox cmbAdvertiser;
        private EditableAutoCompleteComboBox cmbProduct;
        private System.Windows.Forms.DateTimePicker dtpCampaignStart;
        private System.Windows.Forms.DateTimePicker dtpCampaignEnd;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnCreateNew;
        private System.Windows.Forms.Button btnUploadCSV;
        private System.Windows.Forms.Button btnEditBooking;
        private System.Windows.Forms.Button btnBack;
        private CSharpFlexGrid.CustomDataGridView dgvCalendar;
        private System.Windows.Forms.Label lblRecordCount;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.Panel panelGridMode;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.panelInfo = new System.Windows.Forms.Panel();
            this.dtpCampaignEnd = new System.Windows.Forms.DateTimePicker();
            this.dtpCampaignStart = new System.Windows.Forms.DateTimePicker();
            this.cmbProduct = new EditableAutoCompleteComboBox();
            this.cmbAdvertiser = new EditableAutoCompleteComboBox();
            this.cmbAgency = new EditableAutoCompleteComboBox();
            this.lblCampaignEndLabel = new System.Windows.Forms.Label();
            this.lblCampaignStartLabel = new System.Windows.Forms.Label();
            this.lblProductLabel = new System.Windows.Forms.Label();
            this.lblAdvertiserLabel = new System.Windows.Forms.Label();
            this.lblAgencyLabel = new System.Windows.Forms.Label();
            this.lblTotalSpots = new System.Windows.Forms.Label();
            this.lblPackageCostLabel = new System.Windows.Forms.Label();
            this.txtPackageCost = new System.Windows.Forms.TextBox();
            this.panelActions = new System.Windows.Forms.Panel();
            this.btnEditBooking = new System.Windows.Forms.Button();
            this.btnUploadCSV = new System.Windows.Forms.Button();
            this.btnCreateNew = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.panelGridMode = new System.Windows.Forms.Panel();
            this.panelGrid = new System.Windows.Forms.Panel();
            this.lblRecordCount = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.dgvCalendar = new CSharpFlexGrid.CustomDataGridView();
            this.panelHeader.SuspendLayout();
            this.panelInfo.SuspendLayout();
            this.panelGridMode.SuspendLayout();
            this.panelActions.SuspendLayout();
            this.panelGrid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalendar)).BeginInit();
            this.SuspendLayout();
            //
            // panelHeader
            //
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(45)))), ((int)(((byte)(48)))));
            this.panelHeader.Controls.Add(this.lblTitle);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(1400, 60);
            this.panelHeader.TabIndex = 0;
            //
            // lblTitle
            //
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.White;
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(1400, 60);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "FlexGrid - Calendar View";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // panelInfo
            //
            this.panelInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.panelInfo.Controls.Add(this.dtpCampaignEnd);
            this.panelInfo.Controls.Add(this.dtpCampaignStart);
            this.panelInfo.Controls.Add(this.cmbProduct);
            this.panelInfo.Controls.Add(this.cmbAdvertiser);
            this.panelInfo.Controls.Add(this.cmbAgency);
            this.panelInfo.Controls.Add(this.lblCampaignEndLabel);
            this.panelInfo.Controls.Add(this.lblCampaignStartLabel);
            this.panelInfo.Controls.Add(this.lblProductLabel);
            this.panelInfo.Controls.Add(this.lblAdvertiserLabel);
            this.panelInfo.Controls.Add(this.lblAgencyLabel);
            this.panelInfo.Controls.Add(this.lblTotalSpots);
            this.panelInfo.Controls.Add(this.lblPackageCostLabel);
            this.panelInfo.Controls.Add(this.txtPackageCost);
            this.panelInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelInfo.Location = new System.Drawing.Point(0, 60);
            this.panelInfo.Name = "panelInfo";
            this.panelInfo.Padding = new System.Windows.Forms.Padding(20, 10, 20, 10);
            this.panelInfo.Size = new System.Drawing.Size(1400, 100);
            this.panelInfo.TabIndex = 1;
            //
            // dtpCampaignEnd
            //
            this.dtpCampaignEnd.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpCampaignEnd.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCampaignEnd.Location = new System.Drawing.Point(1015, 60);
            this.dtpCampaignEnd.Name = "dtpCampaignEnd";
            this.dtpCampaignEnd.Size = new System.Drawing.Size(150, 23);
            this.dtpCampaignEnd.TabIndex = 9;
            //
            // dtpCampaignStart
            //
            this.dtpCampaignStart.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpCampaignStart.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCampaignStart.Location = new System.Drawing.Point(1015, 15);
            this.dtpCampaignStart.Name = "dtpCampaignStart";
            this.dtpCampaignStart.Size = new System.Drawing.Size(150, 23);
            this.dtpCampaignStart.TabIndex = 8;
            //
            // cmbProduct
            //
            this.cmbProduct.Font = new System.Drawing.Font("Segoe UI", 9F , System.Drawing.FontStyle.Bold);
            this.cmbProduct.FormattingEnabled = true;
            this.cmbProduct.Location = new System.Drawing.Point(660, 40);
            this.cmbProduct.Name = "cmbProduct";
            this.cmbProduct.Size = new System.Drawing.Size(200, 23);
            this.cmbProduct.TabIndex = 7;
            //
            // cmbAdvertiser
            //
            this.cmbProduct.Font = new System.Drawing.Font("Segoe UI", 9F , System.Drawing.FontStyle.Bold);
            this.cmbAdvertiser.FormattingEnabled = true;
            this.cmbAdvertiser.Location = new System.Drawing.Point(400, 40);
            this.cmbAdvertiser.Name = "cmbAdvertiser";
            this.cmbAdvertiser.Size = new System.Drawing.Size(200, 23);
            this.cmbAdvertiser.TabIndex = 6;
            //
            // cmbAgency
            //
            this.cmbAgency.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbAgency.FormattingEnabled = true;
            this.cmbAgency.Location = new System.Drawing.Point(140, 40);
            this.cmbAgency.Name = "cmbAgency";
            this.cmbAgency.Size = new System.Drawing.Size(200, 23);
            this.cmbAgency.TabIndex = 5;
            //
            // lblCampaignEndLabel
            //
            this.lblCampaignEndLabel.AutoSize = true;
            this.lblCampaignEndLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblCampaignEndLabel.Location = new System.Drawing.Point(890, 60);
            this.lblCampaignEndLabel.Name = "lblCampaignEndLabel";
            this.lblCampaignEndLabel.Size = new System.Drawing.Size(91, 15);
            this.lblCampaignEndLabel.TabIndex = 4;
            this.lblCampaignEndLabel.Text = "Campaign End:";
            //
            // lblCampaignStartLabel
            //
            this.lblCampaignStartLabel.AutoSize = true;
            this.lblCampaignStartLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblCampaignStartLabel.Location = new System.Drawing.Point(890, 15);
            this.lblCampaignStartLabel.Name = "lblCampaignStartLabel";
            this.lblCampaignStartLabel.Size = new System.Drawing.Size(99, 15);
            this.lblCampaignStartLabel.TabIndex = 3;
            this.lblCampaignStartLabel.Text = "Campaign Start:";
            //
            // lblProductLabel
            //
            this.lblProductLabel.AutoSize = true;
            this.lblProductLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblProductLabel.Location = new System.Drawing.Point(660, 15);
            this.lblProductLabel.Name = "lblProductLabel";
            this.lblProductLabel.Size = new System.Drawing.Size(54, 15);
            this.lblProductLabel.TabIndex = 2;
            this.lblProductLabel.Text = "Product:";
            //
            // lblAdvertiserLabel
            //
            this.lblAdvertiserLabel.AutoSize = true;
            this.lblAdvertiserLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblAdvertiserLabel.Location = new System.Drawing.Point(400, 15);
            this.lblAdvertiserLabel.Name = "lblAdvertiserLabel";
            this.lblAdvertiserLabel.Size = new System.Drawing.Size(69, 15);
            this.lblAdvertiserLabel.TabIndex = 1;
            this.lblAdvertiserLabel.Text = "Advertiser:";
            //
            // lblAgencyLabel
            //
            this.lblAgencyLabel.AutoSize = true;
            this.lblAgencyLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblAgencyLabel.Location = new System.Drawing.Point(140, 15);
            this.lblAgencyLabel.Name = "lblAgencyLabel";
            this.lblAgencyLabel.Size = new System.Drawing.Size(51, 15);
            this.lblAgencyLabel.TabIndex = 0;
            this.lblAgencyLabel.Text = "Agency:";
            //
            // lblTotalSpots
            //
            this.lblTotalSpots.AutoSize = true;
            this.lblTotalSpots.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblTotalSpots.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.lblTotalSpots.Location = new System.Drawing.Point(1240, 15);
            this.lblTotalSpots.Name = "lblTotalSpots";
            this.lblTotalSpots.Size = new System.Drawing.Size(85, 19);
            this.lblTotalSpots.TabIndex = 10;
            this.lblTotalSpots.Text = "Total Spots: 0";
            //
            // lblPackageCostLabel
            //
            this.lblPackageCostLabel.AutoSize = true;
            this.lblPackageCostLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblPackageCostLabel.Location = new System.Drawing.Point(1240, 60);
            this.lblPackageCostLabel.Name = "lblPackageCostLabel";
            this.lblPackageCostLabel.Size = new System.Drawing.Size(85, 15);
            this.lblPackageCostLabel.TabIndex = 11;
            this.lblPackageCostLabel.Text = "Package Cost:";
            //
            // txtPackageCost
            //
            this.txtPackageCost.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPackageCost.Location = new System.Drawing.Point(1330, 60);
            this.txtPackageCost.Name = "txtPackageCost";
            this.txtPackageCost.Size = new System.Drawing.Size(120, 23);
            this.txtPackageCost.TabIndex = 12;
            this.txtPackageCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            //
            // panelGridMode
            //
            this.panelGridMode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.panelGridMode.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelGridMode.Location = new System.Drawing.Point(0, 160);
            this.panelGridMode.Name = "panelGridMode";
            this.panelGridMode.Padding = new System.Windows.Forms.Padding(20, 5, 20, 5);
            this.panelGridMode.Size = new System.Drawing.Size(1400, 60);
            this.panelGridMode.TabIndex = 4;
            //
            // panelActions
            //
            this.panelActions.BackColor = System.Drawing.Color.White;
            this.panelActions.Controls.Add(this.btnEditBooking);
            this.panelActions.Controls.Add(this.btnUploadCSV);
            this.panelActions.Controls.Add(this.btnCreateNew);
            this.panelActions.Controls.Add(this.btnSelectFile);
            this.panelActions.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelActions.Location = new System.Drawing.Point(0, 220);
            this.panelActions.Name = "panelActions";
            this.panelActions.Padding = new System.Windows.Forms.Padding(20, 10, 20, 10);
            this.panelActions.Size = new System.Drawing.Size(1400, 70);
            this.panelActions.TabIndex = 2;
            //
            // btnEditBooking
            //
            this.btnEditBooking.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(200)))), ((int)(((byte)(200)))));
            this.btnEditBooking.Enabled = false;
            this.btnEditBooking.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEditBooking.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnEditBooking.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.btnEditBooking.Location = new System.Drawing.Point(723, 15);
            this.btnEditBooking.Name = "btnEditBooking";
            this.btnEditBooking.Size = new System.Drawing.Size(220, 40);
            this.btnEditBooking.TabIndex = 3;
            this.btnEditBooking.Text = "Edit Booking Order (Coming Soon)";
            this.btnEditBooking.UseVisualStyleBackColor = false;
            this.btnEditBooking.Click += new System.EventHandler(this.btnEditBooking_Click);
            //
            // btnUploadCSV
            //
            this.btnUploadCSV.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(200)))), ((int)(((byte)(200)))));
            this.btnUploadCSV.Enabled = false;
            this.btnUploadCSV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnUploadCSV.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnUploadCSV.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.btnUploadCSV.Location = new System.Drawing.Point(487, 15);
            this.btnUploadCSV.Name = "btnUploadCSV";
            this.btnUploadCSV.Size = new System.Drawing.Size(220, 40);
            this.btnUploadCSV.TabIndex = 2;
            this.btnUploadCSV.Text = "Upload CSV File (Coming Soon)";
            this.btnUploadCSV.UseVisualStyleBackColor = false;
            this.btnUploadCSV.Click += new System.EventHandler(this.btnUploadCSV_Click);
            //
            // btnCreateNew
            //
            this.btnCreateNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(150)))), ((int)(((byte)(136)))));
            this.btnCreateNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateNew.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnCreateNew.ForeColor = System.Drawing.Color.White;
            this.btnCreateNew.Location = new System.Drawing.Point(251, 15);
            this.btnCreateNew.Name = "btnCreateNew";
            this.btnCreateNew.Size = new System.Drawing.Size(220, 40);
            this.btnCreateNew.TabIndex = 1;
            this.btnCreateNew.Text = "Create New Booking";
            this.btnCreateNew.UseVisualStyleBackColor = false;
            this.btnCreateNew.Click += new System.EventHandler(this.btnCreateNew_Click);
            //
            // btnSelectFile
            //
            this.btnSelectFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.btnSelectFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelectFile.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnSelectFile.ForeColor = System.Drawing.Color.White;
            this.btnSelectFile.Location = new System.Drawing.Point(15, 15);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(220, 40);
            this.btnSelectFile.TabIndex = 0;
            this.btnSelectFile.Text = "Select Booking Order File";
            this.btnSelectFile.UseVisualStyleBackColor = false;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            //
            // panelGrid
            //
            this.panelGrid.Controls.Add(this.lblRecordCount);
            this.panelGrid.Controls.Add(this.btnBack);
            this.panelGrid.Controls.Add(this.btnCalculate);
            this.panelGrid.Controls.Add(this.btnSave);
            this.panelGrid.Controls.Add(this.dgvCalendar);
            this.panelGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelGrid.Location = new System.Drawing.Point(0, 290);
            this.panelGrid.Name = "panelGrid";
            this.panelGrid.Padding = new System.Windows.Forms.Padding(20, 10, 20, 50);
            this.panelGrid.Size = new System.Drawing.Size(1400, 510);
            this.panelGrid.TabIndex = 3;
            //
            // lblRecordCount
            //
            this.lblRecordCount.AutoSize = true;
            this.lblRecordCount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblRecordCount.Location = new System.Drawing.Point(23, 15);
            this.lblRecordCount.Name = "lblRecordCount";
            this.lblRecordCount.Size = new System.Drawing.Size(96, 15);
            this.lblRecordCount.TabIndex = 3;
            this.lblRecordCount.Text = "Total Programs: 0";
            this.lblRecordCount.Visible = false;
            //
            // btnBack
            //
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBack.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.btnBack.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBack.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnBack.ForeColor = System.Drawing.Color.White;
            this.btnBack.Location = new System.Drawing.Point(1300, 10);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(80, 30);
            this.btnBack.TabIndex = 2;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = false;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            //
            // btnSave
            //
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnSave.ForeColor = System.Drawing.Color.White;
            this.btnSave.Location = new System.Drawing.Point(1210, 10);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(80, 30);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Visible = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            //
            // btnCalculate
            //
            this.btnCalculate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCalculate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(152)))), ((int)(((byte)(0)))));
            this.btnCalculate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCalculate.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnCalculate.ForeColor = System.Drawing.Color.White;
            this.btnCalculate.Location = new System.Drawing.Point(1110, 10);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(90, 30);
            this.btnCalculate.TabIndex = 4;
            this.btnCalculate.Text = "Calculate";
            this.btnCalculate.UseVisualStyleBackColor = false;
            this.btnCalculate.Visible = false;
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);
            //
            // dgvCalendar
            //
            this.dgvCalendar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvCalendar.BackgroundColor = System.Drawing.Color.White;
            this.dgvCalendar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCalendar.Location = new System.Drawing.Point(20, 45);
            this.dgvCalendar.Name = "dgvCalendar";
            this.dgvCalendar.RowHeadersWidth = 30;
            this.dgvCalendar.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.dgvCalendar.Size = new System.Drawing.Size(1360, 400);
            this.dgvCalendar.TabIndex = 0;
            this.dgvCalendar.Visible = false;
            //
            // BookingManagerForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1400, 800);
            this.Controls.Add(this.panelGrid);
            this.Controls.Add(this.panelActions);
            this.Controls.Add(this.panelGridMode);
            this.Controls.Add(this.panelInfo);
            this.Controls.Add(this.panelHeader);
            this.MinimumSize = new System.Drawing.Size(1200, 600);
            this.Name = "BookingManagerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FlexGrid - Booking Manager";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panelHeader.ResumeLayout(false);
            this.panelInfo.ResumeLayout(false);
            this.panelInfo.PerformLayout();
            this.panelGridMode.ResumeLayout(false);
            this.panelGridMode.PerformLayout();
            this.panelActions.ResumeLayout(false);
            this.panelGrid.ResumeLayout(false);
            this.panelGrid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalendar)).EndInit();
            this.ResumeLayout(false);
        }
    
    }
}
