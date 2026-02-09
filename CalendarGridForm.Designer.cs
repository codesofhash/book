namespace CSharpFlexGrid
{
    partial class CalendarGridForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Panel panelInfo;
        private System.Windows.Forms.Panel panelControls;
        private System.Windows.Forms.Panel panelGrid;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblAgency;
        private System.Windows.Forms.Label lblAdvertiser;
        private System.Windows.Forms.Label lblProduct;
        private System.Windows.Forms.Label lblCampaign;
        private System.Windows.Forms.Label lblTotalSpots;
        private System.Windows.Forms.Label lblRecordCount;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnBack;
        private CSharpFlexGrid.CustomDataGridView dgvCalendar;

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
            this.lblTotalSpots = new System.Windows.Forms.Label();
            this.lblCampaign = new System.Windows.Forms.Label();
            this.lblProduct = new System.Windows.Forms.Label();
            this.lblAdvertiser = new System.Windows.Forms.Label();
            this.lblAgency = new System.Windows.Forms.Label();
            this.panelControls = new System.Windows.Forms.Panel();
            this.lblRecordCount = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.panelGrid = new System.Windows.Forms.Panel();
            this.dgvCalendar = new CSharpFlexGrid.CustomDataGridView();
            this.panelHeader.SuspendLayout();
            this.panelInfo.SuspendLayout();
            this.panelControls.SuspendLayout();
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
            this.lblTitle.Text = "Calendar View - Monthly Schedule";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // panelInfo
            //
            this.panelInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.panelInfo.Controls.Add(this.lblTotalSpots);
            this.panelInfo.Controls.Add(this.lblCampaign);
            this.panelInfo.Controls.Add(this.lblProduct);
            this.panelInfo.Controls.Add(this.lblAdvertiser);
            this.panelInfo.Controls.Add(this.lblAgency);
            this.panelInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelInfo.Location = new System.Drawing.Point(0, 60);
            this.panelInfo.Name = "panelInfo";
            this.panelInfo.Padding = new System.Windows.Forms.Padding(20, 10, 20, 10);
            this.panelInfo.Size = new System.Drawing.Size(1400, 80);
            this.panelInfo.TabIndex = 1;
            //
            // lblTotalSpots
            //
            this.lblTotalSpots.AutoSize = true;
            this.lblTotalSpots.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblTotalSpots.Location = new System.Drawing.Point(900, 15);
            this.lblTotalSpots.Name = "lblTotalSpots";
            this.lblTotalSpots.Size = new System.Drawing.Size(73, 15);
            this.lblTotalSpots.TabIndex = 4;
            this.lblTotalSpots.Text = "Total Spots:";
            //
            // lblCampaign
            //
            this.lblCampaign.AutoSize = true;
            this.lblCampaign.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCampaign.Location = new System.Drawing.Point(450, 40);
            this.lblCampaign.Name = "lblCampaign";
            this.lblCampaign.Size = new System.Drawing.Size(64, 15);
            this.lblCampaign.TabIndex = 3;
            this.lblCampaign.Text = "Campaign:";
            //
            // lblProduct
            //
            this.lblProduct.AutoSize = true;
            this.lblProduct.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblProduct.Location = new System.Drawing.Point(450, 15);
            this.lblProduct.Name = "lblProduct";
            this.lblProduct.Size = new System.Drawing.Size(52, 15);
            this.lblProduct.TabIndex = 2;
            this.lblProduct.Text = "Product:";
            //
            // lblAdvertiser
            //
            this.lblAdvertiser.AutoSize = true;
            this.lblAdvertiser.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblAdvertiser.Location = new System.Drawing.Point(23, 40);
            this.lblAdvertiser.Name = "lblAdvertiser";
            this.lblAdvertiser.Size = new System.Drawing.Size(65, 15);
            this.lblAdvertiser.TabIndex = 1;
            this.lblAdvertiser.Text = "Advertiser:";
            //
            // lblAgency
            //
            this.lblAgency.AutoSize = true;
            this.lblAgency.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblAgency.Location = new System.Drawing.Point(23, 15);
            this.lblAgency.Name = "lblAgency";
            this.lblAgency.Size = new System.Drawing.Size(51, 15);
            this.lblAgency.TabIndex = 0;
            this.lblAgency.Text = "Agency:";
            //
            // panelControls
            //
            this.panelControls.BackColor = System.Drawing.Color.White;
            this.panelControls.Controls.Add(this.lblRecordCount);
            this.panelControls.Controls.Add(this.btnBack);
            this.panelControls.Controls.Add(this.btnSave);
            this.panelControls.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControls.Location = new System.Drawing.Point(0, 140);
            this.panelControls.Name = "panelControls";
            this.panelControls.Padding = new System.Windows.Forms.Padding(20, 10, 20, 10);
            this.panelControls.Size = new System.Drawing.Size(1400, 60);
            this.panelControls.TabIndex = 2;
            //
            // lblRecordCount
            //
            this.lblRecordCount.AutoSize = true;
            this.lblRecordCount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblRecordCount.Location = new System.Drawing.Point(23, 23);
            this.lblRecordCount.Name = "lblRecordCount";
            this.lblRecordCount.Size = new System.Drawing.Size(96, 15);
            this.lblRecordCount.TabIndex = 2;
            this.lblRecordCount.Text = "Total Programs: 0";
            //
            // btnBack
            //
            this.btnBack.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBack.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(100)))), ((int)(((byte)(100)))));
            this.btnBack.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBack.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnBack.ForeColor = System.Drawing.Color.White;
            this.btnBack.Location = new System.Drawing.Point(1300, 15);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(80, 30);
            this.btnBack.TabIndex = 1;
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
            this.btnSave.Location = new System.Drawing.Point(1210, 15);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(80, 30);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            //
            // panelGrid
            //
            this.panelGrid.Controls.Add(this.dgvCalendar);
            this.panelGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelGrid.Location = new System.Drawing.Point(0, 200);
            this.panelGrid.Name = "panelGrid";
            this.panelGrid.Padding = new System.Windows.Forms.Padding(20, 10, 20, 20);
            this.panelGrid.Size = new System.Drawing.Size(1400, 600);
            this.panelGrid.TabIndex = 3;
            //
            // dgvCalendar
            //
            this.dgvCalendar.BackgroundColor = System.Drawing.Color.White;
            this.dgvCalendar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCalendar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvCalendar.Location = new System.Drawing.Point(20, 10);
            this.dgvCalendar.Name = "dgvCalendar";
            this.dgvCalendar.RowHeadersWidth = 30;
            this.dgvCalendar.Size = new System.Drawing.Size(1360, 570);
            this.dgvCalendar.TabIndex = 0;
            this.dgvCalendar.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvCalendar_CellValueChanged);
            //
            // CalendarGridForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1400, 800);
            this.Controls.Add(this.panelGrid);
            this.Controls.Add(this.panelControls);
            this.Controls.Add(this.panelInfo);
            this.Controls.Add(this.panelHeader);
            this.MinimumSize = new System.Drawing.Size(1200, 600);
            this.Name = "CalendarGridForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FlexGrid - Calendar View";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.panelHeader.ResumeLayout(false);
            this.panelInfo.ResumeLayout(false);
            this.panelInfo.PerformLayout();
            this.panelControls.ResumeLayout(false);
            this.panelControls.PerformLayout();
            this.panelGrid.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCalendar)).EndInit();
            this.ResumeLayout(false);
        }
    }
}
