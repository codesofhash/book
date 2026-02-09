namespace CSharpFlexGrid
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "Excel Grid";

            // 
            // toolStrip
            // 
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.copyButton = new System.Windows.Forms.ToolStripButton();
            this.cutButton = new System.Windows.Forms.ToolStripButton();
            this.pasteButton = new System.Windows.Forms.ToolStripButton();
            this.insertRowButton = new System.Windows.Forms.ToolStripButton();
            this.insertColumnButton = new System.Windows.Forms.ToolStripButton();
            this.deleteRowButton = new System.Windows.Forms.ToolStripButton();
            this.deleteColumnButton = new System.Windows.Forms.ToolStripButton();
            this.renameColumnButton = new System.Windows.Forms.ToolStripButton();
            this.processBookingButton = new System.Windows.Forms.ToolStripButton();
            this.toolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip
            // 
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyButton,
            this.cutButton,
            this.pasteButton,
            new System.Windows.Forms.ToolStripSeparator(),
            this.insertRowButton,
            this.insertColumnButton,
            new System.Windows.Forms.ToolStripSeparator(),
            this.deleteRowButton,
            this.deleteColumnButton,
            new System.Windows.Forms.ToolStripSeparator(),
            this.renameColumnButton,
            new System.Windows.Forms.ToolStripSeparator(),
            this.processBookingButton});
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(800, 25);
            this.toolStrip.TabIndex = 1;
            this.toolStrip.Text = "toolStrip1";
            // 
            // copyButton
            // 
            this.copyButton.Text = "Copy";
            this.copyButton.Click += new System.EventHandler(this.copyButton_Click);
            // 
            // cutButton
            // 
            this.cutButton.Text = "Cut";
            this.cutButton.Click += new System.EventHandler(this.cutButton_Click);
            // 
            // pasteButton
            // 
            this.pasteButton.Text = "Paste";
            this.pasteButton.Click += new System.EventHandler(this.pasteButton_Click);
            // 
            // insertRowButton
            // 
            this.insertRowButton.Text = "Insert Row";
            this.insertRowButton.Click += new System.EventHandler(this.insertRowButton_Click);
            // 
            // insertColumnButton
            // 
            this.insertColumnButton.Text = "Insert Column";
            this.insertColumnButton.Click += new System.EventHandler(this.insertColumnButton_Click);
            // 
            // deleteRowButton
            // 
            this.deleteRowButton.Text = "Delete Row";
            this.deleteRowButton.Click += new System.EventHandler(this.deleteRowButton_Click);
            // 
            // deleteColumnButton
            // 
            this.deleteColumnButton.Text = "Delete Column";
            this.deleteColumnButton.Click += new System.EventHandler(this.deleteColumnButton_Click);
            // 
            // renameColumnButton
            // 
            this.renameColumnButton.Text = "Rename Column";
            this.renameColumnButton.Click += new System.EventHandler(this.renameColumnButton_Click);
            // 
            // processBookingButton
            // 
            this.processBookingButton.Text = "Process Booking";
            this.processBookingButton.Click += new System.EventHandler(this.processBookingButton_Click);
            // 
            // customDataGridView
            // 
            this.customDataGridView = new CSharpFlexGrid.CustomDataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.customDataGridView)).BeginInit();
            this.customDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.customDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.customDataGridView.Location = new System.Drawing.Point(0, 25);
            this.customDataGridView.Name = "customDataGridView";
            this.customDataGridView.Size = new System.Drawing.Size(800, 425);
            this.customDataGridView.TabIndex = 0;
            ((System.ComponentModel.ISupportInitialize)(this.customDataGridView)).EndInit();
            // 
            // MainForm
            // 
            this.Controls.Add(this.customDataGridView);
            this.Controls.Add(this.toolStrip);
            this.Name = "MainForm";
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private CustomDataGridView customDataGridView;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.ToolStripButton copyButton;
        private System.Windows.Forms.ToolStripButton cutButton;
        private System.Windows.Forms.ToolStripButton pasteButton;
        private System.Windows.Forms.ToolStripButton insertRowButton;
        private System.Windows.Forms.ToolStripButton insertColumnButton;
        private System.Windows.Forms.ToolStripButton deleteRowButton;
        private System.Windows.Forms.ToolStripButton deleteColumnButton;
        private System.Windows.Forms.ToolStripButton renameColumnButton;
        private System.Windows.Forms.ToolStripButton processBookingButton;
    }
}