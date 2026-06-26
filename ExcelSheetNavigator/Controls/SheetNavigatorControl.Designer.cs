namespace ExcelSheetNavigator.Controls
{
    partial class SheetNavigatorControl
    {
        /// <summary>Required designer variable.</summary>
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.TableLayoutPanel tableLayoutMain;
        private System.Windows.Forms.TableLayoutPanel tableLayoutHeader;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.FlowLayoutPanel pnlDock;
        private System.Windows.Forms.Label lblDock;
        private System.Windows.Forms.ComboBox cmbDock;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnDensity;
        private System.Windows.Forms.FlowLayoutPanel pnlHeaderButtons;
        private System.Windows.Forms.Panel pnlSearch;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TableLayoutPanel tableLayoutSummary;
        private System.Windows.Forms.Label lblSummary;
        private System.Windows.Forms.Panel pnlListHost;
        private System.Windows.Forms.ListBox lstSheets;
        private System.Windows.Forms.Label lblEmpty;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ToolTip toolTip;

        /// <summary>Clean up any resources being used.</summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
                if (_searchTimer != null) { _searchTimer.Dispose(); _searchTimer = null; }
                if (_boldFont != null) { _boldFont.Dispose(); _boldFont = null; }
                if (_italicFont != null) { _italicFont.Dispose(); _italicFont = null; }
                if (_ownedListFont != null) { _ownedListFont.Dispose(); _ownedListFont = null; }
                if (_brushActiveBg != null) { _brushActiveBg.Dispose(); _brushActiveBg = null; }
                if (_brushHoverBg != null) { _brushHoverBg.Dispose(); _brushHoverBg = null; }
                if (_brushRowBg != null) { _brushRowBg.Dispose(); _brushRowBg = null; }
                if (_brushAccent != null) { _brushAccent.Dispose(); _brushAccent = null; }
                if (_brushStar != null) { _brushStar.Dispose(); _brushStar = null; }
                if (_penIconMuted != null) { _penIconMuted.Dispose(); _penIconMuted = null; }
                if (_penIconActive != null) { _penIconActive.Dispose(); _penIconActive = null; }
                if (_penStarOutline != null) { _penStarOutline.Dispose(); _penStarOutline = null; }
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tableLayoutMain = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutHeader = new System.Windows.Forms.TableLayoutPanel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.pnlDock = new System.Windows.Forms.FlowLayoutPanel();
            this.lblDock = new System.Windows.Forms.Label();
            this.cmbDock = new System.Windows.Forms.ComboBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnDensity = new System.Windows.Forms.Button();
            this.pnlHeaderButtons = new System.Windows.Forms.FlowLayoutPanel();
            this.pnlSearch = new System.Windows.Forms.Panel();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.btnClear = new System.Windows.Forms.Button();
            this.tableLayoutSummary = new System.Windows.Forms.TableLayoutPanel();
            this.lblSummary = new System.Windows.Forms.Label();
            this.pnlListHost = new System.Windows.Forms.Panel();
            this.lstSheets = new System.Windows.Forms.ListBox();
            this.lblEmpty = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.tableLayoutMain.SuspendLayout();
            this.tableLayoutHeader.SuspendLayout();
            this.pnlHeaderButtons.SuspendLayout();
            this.pnlDock.SuspendLayout();
            this.pnlSearch.SuspendLayout();
            this.tableLayoutSummary.SuspendLayout();
            this.pnlListHost.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutMain
            // 
            this.tableLayoutMain.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(229)))), ((int)(((byte)(231)))), ((int)(((byte)(235)))));
            this.tableLayoutMain.ColumnCount = 1;
            this.tableLayoutMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutMain.Controls.Add(this.tableLayoutHeader, 0, 0);
            this.tableLayoutMain.Controls.Add(this.pnlSearch, 0, 1);
            this.tableLayoutMain.Controls.Add(this.tableLayoutSummary, 0, 2);
            this.tableLayoutMain.Controls.Add(this.pnlListHost, 0, 3);
            this.tableLayoutMain.Controls.Add(this.lblStatus, 0, 4);
            this.tableLayoutMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutMain.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutMain.Name = "tableLayoutMain";
            this.tableLayoutMain.Padding = new System.Windows.Forms.Padding(0);
            this.tableLayoutMain.RowCount = 5;
            this.tableLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 52F));
            this.tableLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22F));
            this.tableLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22F));
            this.tableLayoutMain.Size = new System.Drawing.Size(260, 500);
            this.tableLayoutMain.TabIndex = 0;
            // 
            // tableLayoutHeader
            // 
            this.tableLayoutHeader.BackColor = System.Drawing.Color.White;
            this.tableLayoutHeader.ColumnCount = 2;
            this.tableLayoutHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.tableLayoutHeader.Controls.Add(this.lblTitle, 0, 0);
            this.tableLayoutHeader.Controls.Add(this.pnlHeaderButtons, 1, 0);
            this.tableLayoutHeader.Controls.Add(this.pnlDock, 0, 1);
            this.tableLayoutHeader.SetColumnSpan(this.pnlDock, 2);
            this.tableLayoutHeader.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutHeader.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutHeader.Margin = new System.Windows.Forms.Padding(0, 0, 0, 1);
            this.tableLayoutHeader.Name = "tableLayoutHeader";
            this.tableLayoutHeader.Padding = new System.Windows.Forms.Padding(8, 4, 6, 2);
            this.tableLayoutHeader.RowCount = 2;
            this.tableLayoutHeader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23F));
            this.tableLayoutHeader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23F));
            this.tableLayoutHeader.Size = new System.Drawing.Size(260, 52);
            this.tableLayoutHeader.TabIndex = 0;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(23)))), ((int)(((byte)(42)))));
            this.lblTitle.Location = new System.Drawing.Point(8, 2);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(80, 22);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Sheet Navigator";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbDock
            // 
            this.cmbDock.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDock.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmbDock.FormattingEnabled = true;
            this.cmbDock.Items.AddRange(new object[] {
            "Left",
            "Right",
            "Floating"});
            this.cmbDock.Location = new System.Drawing.Point(44, 1);
            this.cmbDock.Margin = new System.Windows.Forms.Padding(0, 1, 0, 0);
            this.cmbDock.Name = "cmbDock";
            this.cmbDock.Size = new System.Drawing.Size(96, 21);
            this.cmbDock.TabIndex = 1;
            this.toolTip.SetToolTip(this.cmbDock, "Dock the panel left, right, or float it");
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right;
            this.btnRefresh.BackColor = System.Drawing.Color.White;
            this.btnRefresh.FlatAppearance.BorderSize = 0;
            this.btnRefresh.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(244)))), ((int)(((byte)(246)))));
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRefresh.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRefresh.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(99)))), ((int)(((byte)(235)))));
            this.btnRefresh.Location = new System.Drawing.Point(224, 0);
            this.btnRefresh.Margin = new System.Windows.Forms.Padding(0);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(24, 22);
            this.btnRefresh.TabIndex = 2;
            this.btnRefresh.Text = "\u21BB";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.toolTip.SetToolTip(this.btnRefresh, "Reload the worksheet list");
            // 
            // btnDensity
            // 
            this.btnDensity.BackColor = System.Drawing.Color.White;
            this.btnDensity.FlatAppearance.BorderSize = 0;
            this.btnDensity.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(243)))), ((int)(((byte)(244)))), ((int)(((byte)(246)))));
            this.btnDensity.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDensity.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDensity.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(116)))), ((int)(((byte)(139)))));
            this.btnDensity.Location = new System.Drawing.Point(0, 0);
            this.btnDensity.Margin = new System.Windows.Forms.Padding(0, 0, 2, 0);
            this.btnDensity.Name = "btnDensity";
            this.btnDensity.Size = new System.Drawing.Size(24, 22);
            this.btnDensity.TabIndex = 3;
            this.btnDensity.Text = "\u2630";
            this.btnDensity.UseVisualStyleBackColor = false;
            this.toolTip.SetToolTip(this.btnDensity, "Compact rows");
            // 
            // pnlHeaderButtons
            // 
            this.pnlHeaderButtons.AutoSize = true;
            this.pnlHeaderButtons.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.pnlHeaderButtons.Controls.Add(this.btnDensity);
            this.pnlHeaderButtons.Controls.Add(this.btnRefresh);
            this.pnlHeaderButtons.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlHeaderButtons.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.pnlHeaderButtons.Location = new System.Drawing.Point(202, 0);
            this.pnlHeaderButtons.Margin = new System.Windows.Forms.Padding(0);
            this.pnlHeaderButtons.Name = "pnlHeaderButtons";
            this.pnlHeaderButtons.Size = new System.Drawing.Size(50, 23);
            this.pnlHeaderButtons.TabIndex = 2;
            this.pnlHeaderButtons.WrapContents = false;
            // 
            // pnlDock
            // 
            this.pnlDock.Controls.Add(this.lblDock);
            this.pnlDock.Controls.Add(this.cmbDock);
            this.pnlDock.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDock.Location = new System.Drawing.Point(10, 30);
            this.pnlDock.Margin = new System.Windows.Forms.Padding(0);
            this.pnlDock.Name = "pnlDock";
            this.pnlDock.Padding = new System.Windows.Forms.Padding(0);
            this.pnlDock.Size = new System.Drawing.Size(244, 27);
            this.pnlDock.TabIndex = 1;
            this.pnlDock.WrapContents = false;
            // 
            // lblDock
            // 
            this.lblDock.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblDock.AutoSize = true;
            this.lblDock.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDock.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(114)))), ((int)(((byte)(128)))));
            this.lblDock.Location = new System.Drawing.Point(0, 5);
            this.lblDock.Margin = new System.Windows.Forms.Padding(0, 5, 5, 0);
            this.lblDock.Name = "lblDock";
            this.lblDock.Size = new System.Drawing.Size(38, 14);
            this.lblDock.TabIndex = 0;
            this.lblDock.Text = "Dock:";
            // 
            // pnlSearch
            // 
            this.pnlSearch.BackColor = System.Drawing.Color.White;
            this.pnlSearch.Controls.Add(this.txtSearch);
            this.pnlSearch.Controls.Add(this.btnClear);
            this.pnlSearch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlSearch.Location = new System.Drawing.Point(0, 62);
            this.pnlSearch.Margin = new System.Windows.Forms.Padding(0, 0, 0, 1);
            this.pnlSearch.Name = "pnlSearch";
            this.pnlSearch.Padding = new System.Windows.Forms.Padding(30, 8, 10, 8);
            this.pnlSearch.Size = new System.Drawing.Size(260, 40);
            this.pnlSearch.TabIndex = 1;
            // 
            // txtSearch
            // 
            this.txtSearch.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtSearch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtSearch.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSearch.Location = new System.Drawing.Point(28, 5);
            this.txtSearch.Margin = new System.Windows.Forms.Padding(0);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(219, 23);
            this.txtSearch.TabIndex = 0;
            // 
            // btnClear
            // 
            this.btnClear.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnClear.FlatAppearance.BorderSize = 0;
            this.btnClear.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClear.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClear.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(114)))), ((int)(((byte)(128)))));
            this.btnClear.Location = new System.Drawing.Point(232, 4);
            this.btnClear.Margin = new System.Windows.Forms.Padding(0);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(20, 21);
            this.btnClear.TabIndex = 1;
            this.btnClear.TabStop = false;
            this.btnClear.Text = "\u00D7";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Visible = false;
            this.toolTip.SetToolTip(this.btnClear, "Clear search");
            // 
            // tableLayoutSummary
            // 
            this.tableLayoutSummary.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(249)))), ((int)(((byte)(250)))), ((int)(((byte)(251)))));
            this.tableLayoutSummary.ColumnCount = 1;
            this.tableLayoutSummary.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutSummary.Controls.Add(this.lblSummary, 0, 0);
            this.tableLayoutSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutSummary.Location = new System.Drawing.Point(0, 72);
            this.tableLayoutSummary.Margin = new System.Windows.Forms.Padding(0, 0, 0, 1);
            this.tableLayoutSummary.Name = "tableLayoutSummary";
            this.tableLayoutSummary.Padding = new System.Windows.Forms.Padding(8, 1, 6, 1);
            this.tableLayoutSummary.RowCount = 1;
            this.tableLayoutSummary.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutSummary.Size = new System.Drawing.Size(260, 19);
            this.tableLayoutSummary.TabIndex = 2;
            // 
            // lblSummary
            // 
            this.lblSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblSummary.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSummary.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(114)))), ((int)(((byte)(128)))));
            this.lblSummary.Location = new System.Drawing.Point(8, 1);
            this.lblSummary.Margin = new System.Windows.Forms.Padding(0);
            this.lblSummary.Name = "lblSummary";
            this.lblSummary.Size = new System.Drawing.Size(186, 17);
            this.lblSummary.TabIndex = 0;
            this.lblSummary.Text = "";
            this.lblSummary.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlListHost
            // 
            this.pnlListHost.BackColor = System.Drawing.Color.White;
            this.pnlListHost.Controls.Add(this.lstSheets);
            this.pnlListHost.Controls.Add(this.lblEmpty);
            this.pnlListHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlListHost.Location = new System.Drawing.Point(0, 98);
            this.pnlListHost.Margin = new System.Windows.Forms.Padding(0, 0, 0, 1);
            this.pnlListHost.Name = "pnlListHost";
            this.pnlListHost.Padding = new System.Windows.Forms.Padding(0);
            this.pnlListHost.Size = new System.Drawing.Size(260, 375);
            this.pnlListHost.TabIndex = 3;
            // 
            // lstSheets
            // 
            this.lstSheets.BackColor = System.Drawing.Color.White;
            this.lstSheets.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lstSheets.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstSheets.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.lstSheets.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstSheets.IntegralHeight = false;
            this.lstSheets.ItemHeight = 30;
            this.lstSheets.Location = new System.Drawing.Point(0, 0);
            this.lstSheets.Margin = new System.Windows.Forms.Padding(0);
            this.lstSheets.Name = "lstSheets";
            this.lstSheets.Size = new System.Drawing.Size(260, 375);
            this.lstSheets.TabIndex = 0;
            // 
            // lblEmpty
            // 
            this.lblEmpty.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblEmpty.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEmpty.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(114)))), ((int)(((byte)(128)))));
            this.lblEmpty.Location = new System.Drawing.Point(0, 0);
            this.lblEmpty.Name = "lblEmpty";
            this.lblEmpty.Size = new System.Drawing.Size(260, 375);
            this.lblEmpty.TabIndex = 1;
            this.lblEmpty.Text = "";
            this.lblEmpty.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblEmpty.Visible = false;
            // 
            // lblStatus
            // 
            this.lblStatus.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(249)))), ((int)(((byte)(250)))), ((int)(((byte)(251)))));
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(114)))), ((int)(((byte)(128)))));
            this.lblStatus.Location = new System.Drawing.Point(0, 474);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Padding = new System.Windows.Forms.Padding(8, 0, 6, 0);
            this.lblStatus.Size = new System.Drawing.Size(260, 20);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "Ready";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // SheetNavigatorControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.tableLayoutMain);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MinimumSize = new System.Drawing.Size(180, 200);
            this.Name = "SheetNavigatorControl";
            this.Size = new System.Drawing.Size(260, 500);
            this.tableLayoutMain.ResumeLayout(false);
            this.pnlHeaderButtons.ResumeLayout(false);
            this.tableLayoutHeader.ResumeLayout(false);
            this.pnlDock.ResumeLayout(false);
            this.pnlDock.PerformLayout();
            this.pnlSearch.ResumeLayout(false);
            this.pnlSearch.PerformLayout();
            this.tableLayoutSummary.ResumeLayout(false);
            this.pnlListHost.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion
    }
}
