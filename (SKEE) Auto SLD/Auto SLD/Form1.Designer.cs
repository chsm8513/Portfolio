namespace Auto_SLD
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TableLayoutPanel rootLayout;
        private System.Windows.Forms.TableLayoutPanel projectLayout;
        private System.Windows.Forms.Label lblProject;
        private System.Windows.Forms.TextBox txtProject;
        private DevExpress.XtraEditors.SimpleButton btnFindProject;
        private System.Windows.Forms.GroupBox worksheetGroup;
        private System.Windows.Forms.DataGridView dataGridView;
        private Auto_SLD.Controls.BusyProgressPanel worksheetProgressPanel;
        private DevExpress.XtraEditors.SimpleButton btnAutoDrawing;
        private DevExpress.XtraEditors.SimpleButton btnExit;
        private System.Windows.Forms.Label lblStatus;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.rootLayout = new System.Windows.Forms.TableLayoutPanel();
            this.projectLayout = new System.Windows.Forms.TableLayoutPanel();
            this.lblProject = new System.Windows.Forms.Label();
            this.txtProject = new System.Windows.Forms.TextBox();
            this.btnFindProject = new DevExpress.XtraEditors.SimpleButton();
            this.worksheetGroup = new System.Windows.Forms.GroupBox();
            this.worksheetProgressPanel = new Auto_SLD.Controls.BusyProgressPanel();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.tablePanel2 = new DevExpress.Utils.Layout.TablePanel();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnAutoDrawing = new DevExpress.XtraEditors.SimpleButton();
            this.btnExit = new DevExpress.XtraEditors.SimpleButton();
            this.rootLayout.SuspendLayout();
            this.projectLayout.SuspendLayout();
            this.worksheetGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.worksheetProgressPanel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tablePanel2)).BeginInit();
            this.tablePanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // rootLayout
            // 
            this.rootLayout.ColumnCount = 1;
            this.rootLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rootLayout.Controls.Add(this.projectLayout, 0, 0);
            this.rootLayout.Controls.Add(this.worksheetGroup, 0, 1);
            this.rootLayout.Controls.Add(this.tablePanel2, 0, 2);
            this.rootLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rootLayout.Location = new System.Drawing.Point(0, 0);
            this.rootLayout.Name = "rootLayout";
            this.rootLayout.Padding = new System.Windows.Forms.Padding(12);
            this.rootLayout.RowCount = 4;
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 70F));
            this.rootLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.rootLayout.Size = new System.Drawing.Size(958, 543);
            this.rootLayout.TabIndex = 0;
            // 
            // projectLayout
            // 
            this.projectLayout.ColumnCount = 3;
            this.projectLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 110F));
            this.projectLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.projectLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 160F));
            this.projectLayout.Controls.Add(this.lblProject, 0, 0);
            this.projectLayout.Controls.Add(this.txtProject, 1, 0);
            this.projectLayout.Controls.Add(this.btnFindProject, 2, 0);
            this.projectLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.projectLayout.Location = new System.Drawing.Point(15, 15);
            this.projectLayout.Name = "projectLayout";
            this.projectLayout.RowCount = 1;
            this.projectLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.projectLayout.Size = new System.Drawing.Size(928, 42);
            this.projectLayout.TabIndex = 0;
            // 
            // lblProject
            // 
            this.lblProject.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblProject.Location = new System.Drawing.Point(3, 0);
            this.lblProject.Name = "lblProject";
            this.lblProject.Size = new System.Drawing.Size(104, 42);
            this.lblProject.TabIndex = 0;
            this.lblProject.Text = "EB Project";
            this.lblProject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtProject
            // 
            this.txtProject.BackColor = System.Drawing.Color.WhiteSmoke;
            this.txtProject.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtProject.Location = new System.Drawing.Point(113, 10);
            this.txtProject.Margin = new System.Windows.Forms.Padding(3, 10, 8, 8);
            this.txtProject.Name = "txtProject";
            this.txtProject.ReadOnly = true;
            this.txtProject.Size = new System.Drawing.Size(647, 22);
            this.txtProject.TabIndex = 1;
            // 
            // btnFindProject
            // 
            this.btnFindProject.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnFindProject.Location = new System.Drawing.Point(771, 7);
            this.btnFindProject.Margin = new System.Windows.Forms.Padding(3, 7, 3, 7);
            this.btnFindProject.Name = "btnFindProject";
            this.btnFindProject.Size = new System.Drawing.Size(154, 28);
            this.btnFindProject.TabIndex = 2;
            this.btnFindProject.Text = "현재 프로젝트 찾기";
            this.btnFindProject.Click += new System.EventHandler(this.btnFindProject_Click);
            // 
            // worksheetGroup
            // 
            this.worksheetGroup.Controls.Add(this.worksheetProgressPanel);
            this.worksheetGroup.Controls.Add(this.dataGridView);
            this.worksheetGroup.Dock = System.Windows.Forms.DockStyle.Fill;
            this.worksheetGroup.Location = new System.Drawing.Point(12, 66);
            this.worksheetGroup.Margin = new System.Windows.Forms.Padding(0, 6, 0, 6);
            this.worksheetGroup.Name = "worksheetGroup";
            this.worksheetGroup.Padding = new System.Windows.Forms.Padding(8, 6, 8, 8);
            this.worksheetGroup.Size = new System.Drawing.Size(934, 381);
            this.worksheetGroup.TabIndex = 1;
            this.worksheetGroup.TabStop = false;
            this.worksheetGroup.Text = "EB Worksheet Objects";
            // 
            // worksheetProgressPanel
            // 
            this.worksheetProgressPanel.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.worksheetProgressPanel.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.worksheetProgressPanel.Caption = "Loading...";
            this.worksheetProgressPanel.ContentAlignment = System.Drawing.ContentAlignment.MiddleCenter;
            this.worksheetProgressPanel.Description = "...";
            this.worksheetProgressPanel.Location = new System.Drawing.Point(327, 157);
            this.worksheetProgressPanel.MinimumSize = new System.Drawing.Size(280, 66);
            this.worksheetProgressPanel.Name = "worksheetProgressPanel";
            this.worksheetProgressPanel.Padding = new System.Windows.Forms.Padding(12, 0, 12, 0);
            this.worksheetProgressPanel.Size = new System.Drawing.Size(280, 66);
            this.worksheetProgressPanel.TabIndex = 1;
            this.worksheetProgressPanel.Visible = false;
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.AllowUserToOrderColumns = true;
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(56)))), ((int)(((byte)(75)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 9F);
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(37)))), ((int)(((byte)(56)))), ((int)(((byte)(75)))));
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView.ColumnHeadersHeight = 34;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.EnableHeadersVisualStyles = false;
            this.dataGridView.Location = new System.Drawing.Point(8, 21);
            this.dataGridView.MultiSelect = false;
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.ReadOnly = true;
            this.dataGridView.RowHeadersVisible = false;
            this.dataGridView.RowHeadersWidth = 58;
            this.dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView.Size = new System.Drawing.Size(918, 352);
            this.dataGridView.TabIndex = 0;
            // 
            // tablePanel2
            // 
            this.tablePanel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tablePanel2.Columns.AddRange(new DevExpress.Utils.Layout.TablePanelColumn[] {
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 21.34F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 11.66F),
            new DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 77F)});
            this.tablePanel2.Controls.Add(this.lblStatus);
            this.tablePanel2.Controls.Add(this.btnAutoDrawing);
            this.tablePanel2.Controls.Add(this.btnExit);
            this.tablePanel2.Location = new System.Drawing.Point(12, 453);
            this.tablePanel2.Margin = new System.Windows.Forms.Padding(0);
            this.tablePanel2.Name = "tablePanel2";
            this.tablePanel2.Rows.AddRange(new DevExpress.Utils.Layout.TablePanelRow[] {
            new DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 26F)});
            this.tablePanel2.Size = new System.Drawing.Size(934, 70);
            this.tablePanel2.TabIndex = 0;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoEllipsis = true;
            this.tablePanel2.SetColumn(this.lblStatus, 2);
            this.lblStatus.ForeColor = System.Drawing.Color.DimGray;
            this.lblStatus.Location = new System.Drawing.Point(283, 16);
            this.lblStatus.Name = "lblStatus";
            this.tablePanel2.SetRow(this.lblStatus, 0);
            this.lblStatus.Size = new System.Drawing.Size(648, 38);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "EB 프로젝트를 선택하세요. Worksheet 연동은 다음 단계에서 구성합니다.";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnAutoDrawing
            // 
            this.tablePanel2.SetColumn(this.btnAutoDrawing, 0);
            this.btnAutoDrawing.Location = new System.Drawing.Point(6, 19);
            this.btnAutoDrawing.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.btnAutoDrawing.Name = "btnAutoDrawing";
            this.tablePanel2.SetRow(this.btnAutoDrawing, 0);
            this.btnAutoDrawing.Size = new System.Drawing.Size(169, 32);
            this.btnAutoDrawing.TabIndex = 1;
            this.btnAutoDrawing.Text = "Auto Drawing";
            this.btnAutoDrawing.Click += new System.EventHandler(this.btnAutoDrawing_Click);
            // 
            // btnExit
            // 
            this.tablePanel2.SetColumn(this.btnExit, 1);
            this.btnExit.Location = new System.Drawing.Point(187, 19);
            this.btnExit.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.btnExit.Name = "btnExit";
            this.tablePanel2.SetRow(this.btnExit, 0);
            this.btnExit.Size = new System.Drawing.Size(87, 32);
            this.btnExit.TabIndex = 0;
            this.btnExit.Text = "닫기";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(958, 543);
            this.Controls.Add(this.rootLayout);
            this.MinimumSize = new System.Drawing.Size(920, 560);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Auto SLD";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.rootLayout.ResumeLayout(false);
            this.projectLayout.ResumeLayout(false);
            this.projectLayout.PerformLayout();
            this.worksheetGroup.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.worksheetProgressPanel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tablePanel2)).EndInit();
            this.tablePanel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        private DevExpress.Utils.Layout.TablePanel tablePanel2;
    }
}
