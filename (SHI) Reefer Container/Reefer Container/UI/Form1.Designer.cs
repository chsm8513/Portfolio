namespace Reefer_Container
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btn_excel_import = new System.Windows.Forms.Button();
            this.btn_create_data = new System.Windows.Forms.Button();
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.btn_setting = new System.Windows.Forms.Button();
            this.panel_setting = new System.Windows.Forms.Panel();
            this.tb_revision = new System.Windows.Forms.TextBox();
            this.sortPriorityPanel = new Reefer_Container.SortPriorityPanel();
            this.btn_default_setting = new System.Windows.Forms.Button();
            this.btn_save_setting = new System.Windows.Forms.Button();
            this.lbl_tr_count = new System.Windows.Forms.Label();
            this.label_tr_count_title = new System.Windows.Forms.Label();
            this.lbl_tr_result = new System.Windows.Forms.Label();
            this.lbl_srd_result = new System.Windows.Forms.Label();
            this.lbl_rdb_result = new System.Windows.Forms.Label();
            this.label_gen_unit = new System.Windows.Forms.Label();
            this.label_tr_unit = new System.Windows.Forms.Label();
            this.label_srd_unit = new System.Windows.Forms.Label();
            this.label_rdb_unit = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label_reefer_unit = new System.Windows.Forms.Label();
            this.label_max_gang_unit = new System.Windows.Forms.Label();
            this.label_max_rdb_unit = new System.Windows.Forms.Label();
            this.label_max_af_unit = new System.Windows.Forms.Label();
            this.label_generator_basis = new System.Windows.Forms.Label();
            this.label_srd_basis = new System.Windows.Forms.Label();
            this.num_generator_basis = new System.Windows.Forms.NumericUpDown();
            this.num_srd_basis = new System.Windows.Forms.NumericUpDown();
            this.num_transformer_basis = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.num_distboard_basis = new System.Windows.Forms.NumericUpDown();
            this.num_TR = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.num_SRD = new System.Windows.Forms.NumericUpDown();
            this.cb_deck_hold_type = new System.Windows.Forms.ComboBox();
            this.num_RDB = new System.Windows.Forms.NumericUpDown();
            this.chk_deck_first = new System.Windows.Forms.CheckBox();
            this.num_Reefer_basis = new System.Windows.Forms.NumericUpDown();
            this.num_Max_gang_cnt = new System.Windows.Forms.NumericUpDown();
            this.num_Max_rdb_reefer_cnt = new System.Windows.Forms.NumericUpDown();
            this.num_Max_AF = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_update_data = new System.Windows.Forms.Button();
            this.btn_change_color = new System.Windows.Forms.Button();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.spreadsheetControl1 = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.button1 = new System.Windows.Forms.Button();
            this.btn_rebuild = new System.Windows.Forms.Button();
            this.btn_callworksheet = new System.Windows.Forms.Button();
            this.lb_projectname = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            this.panel_setting.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.num_generator_basis)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_srd_basis)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_transformer_basis)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_distboard_basis)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_TR)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_SRD)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_RDB)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Reefer_basis)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Max_gang_cnt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Max_rdb_reefer_cnt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Max_AF)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_excel_import
            // 
            this.btn_excel_import.Location = new System.Drawing.Point(93, 24);
            this.btn_excel_import.Name = "btn_excel_import";
            this.btn_excel_import.Size = new System.Drawing.Size(88, 23);
            this.btn_excel_import.TabIndex = 0;
            this.btn_excel_import.Text = "Excel Import";
            this.btn_excel_import.UseVisualStyleBackColor = true;
            this.btn_excel_import.Click += new System.EventHandler(this.btn_excel_import_Click);
            // 
            // btn_create_data
            // 
            this.btn_create_data.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_create_data.Location = new System.Drawing.Point(670, 24);
            this.btn_create_data.Name = "btn_create_data";
            this.btn_create_data.Size = new System.Drawing.Size(86, 23);
            this.btn_create_data.TabIndex = 1;
            this.btn_create_data.Text = "Create Data";
            this.btn_create_data.UseVisualStyleBackColor = true;
            this.btn_create_data.Click += new System.EventHandler(this.btn_create_data_Click);
            // 
            // gridControl1
            // 
            this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridControl1.Location = new System.Drawing.Point(12, 53);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.Size = new System.Drawing.Size(789, 516);
            this.gridControl1.TabIndex = 2;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.Name = "gridView1";
            // 
            // btn_setting
            // 
            this.btn_setting.Location = new System.Drawing.Point(12, 24);
            this.btn_setting.Name = "btn_setting";
            this.btn_setting.Size = new System.Drawing.Size(75, 23);
            this.btn_setting.TabIndex = 3;
            this.btn_setting.Text = "Setting";
            this.btn_setting.UseVisualStyleBackColor = true;
            this.btn_setting.Click += new System.EventHandler(this.bnt_setting_Click);
            // 
            // panel_setting
            // 
            this.panel_setting.BackColor = System.Drawing.Color.Silver;
            this.panel_setting.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel_setting.Controls.Add(this.tb_revision);
            this.panel_setting.Controls.Add(this.sortPriorityPanel);
            this.panel_setting.Controls.Add(this.btn_default_setting);
            this.panel_setting.Controls.Add(this.btn_save_setting);
            this.panel_setting.Controls.Add(this.lbl_tr_count);
            this.panel_setting.Controls.Add(this.label_tr_count_title);
            this.panel_setting.Controls.Add(this.lbl_tr_result);
            this.panel_setting.Controls.Add(this.lbl_srd_result);
            this.panel_setting.Controls.Add(this.lbl_rdb_result);
            this.panel_setting.Controls.Add(this.label_gen_unit);
            this.panel_setting.Controls.Add(this.label_tr_unit);
            this.panel_setting.Controls.Add(this.label_srd_unit);
            this.panel_setting.Controls.Add(this.label_rdb_unit);
            this.panel_setting.Controls.Add(this.label10);
            this.panel_setting.Controls.Add(this.label7);
            this.panel_setting.Controls.Add(this.label6);
            this.panel_setting.Controls.Add(this.label_reefer_unit);
            this.panel_setting.Controls.Add(this.label_max_gang_unit);
            this.panel_setting.Controls.Add(this.label_max_rdb_unit);
            this.panel_setting.Controls.Add(this.label_max_af_unit);
            this.panel_setting.Controls.Add(this.label_generator_basis);
            this.panel_setting.Controls.Add(this.label_srd_basis);
            this.panel_setting.Controls.Add(this.num_generator_basis);
            this.panel_setting.Controls.Add(this.num_srd_basis);
            this.panel_setting.Controls.Add(this.num_transformer_basis);
            this.panel_setting.Controls.Add(this.label12);
            this.panel_setting.Controls.Add(this.num_distboard_basis);
            this.panel_setting.Controls.Add(this.num_TR);
            this.panel_setting.Controls.Add(this.label9);
            this.panel_setting.Controls.Add(this.num_SRD);
            this.panel_setting.Controls.Add(this.cb_deck_hold_type);
            this.panel_setting.Controls.Add(this.num_RDB);
            this.panel_setting.Controls.Add(this.chk_deck_first);
            this.panel_setting.Controls.Add(this.num_Reefer_basis);
            this.panel_setting.Controls.Add(this.num_Max_gang_cnt);
            this.panel_setting.Controls.Add(this.num_Max_rdb_reefer_cnt);
            this.panel_setting.Controls.Add(this.num_Max_AF);
            this.panel_setting.Controls.Add(this.label11);
            this.panel_setting.Controls.Add(this.label8);
            this.panel_setting.Controls.Add(this.label4);
            this.panel_setting.Controls.Add(this.label3);
            this.panel_setting.Controls.Add(this.label2);
            this.panel_setting.Controls.Add(this.label1);
            this.panel_setting.Location = new System.Drawing.Point(12, 53);
            this.panel_setting.Name = "panel_setting";
            this.panel_setting.Size = new System.Drawing.Size(519, 499);
            this.panel_setting.TabIndex = 4;
            this.panel_setting.Visible = false;
            // 
            // tb_revision
            // 
            this.tb_revision.Location = new System.Drawing.Point(170, 291);
            this.tb_revision.Name = "tb_revision";
            this.tb_revision.Size = new System.Drawing.Size(121, 22);
            this.tb_revision.TabIndex = 40;
            this.tb_revision.Text = "Rev0";
            // 
            // sortPriorityPanel
            // 
            this.sortPriorityPanel.AutoSize = true;
            this.sortPriorityPanel.Location = new System.Drawing.Point(29, 326);
            this.sortPriorityPanel.Name = "sortPriorityPanel";
            this.sortPriorityPanel.Size = new System.Drawing.Size(363, 125);
            this.sortPriorityPanel.TabIndex = 6;
            // 
            // btn_default_setting
            // 
            this.btn_default_setting.Location = new System.Drawing.Point(408, 467);
            this.btn_default_setting.Name = "btn_default_setting";
            this.btn_default_setting.Size = new System.Drawing.Size(80, 23);
            this.btn_default_setting.TabIndex = 38;
            this.btn_default_setting.Text = "Default";
            this.btn_default_setting.UseVisualStyleBackColor = true;
            this.btn_default_setting.Click += new System.EventHandler(this.btn_default_setting_Click);
            // 
            // btn_save_setting
            // 
            this.btn_save_setting.Location = new System.Drawing.Point(323, 467);
            this.btn_save_setting.Name = "btn_save_setting";
            this.btn_save_setting.Size = new System.Drawing.Size(80, 23);
            this.btn_save_setting.TabIndex = 37;
            this.btn_save_setting.Text = "Save";
            this.btn_save_setting.UseVisualStyleBackColor = true;
            this.btn_save_setting.Click += new System.EventHandler(this.btn_save_setting_Click);
            // 
            // lbl_tr_count
            // 
            this.lbl_tr_count.AutoSize = true;
            this.lbl_tr_count.BackColor = System.Drawing.Color.Transparent;
            this.lbl_tr_count.Location = new System.Drawing.Point(171, 10);
            this.lbl_tr_count.Name = "lbl_tr_count";
            this.lbl_tr_count.Size = new System.Drawing.Size(33, 14);
            this.lbl_tr_count.TabIndex = 25;
            this.lbl_tr_count.Text = "0 EA";
            // 
            // label_tr_count_title
            // 
            this.label_tr_count_title.AutoSize = true;
            this.label_tr_count_title.Location = new System.Drawing.Point(27, 10);
            this.label_tr_count_title.Name = "label_tr_count_title";
            this.label_tr_count_title.Size = new System.Drawing.Size(89, 14);
            this.label_tr_count_title.TabIndex = 36;
            this.label_tr_count_title.Text = "Reefer TR Q\'ty";
            // 
            // lbl_tr_result
            // 
            this.lbl_tr_result.AutoSize = true;
            this.lbl_tr_result.BackColor = System.Drawing.Color.Transparent;
            this.lbl_tr_result.Location = new System.Drawing.Point(416, 206);
            this.lbl_tr_result.Name = "lbl_tr_result";
            this.lbl_tr_result.Size = new System.Drawing.Size(40, 14);
            this.lbl_tr_result.TabIndex = 24;
            this.lbl_tr_result.Text = "11.8A";
            // 
            // lbl_srd_result
            // 
            this.lbl_srd_result.AutoSize = true;
            this.lbl_srd_result.BackColor = System.Drawing.Color.Transparent;
            this.lbl_srd_result.Location = new System.Drawing.Point(416, 178);
            this.lbl_srd_result.Name = "lbl_srd_result";
            this.lbl_srd_result.Size = new System.Drawing.Size(40, 14);
            this.lbl_srd_result.TabIndex = 23;
            this.lbl_srd_result.Text = "15.2A";
            // 
            // lbl_rdb_result
            // 
            this.lbl_rdb_result.AutoSize = true;
            this.lbl_rdb_result.BackColor = System.Drawing.Color.Transparent;
            this.lbl_rdb_result.Location = new System.Drawing.Point(416, 150);
            this.lbl_rdb_result.Name = "lbl_rdb_result";
            this.lbl_rdb_result.Size = new System.Drawing.Size(40, 14);
            this.lbl_rdb_result.TabIndex = 22;
            this.lbl_rdb_result.Text = "16.9A";
            // 
            // label_gen_unit
            // 
            this.label_gen_unit.AutoSize = true;
            this.label_gen_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_gen_unit.Location = new System.Drawing.Point(296, 236);
            this.label_gen_unit.Name = "label_gen_unit";
            this.label_gen_unit.Size = new System.Drawing.Size(25, 14);
            this.label_gen_unit.TabIndex = 33;
            this.label_gen_unit.Text = "kW";
            // 
            // label_tr_unit
            // 
            this.label_tr_unit.AutoSize = true;
            this.label_tr_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_tr_unit.Location = new System.Drawing.Point(296, 206);
            this.label_tr_unit.Name = "label_tr_unit";
            this.label_tr_unit.Size = new System.Drawing.Size(19, 14);
            this.label_tr_unit.TabIndex = 32;
            this.label_tr_unit.Text = "%";
            // 
            // label_srd_unit
            // 
            this.label_srd_unit.AutoSize = true;
            this.label_srd_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_srd_unit.Location = new System.Drawing.Point(296, 178);
            this.label_srd_unit.Name = "label_srd_unit";
            this.label_srd_unit.Size = new System.Drawing.Size(19, 14);
            this.label_srd_unit.TabIndex = 31;
            this.label_srd_unit.Text = "%";
            // 
            // label_rdb_unit
            // 
            this.label_rdb_unit.AutoSize = true;
            this.label_rdb_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_rdb_unit.Location = new System.Drawing.Point(296, 150);
            this.label_rdb_unit.Name = "label_rdb_unit";
            this.label_rdb_unit.Size = new System.Drawing.Size(19, 14);
            this.label_rdb_unit.TabIndex = 30;
            this.label_rdb_unit.Text = "%";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label10.Location = new System.Drawing.Point(385, 204);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(25, 14);
            this.label10.TabIndex = 29;
            this.label10.Text = "kW";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label7.Location = new System.Drawing.Point(385, 178);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(25, 14);
            this.label7.TabIndex = 29;
            this.label7.Text = "kW";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Location = new System.Drawing.Point(385, 150);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(25, 14);
            this.label6.TabIndex = 29;
            this.label6.Text = "kW";
            // 
            // label_reefer_unit
            // 
            this.label_reefer_unit.AutoSize = true;
            this.label_reefer_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_reefer_unit.Location = new System.Drawing.Point(296, 124);
            this.label_reefer_unit.Name = "label_reefer_unit";
            this.label_reefer_unit.Size = new System.Drawing.Size(25, 14);
            this.label_reefer_unit.TabIndex = 29;
            this.label_reefer_unit.Text = "kW";
            // 
            // label_max_gang_unit
            // 
            this.label_max_gang_unit.AutoSize = true;
            this.label_max_gang_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_max_gang_unit.Location = new System.Drawing.Point(296, 94);
            this.label_max_gang_unit.Name = "label_max_gang_unit";
            this.label_max_gang_unit.Size = new System.Drawing.Size(22, 14);
            this.label_max_gang_unit.TabIndex = 28;
            this.label_max_gang_unit.Text = "EA";
            // 
            // label_max_rdb_unit
            // 
            this.label_max_rdb_unit.AutoSize = true;
            this.label_max_rdb_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_max_rdb_unit.Location = new System.Drawing.Point(296, 65);
            this.label_max_rdb_unit.Name = "label_max_rdb_unit";
            this.label_max_rdb_unit.Size = new System.Drawing.Size(22, 14);
            this.label_max_rdb_unit.TabIndex = 27;
            this.label_max_rdb_unit.Text = "EA";
            // 
            // label_max_af_unit
            // 
            this.label_max_af_unit.AutoSize = true;
            this.label_max_af_unit.BackColor = System.Drawing.Color.Transparent;
            this.label_max_af_unit.Location = new System.Drawing.Point(296, 38);
            this.label_max_af_unit.Name = "label_max_af_unit";
            this.label_max_af_unit.Size = new System.Drawing.Size(21, 14);
            this.label_max_af_unit.TabIndex = 26;
            this.label_max_af_unit.Text = "AF";
            // 
            // label_generator_basis
            // 
            this.label_generator_basis.AutoSize = true;
            this.label_generator_basis.Location = new System.Drawing.Point(27, 234);
            this.label_generator_basis.Name = "label_generator_basis";
            this.label_generator_basis.Size = new System.Drawing.Size(30, 14);
            this.label_generator_basis.TabIndex = 35;
            this.label_generator_basis.Text = "GEN";
            // 
            // label_srd_basis
            // 
            this.label_srd_basis.AutoSize = true;
            this.label_srd_basis.Location = new System.Drawing.Point(27, 178);
            this.label_srd_basis.Name = "label_srd_basis";
            this.label_srd_basis.Size = new System.Drawing.Size(80, 14);
            this.label_srd_basis.TabIndex = 34;
            this.label_srd_basis.Text = "SRD  ->  RDB";
            // 
            // num_generator_basis
            // 
            this.num_generator_basis.DecimalPlaces = 1;
            this.num_generator_basis.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.num_generator_basis.Location = new System.Drawing.Point(171, 232);
            this.num_generator_basis.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.num_generator_basis.Name = "num_generator_basis";
            this.num_generator_basis.Size = new System.Drawing.Size(120, 22);
            this.num_generator_basis.TabIndex = 21;
            this.num_generator_basis.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.num_generator_basis.ValueChanged += new System.EventHandler(this.num_generator_basis_ValueChanged);
            // 
            // num_srd_basis
            // 
            this.num_srd_basis.Location = new System.Drawing.Point(171, 176);
            this.num_srd_basis.Name = "num_srd_basis";
            this.num_srd_basis.Size = new System.Drawing.Size(119, 22);
            this.num_srd_basis.TabIndex = 20;
            this.num_srd_basis.Value = new decimal(new int[] {
            90,
            0,
            0,
            0});
            this.num_srd_basis.ValueChanged += new System.EventHandler(this.num_srd_basis_ValueChanged);
            // 
            // num_transformer_basis
            // 
            this.num_transformer_basis.Location = new System.Drawing.Point(171, 204);
            this.num_transformer_basis.Name = "num_transformer_basis";
            this.num_transformer_basis.Size = new System.Drawing.Size(119, 22);
            this.num_transformer_basis.TabIndex = 19;
            this.num_transformer_basis.Value = new decimal(new int[] {
            70,
            0,
            0,
            0});
            this.num_transformer_basis.ValueChanged += new System.EventHandler(this.num_transformer_basis_ValueChanged);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(27, 206);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(77, 14);
            this.label12.TabIndex = 18;
            this.label12.Text = "TR   ->  SRD";
            // 
            // num_distboard_basis
            // 
            this.num_distboard_basis.Location = new System.Drawing.Point(171, 148);
            this.num_distboard_basis.Name = "num_distboard_basis";
            this.num_distboard_basis.Size = new System.Drawing.Size(119, 22);
            this.num_distboard_basis.TabIndex = 9;
            this.num_distboard_basis.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.num_distboard_basis.ValueChanged += new System.EventHandler(this.num_distboard_basis_ValueChanged);
            // 
            // num_TR
            // 
            this.num_TR.DecimalPlaces = 1;
            this.num_TR.Enabled = false;
            this.num_TR.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.num_TR.Location = new System.Drawing.Point(323, 202);
            this.num_TR.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.num_TR.Name = "num_TR";
            this.num_TR.ReadOnly = true;
            this.num_TR.Size = new System.Drawing.Size(56, 22);
            this.num_TR.TabIndex = 1;
            this.num_TR.Value = new decimal(new int[] {
            103,
            0,
            0,
            65536});
            this.num_TR.ValueChanged += new System.EventHandler(this.num_Reefer_basis_ValueChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(27, 150);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(86, 14);
            this.label9.TabIndex = 8;
            this.label9.Text = "RDB  -> GANG";
            // 
            // num_SRD
            // 
            this.num_SRD.DecimalPlaces = 1;
            this.num_SRD.Enabled = false;
            this.num_SRD.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.num_SRD.Location = new System.Drawing.Point(323, 176);
            this.num_SRD.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.num_SRD.Name = "num_SRD";
            this.num_SRD.ReadOnly = true;
            this.num_SRD.Size = new System.Drawing.Size(56, 22);
            this.num_SRD.TabIndex = 1;
            this.num_SRD.Value = new decimal(new int[] {
            103,
            0,
            0,
            65536});
            this.num_SRD.ValueChanged += new System.EventHandler(this.num_Reefer_basis_ValueChanged);
            // 
            // cb_deck_hold_type
            // 
            this.cb_deck_hold_type.DisplayMember = "DECK / HOLD";
            this.cb_deck_hold_type.FormattingEnabled = true;
            this.cb_deck_hold_type.Items.AddRange(new object[] {
            "DECK / HOLD",
            "DECK + HOLD"});
            this.cb_deck_hold_type.Location = new System.Drawing.Point(170, 261);
            this.cb_deck_hold_type.Name = "cb_deck_hold_type";
            this.cb_deck_hold_type.Size = new System.Drawing.Size(121, 22);
            this.cb_deck_hold_type.TabIndex = 2;
            this.cb_deck_hold_type.SelectedIndexChanged += new System.EventHandler(this.cb_deck_hold_type_SelectedIndexChanged);
            // 
            // num_RDB
            // 
            this.num_RDB.DecimalPlaces = 1;
            this.num_RDB.Enabled = false;
            this.num_RDB.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.num_RDB.Location = new System.Drawing.Point(323, 148);
            this.num_RDB.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.num_RDB.Name = "num_RDB";
            this.num_RDB.ReadOnly = true;
            this.num_RDB.Size = new System.Drawing.Size(56, 22);
            this.num_RDB.TabIndex = 1;
            this.num_RDB.Value = new decimal(new int[] {
            103,
            0,
            0,
            65536});
            this.num_RDB.ValueChanged += new System.EventHandler(this.num_Reefer_basis_ValueChanged);
            // 
            // chk_deck_first
            // 
            this.chk_deck_first.AutoSize = true;
            this.chk_deck_first.Checked = true;
            this.chk_deck_first.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chk_deck_first.Location = new System.Drawing.Point(300, 264);
            this.chk_deck_first.Name = "chk_deck_first";
            this.chk_deck_first.Size = new System.Drawing.Size(98, 18);
            this.chk_deck_first.TabIndex = 39;
            this.chk_deck_first.Text = "Deck -> Hold";
            this.chk_deck_first.UseVisualStyleBackColor = true;
            this.chk_deck_first.CheckedChanged += new System.EventHandler(this.chk_deck_first_CheckedChanged);
            // 
            // num_Reefer_basis
            // 
            this.num_Reefer_basis.DecimalPlaces = 1;
            this.num_Reefer_basis.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.num_Reefer_basis.Location = new System.Drawing.Point(171, 120);
            this.num_Reefer_basis.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.num_Reefer_basis.Name = "num_Reefer_basis";
            this.num_Reefer_basis.Size = new System.Drawing.Size(120, 22);
            this.num_Reefer_basis.TabIndex = 1;
            this.num_Reefer_basis.Value = new decimal(new int[] {
            103,
            0,
            0,
            65536});
            this.num_Reefer_basis.ValueChanged += new System.EventHandler(this.num_Reefer_basis_ValueChanged);
            // 
            // num_Max_gang_cnt
            // 
            this.num_Max_gang_cnt.Location = new System.Drawing.Point(171, 91);
            this.num_Max_gang_cnt.Name = "num_Max_gang_cnt";
            this.num_Max_gang_cnt.Size = new System.Drawing.Size(120, 22);
            this.num_Max_gang_cnt.TabIndex = 1;
            this.num_Max_gang_cnt.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.num_Max_gang_cnt.ValueChanged += new System.EventHandler(this.num_Max_gangbox_cnt_ValueChanged);
            // 
            // num_Max_rdb_reefer_cnt
            // 
            this.num_Max_rdb_reefer_cnt.Location = new System.Drawing.Point(171, 60);
            this.num_Max_rdb_reefer_cnt.Maximum = new decimal(new int[] {
            999,
            0,
            0,
            0});
            this.num_Max_rdb_reefer_cnt.Name = "num_Max_rdb_reefer_cnt";
            this.num_Max_rdb_reefer_cnt.Size = new System.Drawing.Size(120, 22);
            this.num_Max_rdb_reefer_cnt.TabIndex = 1;
            this.num_Max_rdb_reefer_cnt.Value = new decimal(new int[] {
            80,
            0,
            0,
            0});
            this.num_Max_rdb_reefer_cnt.ValueChanged += new System.EventHandler(this.num_Max_rdb_reefer_cnt_ValueChanged);
            // 
            // num_Max_AF
            // 
            this.num_Max_AF.Location = new System.Drawing.Point(171, 33);
            this.num_Max_AF.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.num_Max_AF.Name = "num_Max_AF";
            this.num_Max_AF.Size = new System.Drawing.Size(120, 22);
            this.num_Max_AF.TabIndex = 1;
            this.num_Max_AF.Value = new decimal(new int[] {
            1250,
            0,
            0,
            0});
            this.num_Max_AF.ValueChanged += new System.EventHandler(this.num_Max_AF_ValueChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(27, 294);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(50, 14);
            this.label11.TabIndex = 0;
            this.label11.Text = "Revision";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(26, 267);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 14);
            this.label8.TabIndex = 0;
            this.label8.Text = "Deck/Hold";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 122);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 14);
            this.label4.TabIndex = 0;
            this.label4.Text = "Reefer Basis";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 14);
            this.label3.TabIndex = 0;
            this.label3.Text = "Max Gang Count";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(123, 14);
            this.label2.TabIndex = 0;
            this.label2.Text = "Max RDB Gang Count";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "Max AF";
            // 
            // btn_update_data
            // 
            this.btn_update_data.Location = new System.Drawing.Point(268, 24);
            this.btn_update_data.Name = "btn_update_data";
            this.btn_update_data.Size = new System.Drawing.Size(87, 23);
            this.btn_update_data.TabIndex = 5;
            this.btn_update_data.Text = "Update Data";
            this.btn_update_data.UseVisualStyleBackColor = true;
            this.btn_update_data.Click += new System.EventHandler(this.btn_update_data_Click);
            // 
            // btn_change_color
            // 
            this.btn_change_color.Location = new System.Drawing.Point(363, 24);
            this.btn_change_color.Name = "btn_change_color";
            this.btn_change_color.Size = new System.Drawing.Size(87, 23);
            this.btn_change_color.TabIndex = 10;
            this.btn_change_color.Text = "Change Color";
            this.btn_change_color.UseVisualStyleBackColor = true;
            this.btn_change_color.Click += new System.EventHandler(this.btn_change_color_Click);
            // 
            // simpleButton1
            // 
            this.simpleButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.simpleButton1.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.ImageOptions.Image")));
            this.simpleButton1.Location = new System.Drawing.Point(762, 13);
            this.simpleButton1.Name = "simpleButton1";
            this.simpleButton1.PaintStyle = DevExpress.XtraEditors.Controls.PaintStyles.Light;
            this.simpleButton1.Size = new System.Drawing.Size(39, 34);
            this.simpleButton1.TabIndex = 6;
            this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // spreadsheetControl1
            // 
            this.spreadsheetControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.spreadsheetControl1.Location = new System.Drawing.Point(12, 53);
            this.spreadsheetControl1.Name = "spreadsheetControl1";
            this.spreadsheetControl1.Size = new System.Drawing.Size(789, 516);
            this.spreadsheetControl1.TabIndex = 7;
            this.spreadsheetControl1.Text = "spreadsheetControl1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(187, 24);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_rebuild
            // 
            this.btn_rebuild.Location = new System.Drawing.Point(466, 24);
            this.btn_rebuild.Name = "btn_rebuild";
            this.btn_rebuild.Size = new System.Drawing.Size(87, 23);
            this.btn_rebuild.TabIndex = 5;
            this.btn_rebuild.Text = "ReBuild";
            this.btn_rebuild.UseVisualStyleBackColor = true;
            this.btn_rebuild.Visible = false;
            this.btn_rebuild.Click += new System.EventHandler(this.btn_rebuild_Click);
            // 
            // btn_callworksheet
            // 
            this.btn_callworksheet.Location = new System.Drawing.Point(559, 24);
            this.btn_callworksheet.Name = "btn_callworksheet";
            this.btn_callworksheet.Size = new System.Drawing.Size(75, 23);
            this.btn_callworksheet.TabIndex = 9;
            this.btn_callworksheet.Text = "WorkSheet";
            this.btn_callworksheet.UseVisualStyleBackColor = true;
            this.btn_callworksheet.Click += new System.EventHandler(this.btn_callworksheet_Click);
            // 
            // lb_projectname
            // 
            this.lb_projectname.AutoSize = true;
            this.lb_projectname.Location = new System.Drawing.Point(75, 2);
            this.lb_projectname.Name = "lb_projectname";
            this.lb_projectname.Size = new System.Drawing.Size(11, 14);
            this.lb_projectname.TabIndex = 40;
            this.lb_projectname.Text = "-";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 2);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 14);
            this.label5.TabIndex = 41;
            this.label5.Text = "Project: ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(813, 581);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lb_projectname);
            this.Controls.Add(this.btn_callworksheet);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.simpleButton1);
            this.Controls.Add(this.btn_rebuild);
            this.Controls.Add(this.btn_change_color);
            this.Controls.Add(this.btn_update_data);
            this.Controls.Add(this.panel_setting);
            this.Controls.Add(this.btn_setting);
            this.Controls.Add(this.gridControl1);
            this.Controls.Add(this.btn_create_data);
            this.Controls.Add(this.btn_excel_import);
            this.Controls.Add(this.spreadsheetControl1);
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            this.panel_setting.ResumeLayout(false);
            this.panel_setting.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.num_generator_basis)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_srd_basis)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_transformer_basis)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_distboard_basis)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_TR)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_SRD)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_RDB)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Reefer_basis)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Max_gang_cnt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Max_rdb_reefer_cnt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Max_AF)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_excel_import;
        private System.Windows.Forms.Button btn_create_data;
        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private System.Windows.Forms.Button btn_setting;

        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.Panel panel_setting;
        public System.Windows.Forms.NumericUpDown num_Max_gang_cnt;
        public System.Windows.Forms.NumericUpDown num_Max_rdb_reefer_cnt;
        public System.Windows.Forms.NumericUpDown num_Max_AF;
        public System.Windows.Forms.ComboBox cb_deck_hold_type;
        private System.Windows.Forms.CheckBox chk_deck_first;
        private System.Windows.Forms.Label label8;
        public System.Windows.Forms.NumericUpDown num_Reefer_basis;

        private System.Windows.Forms.Button btn_update_data;
        private System.Windows.Forms.Button btn_change_color;
        private System.Windows.Forms.Button btn_save_setting;
        private System.Windows.Forms.Button btn_default_setting;
        private DevExpress.XtraEditors.SimpleButton simpleButton1;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControl1;
        public System.Windows.Forms.NumericUpDown num_distboard_basis;
        private System.Windows.Forms.Label label9;
        public System.Windows.Forms.NumericUpDown num_transformer_basis;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.NumericUpDown num_srd_basis;
        private System.Windows.Forms.NumericUpDown num_generator_basis;
        private System.Windows.Forms.Label lbl_rdb_result;
        private System.Windows.Forms.Label lbl_srd_result;
        private System.Windows.Forms.Label lbl_tr_result;
        private System.Windows.Forms.Label lbl_tr_count;
        private System.Windows.Forms.Label label_max_af_unit;
        private System.Windows.Forms.Label label_max_rdb_unit;
        private System.Windows.Forms.Label label_max_gang_unit;
        private System.Windows.Forms.Label label_reefer_unit;
        private System.Windows.Forms.Label label_rdb_unit;
        private System.Windows.Forms.Label label_srd_unit;
        private System.Windows.Forms.Label label_tr_unit;
        private System.Windows.Forms.Label label_gen_unit;
        private System.Windows.Forms.Label label_srd_basis;
        private System.Windows.Forms.Label label_generator_basis;
        private System.Windows.Forms.Label label_tr_count_title;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btn_rebuild;
        private System.Windows.Forms.Button btn_callworksheet;
        private System.Windows.Forms.Label lb_projectname;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        public System.Windows.Forms.NumericUpDown num_TR;
        public System.Windows.Forms.NumericUpDown num_SRD;
        public System.Windows.Forms.NumericUpDown num_RDB;
        public SortPriorityPanel sortPriorityPanel;
        private System.Windows.Forms.TextBox tb_revision;
        private System.Windows.Forms.Label label11;
    }
}

