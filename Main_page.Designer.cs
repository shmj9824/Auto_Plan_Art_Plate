namespace Auto_Plan_Art_Plate
{
    partial class Main_page
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.bt_sel_art = new System.Windows.Forms.Button();
            this.bt_sel_plate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.nud_plate_row = new System.Windows.Forms.NumericUpDown();
            this.nud_plate_col = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.bt_plan_start = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.cb_Rotate = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tb_art_plan_explan = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.bt_sel_plan_excel = new System.Windows.Forms.Button();
            this.openFileDialog3 = new System.Windows.Forms.OpenFileDialog();
            this.label_file_art = new System.Windows.Forms.Label();
            this.label_file_plate = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label28 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.CB_id5 = new System.Windows.Forms.CheckBox();
            this.cb_back_plate = new System.Windows.Forms.CheckBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.cb_dataset = new System.Windows.Forms.ComboBox();
            this.bt_VariPrint = new System.Windows.Forms.Button();
            this.bt_PV_table = new System.Windows.Forms.Button();
            this.bt_exam_art = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tp_VP = new System.Windows.Forms.TabPage();
            this.cb_year_p = new System.Windows.Forms.CheckBox();
            this.bt_Reference = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.cb_print_way = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cb_vp_dataset = new System.Windows.Forms.ComboBox();
            this.lab_vp = new System.Windows.Forms.Label();
            this.lab_dt = new System.Windows.Forms.Label();
            this.tp_merge = new System.Windows.Forms.TabPage();
            this.label5 = new System.Windows.Forms.Label();
            this.tp_merge_tb = new System.Windows.Forms.TabPage();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label12 = new System.Windows.Forms.Label();
            this.cb_ziener_acronym = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.tb_care_num_z = new System.Windows.Forms.TextBox();
            this.bt_import_exam = new System.Windows.Forms.Button();
            this.cb_sel_ziener = new System.Windows.Forms.ComboBox();
            this.bt_ziener_refer_excel = new System.Windows.Forms.Button();
            this.bt_ziener_tran = new System.Windows.Forms.Button();
            this.tb_ziener_out = new System.Windows.Forms.TextBox();
            this.tb_ziener_input = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tb_care_num_v = new System.Windows.Forms.TextBox();
            this.cb_sel_vaude = new System.Windows.Forms.ComboBox();
            this.bt_str_apply = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.tb_Vaude_out = new System.Windows.Forms.TextBox();
            this.tb_Vaude_input = new System.Windows.Forms.TextBox();
            this.bt_vaude_refer_excel = new System.Windows.Forms.Button();
            this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
            this.CB_oldver = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nud_plate_row)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nud_plate_col)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tp_VP.SuspendLayout();
            this.tp_merge.SuspendLayout();
            this.tp_merge_tb.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // bt_sel_art
            // 
            this.bt_sel_art.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_sel_art.Location = new System.Drawing.Point(6, 6);
            this.bt_sel_art.Name = "bt_sel_art";
            this.bt_sel_art.Size = new System.Drawing.Size(90, 35);
            this.bt_sel_art.TabIndex = 0;
            this.bt_sel_art.Text = "選擇畫稿";
            this.bt_sel_art.UseVisualStyleBackColor = true;
            this.bt_sel_art.Click += new System.EventHandler(this.bt_sel_art_Click);
            // 
            // bt_sel_plate
            // 
            this.bt_sel_plate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_sel_plate.Location = new System.Drawing.Point(6, 43);
            this.bt_sel_plate.Name = "bt_sel_plate";
            this.bt_sel_plate.Size = new System.Drawing.Size(90, 35);
            this.bt_sel_plate.TabIndex = 1;
            this.bt_sel_plate.Text = "選擇模板";
            this.bt_sel_plate.UseVisualStyleBackColor = true;
            this.bt_sel_plate.Click += new System.EventHandler(this.bt_sel_plate_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(3, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "模板列(橫)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(151, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "模板行(直)";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.nud_plate_row);
            this.panel1.Controls.Add(this.nud_plate_col);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(164, 79);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(297, 45);
            this.panel1.TabIndex = 6;
            // 
            // nud_plate_row
            // 
            this.nud_plate_row.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.nud_plate_row.Location = new System.Drawing.Point(90, 11);
            this.nud_plate_row.Name = "nud_plate_row";
            this.nud_plate_row.Size = new System.Drawing.Size(49, 27);
            this.nud_plate_row.TabIndex = 7;
            this.nud_plate_row.Value = new decimal(new int[] {
            7,
            0,
            0,
            0});
            // 
            // nud_plate_col
            // 
            this.nud_plate_col.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.nud_plate_col.Location = new System.Drawing.Point(238, 11);
            this.nud_plate_col.Name = "nud_plate_col";
            this.nud_plate_col.Size = new System.Drawing.Size(49, 27);
            this.nud_plate_col.TabIndex = 6;
            this.nud_plate_col.Value = new decimal(new int[] {
            7,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(3, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "原稿旋轉";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // bt_plan_start
            // 
            this.bt_plan_start.Enabled = false;
            this.bt_plan_start.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_plan_start.Location = new System.Drawing.Point(547, 47);
            this.bt_plan_start.Name = "bt_plan_start";
            this.bt_plan_start.Size = new System.Drawing.Size(100, 35);
            this.bt_plan_start.TabIndex = 9;
            this.bt_plan_start.Text = "開始";
            this.bt_plan_start.UseVisualStyleBackColor = true;
            this.bt_plan_start.Click += new System.EventHandler(this.bt_plan_start_Click);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // cb_Rotate
            // 
            this.cb_Rotate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cb_Rotate.FormattingEnabled = true;
            this.cb_Rotate.Items.AddRange(new object[] {
            "0",
            "90"});
            this.cb_Rotate.Location = new System.Drawing.Point(80, 10);
            this.cb_Rotate.Name = "cb_Rotate";
            this.cb_Rotate.Size = new System.Drawing.Size(63, 24);
            this.cb_Rotate.TabIndex = 10;
            this.cb_Rotate.Text = "90";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.cb_Rotate);
            this.panel2.Location = new System.Drawing.Point(3, 79);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(155, 45);
            this.panel2.TabIndex = 11;
            // 
            // tb_art_plan_explan
            // 
            this.tb_art_plan_explan.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tb_art_plan_explan.Location = new System.Drawing.Point(10, 30);
            this.tb_art_plan_explan.Name = "tb_art_plan_explan";
            this.tb_art_plan_explan.Size = new System.Drawing.Size(517, 27);
            this.tb_art_plan_explan.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(5, 5);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 16);
            this.label4.TabIndex = 11;
            this.label4.Text = "模板說明";
            // 
            // bt_sel_plan_excel
            // 
            this.bt_sel_plan_excel.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_sel_plan_excel.Location = new System.Drawing.Point(225, 104);
            this.bt_sel_plan_excel.Name = "bt_sel_plan_excel";
            this.bt_sel_plan_excel.Size = new System.Drawing.Size(100, 35);
            this.bt_sel_plan_excel.TabIndex = 12;
            this.bt_sel_plan_excel.Text = "選擇拼模表";
            this.bt_sel_plan_excel.UseVisualStyleBackColor = true;
            this.bt_sel_plan_excel.Click += new System.EventHandler(this.bt_sel_plan_excel_Click);
            // 
            // openFileDialog3
            // 
            this.openFileDialog3.FileName = "openFileDialog3";
            // 
            // label_file_art
            // 
            this.label_file_art.AutoSize = true;
            this.label_file_art.Location = new System.Drawing.Point(102, 18);
            this.label_file_art.Name = "label_file_art";
            this.label_file_art.Size = new System.Drawing.Size(87, 16);
            this.label_file_art.TabIndex = 14;
            this.label_file_art.Text = "畫稿檔路徑";
            // 
            // label_file_plate
            // 
            this.label_file_plate.AutoSize = true;
            this.label_file_plate.Location = new System.Drawing.Point(102, 55);
            this.label_file_plate.Name = "label_file_plate";
            this.label_file_plate.Size = new System.Drawing.Size(87, 16);
            this.label_file_plate.TabIndex = 15;
            this.label_file_plate.Text = "模板檔路徑";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label28.Location = new System.Drawing.Point(1, 351);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(150, 16);
            this.label28.TabIndex = 38;
            this.label28.Text = "畫稿自動拼模-V1.3";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.CB_id5);
            this.panel3.Controls.Add(this.cb_back_plate);
            this.panel3.Controls.Add(this.panel4);
            this.panel3.Controls.Add(this.panel1);
            this.panel3.Controls.Add(this.panel2);
            this.panel3.Controls.Add(this.bt_plan_start);
            this.panel3.Location = new System.Drawing.Point(7, 156);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(659, 138);
            this.panel3.TabIndex = 39;
            // 
            // CB_id5
            // 
            this.CB_id5.AutoSize = true;
            this.CB_id5.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.CB_id5.Location = new System.Drawing.Point(556, 104);
            this.CB_id5.Name = "CB_id5";
            this.CB_id5.Size = new System.Drawing.Size(88, 20);
            this.CB_id5.TabIndex = 16;
            this.CB_id5.Text = "Reiid-5舊";
            this.CB_id5.UseVisualStyleBackColor = true;
            // 
            // cb_back_plate
            // 
            this.cb_back_plate.AutoSize = true;
            this.cb_back_plate.Enabled = false;
            this.cb_back_plate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cb_back_plate.Location = new System.Drawing.Point(472, 93);
            this.cb_back_plate.Name = "cb_back_plate";
            this.cb_back_plate.Size = new System.Drawing.Size(58, 20);
            this.cb_back_plate.TabIndex = 16;
            this.cb_back_plate.Text = "背板";
            this.cb_back_plate.UseVisualStyleBackColor = true;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.tb_art_plan_explan);
            this.panel4.Controls.Add(this.label4);
            this.panel4.Location = new System.Drawing.Point(3, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(538, 70);
            this.panel4.TabIndex = 15;
            // 
            // cb_dataset
            // 
            this.cb_dataset.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cb_dataset.FormattingEnabled = true;
            this.cb_dataset.Location = new System.Drawing.Point(7, 112);
            this.cb_dataset.Name = "cb_dataset";
            this.cb_dataset.Size = new System.Drawing.Size(198, 27);
            this.cb_dataset.TabIndex = 40;
            this.cb_dataset.SelectedValueChanged += new System.EventHandler(this.cb_dataset_SelectedValueChanged);
            // 
            // bt_VariPrint
            // 
            this.bt_VariPrint.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_VariPrint.Location = new System.Drawing.Point(222, 136);
            this.bt_VariPrint.Name = "bt_VariPrint";
            this.bt_VariPrint.Size = new System.Drawing.Size(110, 35);
            this.bt_VariPrint.TabIndex = 43;
            this.bt_VariPrint.Text = "火併";
            this.bt_VariPrint.UseVisualStyleBackColor = true;
            this.bt_VariPrint.Click += new System.EventHandler(this.bt_VariPrint_Click);
            // 
            // bt_PV_table
            // 
            this.bt_PV_table.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_PV_table.Location = new System.Drawing.Point(6, 6);
            this.bt_PV_table.Name = "bt_PV_table";
            this.bt_PV_table.Size = new System.Drawing.Size(110, 35);
            this.bt_PV_table.TabIndex = 44;
            this.bt_PV_table.Text = "資料表";
            this.bt_PV_table.UseVisualStyleBackColor = true;
            this.bt_PV_table.Click += new System.EventHandler(this.bt_PV_table_Click);
            // 
            // bt_exam_art
            // 
            this.bt_exam_art.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_exam_art.Location = new System.Drawing.Point(6, 66);
            this.bt_exam_art.Name = "bt_exam_art";
            this.bt_exam_art.Size = new System.Drawing.Size(110, 35);
            this.bt_exam_art.TabIndex = 44;
            this.bt_exam_art.Text = "母版";
            this.bt_exam_art.UseVisualStyleBackColor = true;
            this.bt_exam_art.Click += new System.EventHandler(this.bt_exam_art_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tp_VP);
            this.tabControl1.Controls.Add(this.tp_merge);
            this.tabControl1.Controls.Add(this.tp_merge_tb);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(695, 348);
            this.tabControl1.TabIndex = 46;
            // 
            // tp_VP
            // 
            this.tp_VP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tp_VP.Controls.Add(this.CB_oldver);
            this.tp_VP.Controls.Add(this.cb_year_p);
            this.tp_VP.Controls.Add(this.bt_Reference);
            this.tp_VP.Controls.Add(this.label7);
            this.tp_VP.Controls.Add(this.cb_print_way);
            this.tp_VP.Controls.Add(this.label6);
            this.tp_VP.Controls.Add(this.cb_vp_dataset);
            this.tp_VP.Controls.Add(this.lab_vp);
            this.tp_VP.Controls.Add(this.lab_dt);
            this.tp_VP.Controls.Add(this.bt_exam_art);
            this.tp_VP.Controls.Add(this.bt_VariPrint);
            this.tp_VP.Controls.Add(this.bt_PV_table);
            this.tp_VP.Location = new System.Drawing.Point(4, 29);
            this.tp_VP.Name = "tp_VP";
            this.tp_VP.Padding = new System.Windows.Forms.Padding(3);
            this.tp_VP.Size = new System.Drawing.Size(687, 315);
            this.tp_VP.TabIndex = 1;
            this.tp_VP.Text = "火併";
            this.tp_VP.UseVisualStyleBackColor = true;
            // 
            // cb_year_p
            // 
            this.cb_year_p.AutoSize = true;
            this.cb_year_p.Enabled = false;
            this.cb_year_p.Location = new System.Drawing.Point(153, 248);
            this.cb_year_p.Name = "cb_year_p";
            this.cb_year_p.Size = new System.Drawing.Size(90, 20);
            this.cb_year_p.TabIndex = 53;
            this.cb_year_p.Text = "年份增加";
            this.cb_year_p.UseVisualStyleBackColor = true;
            // 
            // bt_Reference
            // 
            this.bt_Reference.Enabled = false;
            this.bt_Reference.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.bt_Reference.Location = new System.Drawing.Point(142, 194);
            this.bt_Reference.Name = "bt_Reference";
            this.bt_Reference.Size = new System.Drawing.Size(110, 35);
            this.bt_Reference.TabIndex = 52;
            this.bt_Reference.Text = "參考檔";
            this.bt_Reference.UseVisualStyleBackColor = true;
            this.bt_Reference.Click += new System.EventHandler(this.bt_Reference_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(7, 183);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 16);
            this.label7.TabIndex = 50;
            this.label7.Text = "方式";
            // 
            // cb_print_way
            // 
            this.cb_print_way.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cb_print_way.FormattingEnabled = true;
            this.cb_print_way.Items.AddRange(new object[] {
            "核稿",
            "備拼檔"});
            this.cb_print_way.Location = new System.Drawing.Point(7, 202);
            this.cb_print_way.Name = "cb_print_way";
            this.cb_print_way.Size = new System.Drawing.Size(109, 27);
            this.cb_print_way.TabIndex = 49;
            this.cb_print_way.Text = "核稿";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 121);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(87, 16);
            this.label6.TabIndex = 48;
            this.label6.Text = "選擇資料集";
            // 
            // cb_vp_dataset
            // 
            this.cb_vp_dataset.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cb_vp_dataset.FormattingEnabled = true;
            this.cb_vp_dataset.Location = new System.Drawing.Point(7, 140);
            this.cb_vp_dataset.Name = "cb_vp_dataset";
            this.cb_vp_dataset.Size = new System.Drawing.Size(198, 27);
            this.cb_vp_dataset.TabIndex = 47;
            this.cb_vp_dataset.SelectedValueChanged += new System.EventHandler(this.cb_vp_dataset_SelectedChanged);
            // 
            // lab_vp
            // 
            this.lab_vp.AutoSize = true;
            this.lab_vp.Location = new System.Drawing.Point(122, 75);
            this.lab_vp.Name = "lab_vp";
            this.lab_vp.Size = new System.Drawing.Size(71, 16);
            this.lab_vp.TabIndex = 46;
            this.lab_vp.Text = "母版路徑";
            // 
            // lab_dt
            // 
            this.lab_dt.AutoSize = true;
            this.lab_dt.Location = new System.Drawing.Point(122, 15);
            this.lab_dt.Name = "lab_dt";
            this.lab_dt.Size = new System.Drawing.Size(87, 16);
            this.lab_dt.TabIndex = 45;
            this.lab_dt.Text = "資料表路徑";
            // 
            // tp_merge
            // 
            this.tp_merge.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tp_merge.Controls.Add(this.label5);
            this.tp_merge.Controls.Add(this.bt_sel_art);
            this.tp_merge.Controls.Add(this.bt_sel_plate);
            this.tp_merge.Controls.Add(this.bt_sel_plan_excel);
            this.tp_merge.Controls.Add(this.cb_dataset);
            this.tp_merge.Controls.Add(this.label_file_art);
            this.tp_merge.Controls.Add(this.panel3);
            this.tp_merge.Controls.Add(this.label_file_plate);
            this.tp_merge.Location = new System.Drawing.Point(4, 29);
            this.tp_merge.Name = "tp_merge";
            this.tp_merge.Padding = new System.Windows.Forms.Padding(3);
            this.tp_merge.Size = new System.Drawing.Size(687, 315);
            this.tp_merge.TabIndex = 0;
            this.tp_merge.Text = "合板";
            this.tp_merge.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 89);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(87, 16);
            this.label5.TabIndex = 41;
            this.label5.Text = "選擇資料集";
            // 
            // tp_merge_tb
            // 
            this.tp_merge_tb.Controls.Add(this.button2);
            this.tp_merge_tb.Controls.Add(this.comboBox1);
            this.tp_merge_tb.Location = new System.Drawing.Point(4, 29);
            this.tp_merge_tb.Name = "tp_merge_tb";
            this.tp_merge_tb.Size = new System.Drawing.Size(687, 315);
            this.tp_merge_tb.TabIndex = 2;
            this.tp_merge_tb.Text = "合板(進階)";
            this.tp_merge_tb.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.button2.Location = new System.Drawing.Point(234, 55);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(100, 35);
            this.button2.TabIndex = 41;
            this.button2.Text = "選擇拼模表";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.comboBox1.Enabled = false;
            this.comboBox1.Font = new System.Drawing.Font("新細明體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(18, 59);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(198, 27);
            this.comboBox1.TabIndex = 42;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.cb_ziener_acronym);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.tb_care_num_z);
            this.tabPage1.Controls.Add(this.bt_import_exam);
            this.tabPage1.Controls.Add(this.cb_sel_ziener);
            this.tabPage1.Controls.Add(this.bt_ziener_refer_excel);
            this.tabPage1.Controls.Add(this.bt_ziener_tran);
            this.tabPage1.Controls.Add(this.tb_ziener_out);
            this.tabPage1.Controls.Add(this.tb_ziener_input);
            this.tabPage1.Location = new System.Drawing.Point(4, 29);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(687, 315);
            this.tabPage1.TabIndex = 3;
            this.tabPage1.Text = "Ziener字串";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(366, 25);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(39, 16);
            this.label12.TabIndex = 9;
            this.label12.Text = "選項";
            // 
            // cb_ziener_acronym
            // 
            this.cb_ziener_acronym.AutoSize = true;
            this.cb_ziener_acronym.Location = new System.Drawing.Point(133, 24);
            this.cb_ziener_acronym.Name = "cb_ziener_acronym";
            this.cb_ziener_acronym.Size = new System.Drawing.Size(58, 20);
            this.cb_ziener_acronym.TabIndex = 8;
            this.cb_ziener_acronym.Text = "縮寫";
            this.cb_ziener_acronym.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(495, 73);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(71, 16);
            this.label9.TabIndex = 7;
            this.label9.Text = "輸出字串";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(139, 73);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(71, 16);
            this.label8.TabIndex = 7;
            this.label8.Text = "輸入字串";
            // 
            // tb_care_num_z
            // 
            this.tb_care_num_z.Location = new System.Drawing.Point(80, 285);
            this.tb_care_num_z.Name = "tb_care_num_z";
            this.tb_care_num_z.Size = new System.Drawing.Size(393, 27);
            this.tb_care_num_z.TabIndex = 6;
            // 
            // bt_import_exam
            // 
            this.bt_import_exam.Location = new System.Drawing.Point(207, 15);
            this.bt_import_exam.Name = "bt_import_exam";
            this.bt_import_exam.Size = new System.Drawing.Size(89, 36);
            this.bt_import_exam.TabIndex = 5;
            this.bt_import_exam.Text = "匯入Excel";
            this.bt_import_exam.UseVisualStyleBackColor = true;
            this.bt_import_exam.Click += new System.EventHandler(this.bt_import_exam_Click);
            // 
            // cb_sel_ziener
            // 
            this.cb_sel_ziener.FormattingEnabled = true;
            this.cb_sel_ziener.Items.AddRange(new object[] {
            "成份",
            "洗語"});
            this.cb_sel_ziener.Location = new System.Drawing.Point(411, 20);
            this.cb_sel_ziener.Name = "cb_sel_ziener";
            this.cb_sel_ziener.Size = new System.Drawing.Size(121, 24);
            this.cb_sel_ziener.TabIndex = 4;
            this.cb_sel_ziener.SelectedIndexChanged += new System.EventHandler(this.Cb_sel_ziener_SelectedValueChanged);
            // 
            // bt_ziener_refer_excel
            // 
            this.bt_ziener_refer_excel.Location = new System.Drawing.Point(8, 13);
            this.bt_ziener_refer_excel.Name = "bt_ziener_refer_excel";
            this.bt_ziener_refer_excel.Size = new System.Drawing.Size(89, 36);
            this.bt_ziener_refer_excel.TabIndex = 3;
            this.bt_ziener_refer_excel.Text = "讀檔";
            this.bt_ziener_refer_excel.UseVisualStyleBackColor = true;
            this.bt_ziener_refer_excel.Click += new System.EventHandler(this.bt_ziener_excel_Click);
            // 
            // bt_ziener_tran
            // 
            this.bt_ziener_tran.Location = new System.Drawing.Point(538, 39);
            this.bt_ziener_tran.Name = "bt_ziener_tran";
            this.bt_ziener_tran.Size = new System.Drawing.Size(91, 31);
            this.bt_ziener_tran.TabIndex = 2;
            this.bt_ziener_tran.Text = "產生";
            this.bt_ziener_tran.UseVisualStyleBackColor = true;
            this.bt_ziener_tran.Click += new System.EventHandler(this.bt_tran_Click);
            // 
            // tb_ziener_out
            // 
            this.tb_ziener_out.Enabled = false;
            this.tb_ziener_out.Location = new System.Drawing.Point(375, 92);
            this.tb_ziener_out.Multiline = true;
            this.tb_ziener_out.Name = "tb_ziener_out";
            this.tb_ziener_out.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_ziener_out.Size = new System.Drawing.Size(309, 187);
            this.tb_ziener_out.TabIndex = 1;
            // 
            // tb_ziener_input
            // 
            this.tb_ziener_input.Enabled = false;
            this.tb_ziener_input.Location = new System.Drawing.Point(8, 92);
            this.tb_ziener_input.Multiline = true;
            this.tb_ziener_input.Name = "tb_ziener_input";
            this.tb_ziener_input.Size = new System.Drawing.Size(361, 187);
            this.tb_ziener_input.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.tb_care_num_v);
            this.tabPage2.Controls.Add(this.cb_sel_vaude);
            this.tabPage2.Controls.Add(this.bt_str_apply);
            this.tabPage2.Controls.Add(this.label10);
            this.tabPage2.Controls.Add(this.label11);
            this.tabPage2.Controls.Add(this.tb_Vaude_out);
            this.tabPage2.Controls.Add(this.tb_Vaude_input);
            this.tabPage2.Controls.Add(this.bt_vaude_refer_excel);
            this.tabPage2.Location = new System.Drawing.Point(4, 29);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(687, 315);
            this.tabPage2.TabIndex = 4;
            this.tabPage2.Text = "Vaude字串";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tb_care_num_v
            // 
            this.tb_care_num_v.Location = new System.Drawing.Point(78, 285);
            this.tb_care_num_v.Name = "tb_care_num_v";
            this.tb_care_num_v.Size = new System.Drawing.Size(393, 27);
            this.tb_care_num_v.TabIndex = 14;
            // 
            // cb_sel_vaude
            // 
            this.cb_sel_vaude.FormattingEnabled = true;
            this.cb_sel_vaude.Items.AddRange(new object[] {
            "成份",
            "洗語"});
            this.cb_sel_vaude.Location = new System.Drawing.Point(19, 25);
            this.cb_sel_vaude.Name = "cb_sel_vaude";
            this.cb_sel_vaude.Size = new System.Drawing.Size(121, 24);
            this.cb_sel_vaude.TabIndex = 13;
            this.cb_sel_vaude.SelectedValueChanged += new System.EventHandler(this.Cb_sel_vaude_SelectedValueChanged);
            // 
            // bt_str_apply
            // 
            this.bt_str_apply.Location = new System.Drawing.Point(382, 18);
            this.bt_str_apply.Name = "bt_str_apply";
            this.bt_str_apply.Size = new System.Drawing.Size(89, 36);
            this.bt_str_apply.TabIndex = 12;
            this.bt_str_apply.Text = "執行";
            this.bt_str_apply.UseVisualStyleBackColor = true;
            this.bt_str_apply.Click += new System.EventHandler(this.bt_str_apply_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(474, 73);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(71, 16);
            this.label10.TabIndex = 10;
            this.label10.Text = "輸出字串";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(128, 73);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(71, 16);
            this.label11.TabIndex = 11;
            this.label11.Text = "輸入字串";
            // 
            // tb_Vaude_out
            // 
            this.tb_Vaude_out.Location = new System.Drawing.Point(367, 92);
            this.tb_Vaude_out.Multiline = true;
            this.tb_Vaude_out.Name = "tb_Vaude_out";
            this.tb_Vaude_out.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_Vaude_out.Size = new System.Drawing.Size(307, 187);
            this.tb_Vaude_out.TabIndex = 9;
            // 
            // tb_Vaude_input
            // 
            this.tb_Vaude_input.Location = new System.Drawing.Point(19, 93);
            this.tb_Vaude_input.Multiline = true;
            this.tb_Vaude_input.Name = "tb_Vaude_input";
            this.tb_Vaude_input.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tb_Vaude_input.Size = new System.Drawing.Size(342, 187);
            this.tb_Vaude_input.TabIndex = 8;
            // 
            // bt_vaude_refer_excel
            // 
            this.bt_vaude_refer_excel.Enabled = false;
            this.bt_vaude_refer_excel.Location = new System.Drawing.Point(168, 18);
            this.bt_vaude_refer_excel.Name = "bt_vaude_refer_excel";
            this.bt_vaude_refer_excel.Size = new System.Drawing.Size(89, 36);
            this.bt_vaude_refer_excel.TabIndex = 4;
            this.bt_vaude_refer_excel.Text = "讀檔";
            this.bt_vaude_refer_excel.UseVisualStyleBackColor = true;
            this.bt_vaude_refer_excel.Click += new System.EventHandler(this.bt_vaude_refer_excel_Click);
            // 
            // CB_oldver
            // 
            this.CB_oldver.AutoSize = true;
            this.CB_oldver.Location = new System.Drawing.Point(275, 202);
            this.CB_oldver.Name = "CB_oldver";
            this.CB_oldver.Size = new System.Drawing.Size(90, 20);
            this.CB_oldver.TabIndex = 53;
            this.CB_oldver.Text = "舊版相容";
            this.CB_oldver.UseVisualStyleBackColor = true;
            // 
            // Main_page
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(707, 377);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label28);
            this.Name = "Main_page";
            this.Text = "Auto_Plan_Art_Plate";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nud_plate_row)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nud_plate_col)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tp_VP.ResumeLayout(false);
            this.tp_VP.PerformLayout();
            this.tp_merge.ResumeLayout(false);
            this.tp_merge.PerformLayout();
            this.tp_merge_tb.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bt_sel_art;
        private System.Windows.Forms.Button bt_sel_plate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button bt_plan_start;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.NumericUpDown nud_plate_col;
        private System.Windows.Forms.NumericUpDown nud_plate_row;
        private System.Windows.Forms.ComboBox cb_Rotate;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button bt_sel_plan_excel;
        private System.Windows.Forms.OpenFileDialog openFileDialog3;
        private System.Windows.Forms.Label label_file_art;
        private System.Windows.Forms.Label label_file_plate;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.TextBox tb_art_plan_explan;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.ComboBox cb_dataset;
        private System.Windows.Forms.CheckBox cb_back_plate;
        private System.Windows.Forms.Button bt_VariPrint;
        private System.Windows.Forms.Button bt_PV_table;
        private System.Windows.Forms.Button bt_exam_art;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tp_merge;
        private System.Windows.Forms.TabPage tp_VP;
        private System.Windows.Forms.Label lab_dt;
        private System.Windows.Forms.ComboBox cb_vp_dataset;
        private System.Windows.Forms.Label lab_vp;
        private System.Windows.Forms.TabPage tp_merge_tb;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cb_print_way;
        private System.Windows.Forms.Button bt_Reference;
        private System.Windows.Forms.CheckBox cb_year_p;
        private System.Windows.Forms.SaveFileDialog saveFileDialog2;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox tb_ziener_input;
        private System.Windows.Forms.Button bt_ziener_tran;
        private System.Windows.Forms.TextBox tb_ziener_out;
        private System.Windows.Forms.Button bt_ziener_refer_excel;
        private System.Windows.Forms.ComboBox cb_sel_ziener;
        private System.Windows.Forms.Button bt_import_exam;
        private System.Windows.Forms.TextBox tb_care_num_z;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button bt_vaude_refer_excel;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox tb_Vaude_out;
        private System.Windows.Forms.TextBox tb_Vaude_input;
        private System.Windows.Forms.Button bt_str_apply;
        private System.Windows.Forms.ComboBox cb_sel_vaude;
        private System.Windows.Forms.TextBox tb_care_num_v;
        private System.Windows.Forms.CheckBox cb_ziener_acronym;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.CheckBox CB_id5;
        private System.Windows.Forms.CheckBox CB_oldver;
    }
}

