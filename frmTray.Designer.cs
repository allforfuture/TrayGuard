namespace TrayGuard
{
    partial class frmTray
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTray));
            this.btnAddTray = new System.Windows.Forms.Button();
            this.btnSearchTray = new System.Windows.Forms.Button();
            this.dgvTray = new System.Windows.Forms.DataGridView();
            this.label12 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtTrayId = new System.Windows.Forms.TextBox();
            this.dtpRegsterDateFrom = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnAddReturn = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbLine = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbRegisterDept = new System.Windows.Forms.ComboBox();
            this.cbxHideCancel = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtLoginDept = new System.Windows.Forms.TextBox();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.cbxTrayId = new System.Windows.Forms.CheckBox();
            this.cbxRegisterDateFrom = new System.Windows.Forms.CheckBox();
            this.cbxRegisterDept = new System.Windows.Forms.CheckBox();
            this.cbxLine = new System.Windows.Forms.CheckBox();
            this.cbxShift = new System.Windows.Forms.CheckBox();
            this.dtpRegisterDateTo = new System.Windows.Forms.DateTimePicker();
            this.cbxRegisterDateTo = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.dtpUpdateDateFrom = new System.Windows.Forms.DateTimePicker();
            this.dtpUpdateDateTo = new System.Windows.Forms.DateTimePicker();
            this.cmbUpdateDept = new System.Windows.Forms.ComboBox();
            this.cbxUpdateDateFrom = new System.Windows.Forms.CheckBox();
            this.cbxUpdateDateTo = new System.Windows.Forms.CheckBox();
            this.cbxUpdateDept = new System.Windows.Forms.CheckBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtLot = new System.Windows.Forms.TextBox();
            this.cbxLot = new System.Windows.Forms.CheckBox();
            this.label14 = new System.Windows.Forms.Label();
            this.cbxMultiLot = new System.Windows.Forms.CheckBox();
            this.btnCancelMultiTray = new System.Windows.Forms.Button();
            this.label15 = new System.Windows.Forms.Label();
            this.txtModuleId = new System.Windows.Forms.TextBox();
            this.cbxModuleId = new System.Windows.Forms.CheckBox();
            this.txtShift = new System.Windows.Forms.TextBox();
            this.btnPartiallyCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTray)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAddTray
            // 
            this.btnAddTray.Location = new System.Drawing.Point(773, 137);
            this.btnAddTray.Name = "btnAddTray";
            this.btnAddTray.Size = new System.Drawing.Size(100, 23);
            this.btnAddTray.TabIndex = 6;
            this.btnAddTray.Text = "新增马达盆";
            this.btnAddTray.UseVisualStyleBackColor = true;
            this.btnAddTray.Click += new System.EventHandler(this.btnAddBoxId_Click);
            // 
            // btnSearchTray
            // 
            this.btnSearchTray.Location = new System.Drawing.Point(640, 137);
            this.btnSearchTray.Name = "btnSearchTray";
            this.btnSearchTray.Size = new System.Drawing.Size(100, 23);
            this.btnSearchTray.TabIndex = 2;
            this.btnSearchTray.Text = "搜索马达盆";
            this.btnSearchTray.UseVisualStyleBackColor = true;
            this.btnSearchTray.Click += new System.EventHandler(this.btnSearchTray_Click);
            // 
            // dgvTray
            // 
            this.dgvTray.AllowUserToAddRows = false;
            this.dgvTray.AllowUserToDeleteRows = false;
            this.dgvTray.BackgroundColor = System.Drawing.Color.LightSalmon;
            this.dgvTray.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTray.Location = new System.Drawing.Point(12, 175);
            this.dgvTray.Name = "dgvTray";
            this.dgvTray.ReadOnly = true;
            this.dgvTray.RowTemplate.Height = 21;
            this.dgvTray.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvTray.Size = new System.Drawing.Size(1032, 465);
            this.dgvTray.TabIndex = 9;
            this.dgvTray.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBoxId_CellContentClick);
            this.dgvTray.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTray_CellContentDoubleClick);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(64, 15);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(95, 12);
            this.label12.TabIndex = 6;
            this.label12.Text = "登记日期（起始）: ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(66, 73);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "马达盆ID: ";
            // 
            // txtTrayId
            // 
            this.txtTrayId.Location = new System.Drawing.Point(169, 71);
            this.txtTrayId.Name = "txtTrayId";
            this.txtTrayId.Size = new System.Drawing.Size(110, 19);
            this.txtTrayId.TabIndex = 1;
            // 
            // dtpRegsterDateFrom
            // 
            this.dtpRegsterDateFrom.CustomFormat = "yyyy/MM/dd";
            this.dtpRegsterDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegsterDateFrom.Location = new System.Drawing.Point(169, 12);
            this.dtpRegsterDateFrom.Name = "dtpRegsterDateFrom";
            this.dtpRegsterDateFrom.Size = new System.Drawing.Size(110, 19);
            this.dtpRegsterDateFrom.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(627, 104);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "登陆名: ";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(905, 137);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 23);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnAddReturn
            // 
            this.btnAddReturn.Enabled = false;
            this.btnAddReturn.Location = new System.Drawing.Point(905, 95);
            this.btnAddReturn.Name = "btnAddReturn";
            this.btnAddReturn.Size = new System.Drawing.Size(100, 23);
            this.btnAddReturn.TabIndex = 6;
            this.btnAddReturn.Text = "追加返回";
            this.btnAddReturn.UseVisualStyleBackColor = true;
            this.btnAddReturn.Visible = false;
            this.btnAddReturn.Click += new System.EventHandler(this.btnAddReturn_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(627, 15);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "登记部门: ";
            // 
            // cmbLine
            // 
            this.cmbLine.FormattingEnabled = true;
            this.cmbLine.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "A",
            "B",
            "C",
            "D",
            "E",
            "F",
            "G",
            "H",
            "I",
            "J",
            "K",
            "L",
            "M",
            "X"});
            this.cmbLine.Location = new System.Drawing.Point(445, 70);
            this.cmbLine.Name = "cmbLine";
            this.cmbLine.Size = new System.Drawing.Size(110, 20);
            this.cmbLine.TabIndex = 49;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(347, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "线别: ";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(348, 104);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "BIN, 班次: ";
            // 
            // cmbRegisterDept
            // 
            this.cmbRegisterDept.FormattingEnabled = true;
            this.cmbRegisterDept.Items.AddRange(new object[] {
            "MFG",
            "QA",
            "PC"});
            this.cmbRegisterDept.Location = new System.Drawing.Point(726, 12);
            this.cmbRegisterDept.Name = "cmbRegisterDept";
            this.cmbRegisterDept.Size = new System.Drawing.Size(110, 20);
            this.cmbRegisterDept.TabIndex = 49;
            // 
            // cbxHideCancel
            // 
            this.cbxHideCancel.AutoSize = true;
            this.cbxHideCancel.Location = new System.Drawing.Point(573, 138);
            this.cbxHideCancel.Name = "cbxHideCancel";
            this.cbxHideCancel.Size = new System.Drawing.Size(15, 14);
            this.cbxHideCancel.TabIndex = 50;
            this.cbxHideCancel.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(482, 140);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "隐藏取消: ";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(627, 76);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 6;
            this.label5.Text = "登陆部门: ";
            // 
            // txtLoginDept
            // 
            this.txtLoginDept.Location = new System.Drawing.Point(726, 71);
            this.txtLoginDept.Name = "txtLoginDept";
            this.txtLoginDept.ReadOnly = true;
            this.txtLoginDept.Size = new System.Drawing.Size(110, 19);
            this.txtLoginDept.TabIndex = 1;
            this.txtLoginDept.DoubleClick += new System.EventHandler(this.txtLoginDept_DoubleClick);
            // 
            // txtLoginName
            // 
            this.txtLoginName.Location = new System.Drawing.Point(726, 101);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.ReadOnly = true;
            this.txtLoginName.Size = new System.Drawing.Size(110, 19);
            this.txtLoginName.TabIndex = 1;
            this.txtLoginName.DoubleClick += new System.EventHandler(this.txtLoginName_DoubleClick);
            // 
            // cbxTrayId
            // 
            this.cbxTrayId.AutoSize = true;
            this.cbxTrayId.Location = new System.Drawing.Point(297, 74);
            this.cbxTrayId.Name = "cbxTrayId";
            this.cbxTrayId.Size = new System.Drawing.Size(15, 14);
            this.cbxTrayId.TabIndex = 50;
            this.cbxTrayId.UseVisualStyleBackColor = true;
            // 
            // cbxRegisterDateFrom
            // 
            this.cbxRegisterDateFrom.AutoSize = true;
            this.cbxRegisterDateFrom.Checked = true;
            this.cbxRegisterDateFrom.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbxRegisterDateFrom.Location = new System.Drawing.Point(297, 15);
            this.cbxRegisterDateFrom.Name = "cbxRegisterDateFrom";
            this.cbxRegisterDateFrom.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDateFrom.TabIndex = 50;
            this.cbxRegisterDateFrom.UseVisualStyleBackColor = true;
            // 
            // cbxRegisterDept
            // 
            this.cbxRegisterDept.AutoSize = true;
            this.cbxRegisterDept.Location = new System.Drawing.Point(851, 15);
            this.cbxRegisterDept.Name = "cbxRegisterDept";
            this.cbxRegisterDept.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDept.TabIndex = 50;
            this.cbxRegisterDept.UseVisualStyleBackColor = true;
            // 
            // cbxLine
            // 
            this.cbxLine.AutoSize = true;
            this.cbxLine.Location = new System.Drawing.Point(573, 73);
            this.cbxLine.Name = "cbxLine";
            this.cbxLine.Size = new System.Drawing.Size(15, 14);
            this.cbxLine.TabIndex = 50;
            this.cbxLine.UseVisualStyleBackColor = true;
            // 
            // cbxShift
            // 
            this.cbxShift.AutoSize = true;
            this.cbxShift.Location = new System.Drawing.Point(573, 105);
            this.cbxShift.Name = "cbxShift";
            this.cbxShift.Size = new System.Drawing.Size(15, 14);
            this.cbxShift.TabIndex = 50;
            this.cbxShift.UseVisualStyleBackColor = true;
            // 
            // dtpRegisterDateTo
            // 
            this.dtpRegisterDateTo.CustomFormat = "yyyy/MM/dd";
            this.dtpRegisterDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegisterDateTo.Location = new System.Drawing.Point(445, 12);
            this.dtpRegisterDateTo.Name = "dtpRegisterDateTo";
            this.dtpRegisterDateTo.Size = new System.Drawing.Size(110, 19);
            this.dtpRegisterDateTo.TabIndex = 10;
            // 
            // cbxRegisterDateTo
            // 
            this.cbxRegisterDateTo.AutoSize = true;
            this.cbxRegisterDateTo.Location = new System.Drawing.Point(573, 15);
            this.cbxRegisterDateTo.Name = "cbxRegisterDateTo";
            this.cbxRegisterDateTo.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDateTo.TabIndex = 50;
            this.cbxRegisterDateTo.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(346, 15);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(89, 12);
            this.label8.TabIndex = 6;
            this.label8.Text = "登记日期（截止: ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(66, 44);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(95, 12);
            this.label9.TabIndex = 6;
            this.label9.Text = "更新日期（起始）: ";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(346, 44);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(95, 12);
            this.label10.TabIndex = 6;
            this.label10.Text = "更新日期（截止）: ";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(627, 44);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(59, 12);
            this.label11.TabIndex = 6;
            this.label11.Text = "更新部门: ";
            // 
            // dtpUpdateDateFrom
            // 
            this.dtpUpdateDateFrom.CustomFormat = "yyyy/MM/dd";
            this.dtpUpdateDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpUpdateDateFrom.Location = new System.Drawing.Point(169, 41);
            this.dtpUpdateDateFrom.Name = "dtpUpdateDateFrom";
            this.dtpUpdateDateFrom.Size = new System.Drawing.Size(110, 19);
            this.dtpUpdateDateFrom.TabIndex = 10;
            // 
            // dtpUpdateDateTo
            // 
            this.dtpUpdateDateTo.CustomFormat = "yyyy/MM/dd";
            this.dtpUpdateDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpUpdateDateTo.Location = new System.Drawing.Point(445, 41);
            this.dtpUpdateDateTo.Name = "dtpUpdateDateTo";
            this.dtpUpdateDateTo.Size = new System.Drawing.Size(110, 19);
            this.dtpUpdateDateTo.TabIndex = 10;
            // 
            // cmbUpdateDept
            // 
            this.cmbUpdateDept.FormattingEnabled = true;
            this.cmbUpdateDept.Items.AddRange(new object[] {
            "MFG",
            "QA",
            "PC"});
            this.cmbUpdateDept.Location = new System.Drawing.Point(726, 41);
            this.cmbUpdateDept.Name = "cmbUpdateDept";
            this.cmbUpdateDept.Size = new System.Drawing.Size(110, 20);
            this.cmbUpdateDept.TabIndex = 49;
            // 
            // cbxUpdateDateFrom
            // 
            this.cbxUpdateDateFrom.AutoSize = true;
            this.cbxUpdateDateFrom.Location = new System.Drawing.Point(297, 44);
            this.cbxUpdateDateFrom.Name = "cbxUpdateDateFrom";
            this.cbxUpdateDateFrom.Size = new System.Drawing.Size(15, 14);
            this.cbxUpdateDateFrom.TabIndex = 50;
            this.cbxUpdateDateFrom.UseVisualStyleBackColor = true;
            // 
            // cbxUpdateDateTo
            // 
            this.cbxUpdateDateTo.AutoSize = true;
            this.cbxUpdateDateTo.Location = new System.Drawing.Point(573, 44);
            this.cbxUpdateDateTo.Name = "cbxUpdateDateTo";
            this.cbxUpdateDateTo.Size = new System.Drawing.Size(15, 14);
            this.cbxUpdateDateTo.TabIndex = 50;
            this.cbxUpdateDateTo.UseVisualStyleBackColor = true;
            // 
            // cbxUpdateDept
            // 
            this.cbxUpdateDept.AutoSize = true;
            this.cbxUpdateDept.Location = new System.Drawing.Point(851, 46);
            this.cbxUpdateDept.Name = "cbxUpdateDept";
            this.cbxUpdateDept.Size = new System.Drawing.Size(15, 14);
            this.cbxUpdateDept.TabIndex = 50;
            this.cbxUpdateDept.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(66, 104);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(35, 12);
            this.label13.TabIndex = 6;
            this.label13.Text = "批次: ";
            // 
            // txtLot
            // 
            this.txtLot.Location = new System.Drawing.Point(169, 102);
            this.txtLot.Name = "txtLot";
            this.txtLot.Size = new System.Drawing.Size(110, 19);
            this.txtLot.TabIndex = 1;
            // 
            // cbxLot
            // 
            this.cbxLot.AutoSize = true;
            this.cbxLot.Location = new System.Drawing.Point(297, 104);
            this.cbxLot.Name = "cbxLot";
            this.cbxLot.Size = new System.Drawing.Size(15, 14);
            this.cbxLot.TabIndex = 50;
            this.cbxLot.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(347, 140);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(59, 12);
            this.label14.TabIndex = 6;
            this.label14.Text = "多个批号: ";
            // 
            // cbxMultiLot
            // 
            this.cbxMultiLot.AutoSize = true;
            this.cbxMultiLot.Location = new System.Drawing.Point(428, 140);
            this.cbxMultiLot.Name = "cbxMultiLot";
            this.cbxMultiLot.Size = new System.Drawing.Size(15, 14);
            this.cbxMultiLot.TabIndex = 50;
            this.cbxMultiLot.UseVisualStyleBackColor = true;
            // 
            // btnCancelMultiTray
            // 
            this.btnCancelMultiTray.Enabled = false;
            this.btnCancelMultiTray.Location = new System.Drawing.Point(905, 16);
            this.btnCancelMultiTray.Name = "btnCancelMultiTray";
            this.btnCancelMultiTray.Size = new System.Drawing.Size(100, 23);
            this.btnCancelMultiTray.TabIndex = 6;
            this.btnCancelMultiTray.Text = "取消马达盆";
            this.btnCancelMultiTray.UseVisualStyleBackColor = true;
            this.btnCancelMultiTray.Visible = false;
            this.btnCancelMultiTray.Click += new System.EventHandler(this.btnCancelMultiTray_Click);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(66, 137);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(46, 12);
            this.label15.TabIndex = 6;
            this.label15.Text = "马达ID: ";
            // 
            // txtModuleId
            // 
            this.txtModuleId.Location = new System.Drawing.Point(169, 135);
            this.txtModuleId.Name = "txtModuleId";
            this.txtModuleId.Size = new System.Drawing.Size(110, 19);
            this.txtModuleId.TabIndex = 1;
            // 
            // cbxModuleId
            // 
            this.cbxModuleId.AutoSize = true;
            this.cbxModuleId.Location = new System.Drawing.Point(297, 137);
            this.cbxModuleId.Name = "cbxModuleId";
            this.cbxModuleId.Size = new System.Drawing.Size(15, 14);
            this.cbxModuleId.TabIndex = 50;
            this.cbxModuleId.UseVisualStyleBackColor = true;
            // 
            // txtShift
            // 
            this.txtShift.Location = new System.Drawing.Point(445, 102);
            this.txtShift.Name = "txtShift";
            this.txtShift.Size = new System.Drawing.Size(110, 19);
            this.txtShift.TabIndex = 1;
            // 
            // btnPartiallyCancel
            // 
            this.btnPartiallyCancel.Enabled = false;
            this.btnPartiallyCancel.Location = new System.Drawing.Point(905, 55);
            this.btnPartiallyCancel.Name = "btnPartiallyCancel";
            this.btnPartiallyCancel.Size = new System.Drawing.Size(100, 23);
            this.btnPartiallyCancel.TabIndex = 6;
            this.btnPartiallyCancel.Text = "部分取消";
            this.btnPartiallyCancel.UseVisualStyleBackColor = true;
            this.btnPartiallyCancel.Visible = false;
            this.btnPartiallyCancel.Click += new System.EventHandler(this.btnPartiallyCancel_Click);
            // 
            // frmTray
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1056, 650);
            this.Controls.Add(this.cbxShift);
            this.Controls.Add(this.cbxLine);
            this.Controls.Add(this.cbxUpdateDept);
            this.Controls.Add(this.cbxRegisterDept);
            this.Controls.Add(this.cbxUpdateDateTo);
            this.Controls.Add(this.cbxRegisterDateTo);
            this.Controls.Add(this.cbxUpdateDateFrom);
            this.Controls.Add(this.cbxRegisterDateFrom);
            this.Controls.Add(this.cbxModuleId);
            this.Controls.Add(this.cbxLot);
            this.Controls.Add(this.cbxTrayId);
            this.Controls.Add(this.cbxMultiLot);
            this.Controls.Add(this.cbxHideCancel);
            this.Controls.Add(this.cmbUpdateDept);
            this.Controls.Add(this.cmbRegisterDept);
            this.Controls.Add(this.cmbLine);
            this.Controls.Add(this.dtpUpdateDateTo);
            this.Controls.Add(this.dtpUpdateDateFrom);
            this.Controls.Add(this.dtpRegisterDateTo);
            this.Controls.Add(this.dtpRegsterDateFrom);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.txtShift);
            this.Controls.Add(this.txtModuleId);
            this.Controls.Add(this.txtLot);
            this.Controls.Add(this.txtTrayId);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtLoginDept);
            this.Controls.Add(this.txtLoginName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.btnPartiallyCancel);
            this.Controls.Add(this.btnCancelMultiTray);
            this.Controls.Add(this.btnAddReturn);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnAddTray);
            this.Controls.Add(this.btnSearchTray);
            this.Controls.Add(this.dgvTray);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmTray";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "1-A. Tray";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmTray_FormClosed);
            this.Load += new System.EventHandler(this.frmTray_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTray)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dgvTray;
        private System.Windows.Forms.Button btnSearchTray;
        private System.Windows.Forms.Button btnAddTray;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtTrayId;
        private System.Windows.Forms.DateTimePicker dtpRegsterDateFrom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnAddReturn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbLine;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbRegisterDept;
        private System.Windows.Forms.CheckBox cbxHideCancel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtLoginDept;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.CheckBox cbxTrayId;
        private System.Windows.Forms.CheckBox cbxRegisterDateFrom;
        private System.Windows.Forms.CheckBox cbxRegisterDept;
        private System.Windows.Forms.CheckBox cbxLine;
        private System.Windows.Forms.CheckBox cbxShift;
        private System.Windows.Forms.DateTimePicker dtpRegisterDateTo;
        private System.Windows.Forms.CheckBox cbxRegisterDateTo;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.DateTimePicker dtpUpdateDateFrom;
        private System.Windows.Forms.DateTimePicker dtpUpdateDateTo;
        private System.Windows.Forms.ComboBox cmbUpdateDept;
        private System.Windows.Forms.CheckBox cbxUpdateDateFrom;
        private System.Windows.Forms.CheckBox cbxUpdateDateTo;
        private System.Windows.Forms.CheckBox cbxUpdateDept;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtLot;
        private System.Windows.Forms.CheckBox cbxLot;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.CheckBox cbxMultiLot;
        private System.Windows.Forms.Button btnCancelMultiTray;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtModuleId;
        private System.Windows.Forms.CheckBox cbxModuleId;
        private System.Windows.Forms.TextBox txtShift;
        private System.Windows.Forms.Button btnPartiallyCancel;
    }
}

