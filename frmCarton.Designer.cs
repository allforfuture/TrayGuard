namespace TrayGuard
{
    partial class frmCarton
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCarton));
            this.btnAddCarton = new System.Windows.Forms.Button();
            this.btnSearchCarton = new System.Windows.Forms.Button();
            this.dgvCarton = new System.Windows.Forms.DataGridView();
            this.label12 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCartonId = new System.Windows.Forms.TextBox();
            this.dtpRegsterDateFrom = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbxHideCancel = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtLoginDept = new System.Windows.Forms.TextBox();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.cbxCartonId = new System.Windows.Forms.CheckBox();
            this.cbxRegisterDateFrom = new System.Windows.Forms.CheckBox();
            this.cbxPackId = new System.Windows.Forms.CheckBox();
            this.cbxBatch = new System.Windows.Forms.CheckBox();
            this.dtpRegisterDateTo = new System.Windows.Forms.DateTimePicker();
            this.cbxRegisterDateTo = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.txtLot = new System.Windows.Forms.TextBox();
            this.cbxLot = new System.Windows.Forms.CheckBox();
            this.label14 = new System.Windows.Forms.Label();
            this.cbxMultiLot = new System.Windows.Forms.CheckBox();
            this.txtPackId = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtPalletId = new System.Windows.Forms.TextBox();
            this.cbxPalletId = new System.Windows.Forms.CheckBox();
            this.txtBatch = new System.Windows.Forms.TextBox();
            this.btnPrintCarton = new System.Windows.Forms.Button();
            this.btnDeepCancelCartonPackTray = new System.Windows.Forms.Button();
            this.btnImportForDeepCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCarton)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAddCarton
            // 
            this.btnAddCarton.Location = new System.Drawing.Point(524, 119);
            this.btnAddCarton.Name = "btnAddCarton";
            this.btnAddCarton.Size = new System.Drawing.Size(100, 23);
            this.btnAddCarton.TabIndex = 6;
            this.btnAddCarton.Text = "新增纸箱";
            this.btnAddCarton.UseVisualStyleBackColor = true;
            this.btnAddCarton.Click += new System.EventHandler(this.btnAddBoxId_Click);
            // 
            // btnSearchCarton
            // 
            this.btnSearchCarton.Location = new System.Drawing.Point(383, 119);
            this.btnSearchCarton.Name = "btnSearchCarton";
            this.btnSearchCarton.Size = new System.Drawing.Size(100, 23);
            this.btnSearchCarton.TabIndex = 2;
            this.btnSearchCarton.Text = "搜索纸箱";
            this.btnSearchCarton.UseVisualStyleBackColor = true;
            this.btnSearchCarton.Click += new System.EventHandler(this.btnSearchPack_Click);
            // 
            // dgvCarton
            // 
            this.dgvCarton.AllowUserToAddRows = false;
            this.dgvCarton.BackgroundColor = System.Drawing.Color.MediumSpringGreen;
            this.dgvCarton.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCarton.Location = new System.Drawing.Point(12, 161);
            this.dgvCarton.Name = "dgvCarton";
            this.dgvCarton.ReadOnly = true;
            this.dgvCarton.RowTemplate.Height = 21;
            this.dgvCarton.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvCarton.Size = new System.Drawing.Size(971, 479);
            this.dgvCarton.TabIndex = 9;
            this.dgvCarton.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBoxId_CellContentClick);
            this.dgvCarton.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvCarton_CellContentDoubleClick);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(76, 15);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(91, 12);
            this.label12.TabIndex = 6;
            this.label12.Text = "登记日期(起始): ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(78, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "纸箱ID: ";
            // 
            // txtCartonId
            // 
            this.txtCartonId.Location = new System.Drawing.Point(181, 48);
            this.txtCartonId.Name = "txtCartonId";
            this.txtCartonId.Size = new System.Drawing.Size(110, 19);
            this.txtCartonId.TabIndex = 1;
            // 
            // dtpRegsterDateFrom
            // 
            this.dtpRegsterDateFrom.CustomFormat = "yyyy/MM/dd";
            this.dtpRegsterDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegsterDateFrom.Location = new System.Drawing.Point(181, 12);
            this.dtpRegsterDateFrom.Name = "dtpRegsterDateFrom";
            this.dtpRegsterDateFrom.Size = new System.Drawing.Size(110, 19);
            this.dtpRegsterDateFrom.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(634, 81);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "登录名: ";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(819, 119);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 23);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(355, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "捆扎ID: ";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(356, 81);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "批(版本): ";
            // 
            // cbxHideCancel
            // 
            this.cbxHideCancel.AutoSize = true;
            this.cbxHideCancel.Location = new System.Drawing.Point(309, 117);
            this.cbxHideCancel.Name = "cbxHideCancel";
            this.cbxHideCancel.Size = new System.Drawing.Size(15, 14);
            this.cbxHideCancel.TabIndex = 50;
            this.cbxHideCancel.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(214, 119);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "隐藏取消: ";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(634, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 6;
            this.label5.Text = "登陆部门: ";
            // 
            // txtLoginDept
            // 
            this.txtLoginDept.Enabled = false;
            this.txtLoginDept.Location = new System.Drawing.Point(733, 44);
            this.txtLoginDept.Name = "txtLoginDept";
            this.txtLoginDept.Size = new System.Drawing.Size(110, 19);
            this.txtLoginDept.TabIndex = 1;
            // 
            // txtLoginName
            // 
            this.txtLoginName.Enabled = false;
            this.txtLoginName.Location = new System.Drawing.Point(733, 78);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.Size = new System.Drawing.Size(110, 19);
            this.txtLoginName.TabIndex = 1;
            // 
            // cbxCartonId
            // 
            this.cbxCartonId.AutoSize = true;
            this.cbxCartonId.Location = new System.Drawing.Point(309, 51);
            this.cbxCartonId.Name = "cbxCartonId";
            this.cbxCartonId.Size = new System.Drawing.Size(15, 14);
            this.cbxCartonId.TabIndex = 50;
            this.cbxCartonId.UseVisualStyleBackColor = true;
            // 
            // cbxRegisterDateFrom
            // 
            this.cbxRegisterDateFrom.AutoSize = true;
            this.cbxRegisterDateFrom.Checked = true;
            this.cbxRegisterDateFrom.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbxRegisterDateFrom.Location = new System.Drawing.Point(309, 15);
            this.cbxRegisterDateFrom.Name = "cbxRegisterDateFrom";
            this.cbxRegisterDateFrom.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDateFrom.TabIndex = 50;
            this.cbxRegisterDateFrom.UseVisualStyleBackColor = true;
            // 
            // cbxPackId
            // 
            this.cbxPackId.AutoSize = true;
            this.cbxPackId.Location = new System.Drawing.Point(582, 50);
            this.cbxPackId.Name = "cbxPackId";
            this.cbxPackId.Size = new System.Drawing.Size(15, 14);
            this.cbxPackId.TabIndex = 50;
            this.cbxPackId.UseVisualStyleBackColor = true;
            // 
            // cbxBatch
            // 
            this.cbxBatch.AutoSize = true;
            this.cbxBatch.Location = new System.Drawing.Point(582, 82);
            this.cbxBatch.Name = "cbxBatch";
            this.cbxBatch.Size = new System.Drawing.Size(15, 14);
            this.cbxBatch.TabIndex = 50;
            this.cbxBatch.UseVisualStyleBackColor = true;
            // 
            // dtpRegisterDateTo
            // 
            this.dtpRegisterDateTo.CustomFormat = "yyyy/MM/dd";
            this.dtpRegisterDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegisterDateTo.Location = new System.Drawing.Point(454, 12);
            this.dtpRegisterDateTo.Name = "dtpRegisterDateTo";
            this.dtpRegisterDateTo.Size = new System.Drawing.Size(110, 19);
            this.dtpRegisterDateTo.TabIndex = 10;
            // 
            // cbxRegisterDateTo
            // 
            this.cbxRegisterDateTo.AutoSize = true;
            this.cbxRegisterDateTo.Location = new System.Drawing.Point(582, 15);
            this.cbxRegisterDateTo.Name = "cbxRegisterDateTo";
            this.cbxRegisterDateTo.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDateTo.TabIndex = 50;
            this.cbxRegisterDateTo.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(355, 15);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(91, 12);
            this.label8.TabIndex = 6;
            this.label8.Text = "登记日期(截止): ";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(78, 81);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(35, 12);
            this.label13.TabIndex = 6;
            this.label13.Text = "批号: ";
            // 
            // txtLot
            // 
            this.txtLot.Location = new System.Drawing.Point(181, 79);
            this.txtLot.Name = "txtLot";
            this.txtLot.Size = new System.Drawing.Size(110, 19);
            this.txtLot.TabIndex = 1;
            // 
            // cbxLot
            // 
            this.cbxLot.AutoSize = true;
            this.cbxLot.Location = new System.Drawing.Point(309, 81);
            this.cbxLot.Name = "cbxLot";
            this.cbxLot.Size = new System.Drawing.Size(15, 14);
            this.cbxLot.TabIndex = 50;
            this.cbxLot.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(78, 119);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(59, 12);
            this.label14.TabIndex = 6;
            this.label14.Text = "多个批号: ";
            // 
            // cbxMultiLot
            // 
            this.cbxMultiLot.AutoSize = true;
            this.cbxMultiLot.Location = new System.Drawing.Point(173, 117);
            this.cbxMultiLot.Name = "cbxMultiLot";
            this.cbxMultiLot.Size = new System.Drawing.Size(15, 14);
            this.cbxMultiLot.TabIndex = 50;
            this.cbxMultiLot.UseVisualStyleBackColor = true;
            // 
            // txtPackId
            // 
            this.txtPackId.Location = new System.Drawing.Point(454, 47);
            this.txtPackId.Name = "txtPackId";
            this.txtPackId.Size = new System.Drawing.Size(110, 19);
            this.txtPackId.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(634, 15);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "卡板ID: ";
            // 
            // txtPalletId
            // 
            this.txtPalletId.Enabled = false;
            this.txtPalletId.Location = new System.Drawing.Point(733, 12);
            this.txtPalletId.Name = "txtPalletId";
            this.txtPalletId.Size = new System.Drawing.Size(110, 19);
            this.txtPalletId.TabIndex = 1;
            // 
            // cbxPalletId
            // 
            this.cbxPalletId.AutoSize = true;
            this.cbxPalletId.Enabled = false;
            this.cbxPalletId.Location = new System.Drawing.Point(861, 15);
            this.cbxPalletId.Name = "cbxPalletId";
            this.cbxPalletId.Size = new System.Drawing.Size(15, 14);
            this.cbxPalletId.TabIndex = 50;
            this.cbxPalletId.UseVisualStyleBackColor = true;
            // 
            // txtBatch
            // 
            this.txtBatch.Enabled = false;
            this.txtBatch.Location = new System.Drawing.Point(454, 79);
            this.txtBatch.Name = "txtBatch";
            this.txtBatch.Size = new System.Drawing.Size(110, 19);
            this.txtBatch.TabIndex = 1;
            // 
            // btnPrintCarton
            // 
            this.btnPrintCarton.Location = new System.Drawing.Point(676, 119);
            this.btnPrintCarton.Name = "btnPrintCarton";
            this.btnPrintCarton.Size = new System.Drawing.Size(100, 23);
            this.btnPrintCarton.TabIndex = 6;
            this.btnPrintCarton.Text = "打印纸箱标签";
            this.btnPrintCarton.UseVisualStyleBackColor = true;
            this.btnPrintCarton.Click += new System.EventHandler(this.btnPrintCarton_Click);
            // 
            // btnDeepCancelCartonPackTray
            // 
            this.btnDeepCancelCartonPackTray.Enabled = false;
            this.btnDeepCancelCartonPackTray.Location = new System.Drawing.Point(887, 61);
            this.btnDeepCancelCartonPackTray.Name = "btnDeepCancelCartonPackTray";
            this.btnDeepCancelCartonPackTray.Size = new System.Drawing.Size(77, 37);
            this.btnDeepCancelCartonPackTray.TabIndex = 51;
            this.btnDeepCancelCartonPackTray.Text = "深取消纸箱捆扎马达盆";
            this.btnDeepCancelCartonPackTray.UseVisualStyleBackColor = true;
            this.btnDeepCancelCartonPackTray.Visible = false;
            this.btnDeepCancelCartonPackTray.Click += new System.EventHandler(this.btnDeepCancelCartonPackTray_Click);
            // 
            // btnImportForDeepCancel
            // 
            this.btnImportForDeepCancel.Enabled = false;
            this.btnImportForDeepCancel.Location = new System.Drawing.Point(887, 33);
            this.btnImportForDeepCancel.Name = "btnImportForDeepCancel";
            this.btnImportForDeepCancel.Size = new System.Drawing.Size(77, 22);
            this.btnImportForDeepCancel.TabIndex = 51;
            this.btnImportForDeepCancel.Text = "导入列表";
            this.btnImportForDeepCancel.UseVisualStyleBackColor = true;
            this.btnImportForDeepCancel.Visible = false;
            this.btnImportForDeepCancel.Click += new System.EventHandler(this.btnImportForDeepCancel_Click);
            // 
            // frmCarton
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(995, 650);
            this.Controls.Add(this.btnImportForDeepCancel);
            this.Controls.Add(this.btnDeepCancelCartonPackTray);
            this.Controls.Add(this.cbxBatch);
            this.Controls.Add(this.cbxPalletId);
            this.Controls.Add(this.cbxPackId);
            this.Controls.Add(this.cbxRegisterDateTo);
            this.Controls.Add(this.cbxRegisterDateFrom);
            this.Controls.Add(this.cbxLot);
            this.Controls.Add(this.cbxCartonId);
            this.Controls.Add(this.cbxMultiLot);
            this.Controls.Add(this.cbxHideCancel);
            this.Controls.Add(this.dtpRegisterDateTo);
            this.Controls.Add(this.dtpRegsterDateFrom);
            this.Controls.Add(this.txtPalletId);
            this.Controls.Add(this.txtBatch);
            this.Controls.Add(this.txtPackId);
            this.Controls.Add(this.txtLot);
            this.Controls.Add(this.txtCartonId);
            this.Controls.Add(this.txtLoginDept);
            this.Controls.Add(this.txtLoginName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.btnPrintCarton);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnAddCarton);
            this.Controls.Add(this.btnSearchCarton);
            this.Controls.Add(this.dgvCarton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmCarton";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "3-A. Carton";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmCarton_FormClosed);
            this.Load += new System.EventHandler(this.frmCarton_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCarton)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dgvCarton;
        private System.Windows.Forms.Button btnSearchCarton;
        private System.Windows.Forms.Button btnAddCarton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCartonId;
        private System.Windows.Forms.DateTimePicker dtpRegsterDateFrom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox cbxHideCancel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtLoginDept;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.CheckBox cbxCartonId;
        private System.Windows.Forms.CheckBox cbxRegisterDateFrom;
        private System.Windows.Forms.CheckBox cbxPackId;
        private System.Windows.Forms.CheckBox cbxBatch;
        private System.Windows.Forms.DateTimePicker dtpRegisterDateTo;
        private System.Windows.Forms.CheckBox cbxRegisterDateTo;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtLot;
        private System.Windows.Forms.CheckBox cbxLot;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.CheckBox cbxMultiLot;
        private System.Windows.Forms.TextBox txtPackId;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtPalletId;
        private System.Windows.Forms.CheckBox cbxPalletId;
        private System.Windows.Forms.TextBox txtBatch;
        private System.Windows.Forms.Button btnPrintCarton;
        private System.Windows.Forms.Button btnDeepCancelCartonPackTray;
        private System.Windows.Forms.Button btnImportForDeepCancel;
    }
}

