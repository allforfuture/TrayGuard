namespace TrayGuard
{
    partial class frmPack
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPack));
            this.btnAddTray = new System.Windows.Forms.Button();
            this.btnSearchPack = new System.Windows.Forms.Button();
            this.dgvPack = new System.Windows.Forms.DataGridView();
            this.label12 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPackId = new System.Windows.Forms.TextBox();
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
            this.cbxPackId = new System.Windows.Forms.CheckBox();
            this.cbxRegisterDateFrom = new System.Windows.Forms.CheckBox();
            this.cbxTrayId = new System.Windows.Forms.CheckBox();
            this.cbxBatch = new System.Windows.Forms.CheckBox();
            this.dtpRegisterDateTo = new System.Windows.Forms.DateTimePicker();
            this.cbxRegisterDateTo = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.txtLot = new System.Windows.Forms.TextBox();
            this.cbxLot = new System.Windows.Forms.CheckBox();
            this.label14 = new System.Windows.Forms.Label();
            this.cbxMultiLot = new System.Windows.Forms.CheckBox();
            this.txtTrayId = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtCarton = new System.Windows.Forms.TextBox();
            this.cbxCarton = new System.Windows.Forms.CheckBox();
            this.btnPrint = new System.Windows.Forms.Button();
            this.btnSetUpLabel = new System.Windows.Forms.Button();
            this.txtBatch = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPack)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAddTray
            // 
            this.btnAddTray.Location = new System.Drawing.Point(493, 119);
            this.btnAddTray.Name = "btnAddTray";
            this.btnAddTray.Size = new System.Drawing.Size(100, 23);
            this.btnAddTray.TabIndex = 6;
            this.btnAddTray.Text = "新增捆扎";
            this.btnAddTray.UseVisualStyleBackColor = true;
            this.btnAddTray.Click += new System.EventHandler(this.btnAddBoxId_Click);
            // 
            // btnSearchPack
            // 
            this.btnSearchPack.Location = new System.Drawing.Point(372, 119);
            this.btnSearchPack.Name = "btnSearchPack";
            this.btnSearchPack.Size = new System.Drawing.Size(100, 23);
            this.btnSearchPack.TabIndex = 2;
            this.btnSearchPack.Text = "搜索捆扎";
            this.btnSearchPack.UseVisualStyleBackColor = true;
            this.btnSearchPack.Click += new System.EventHandler(this.btnSearchPack_Click);
            // 
            // dgvPack
            // 
            this.dgvPack.AllowUserToAddRows = false;
            this.dgvPack.BackgroundColor = System.Drawing.Color.Orange;
            this.dgvPack.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPack.Location = new System.Drawing.Point(12, 161);
            this.dgvPack.Name = "dgvPack";
            this.dgvPack.ReadOnly = true;
            this.dgvPack.RowTemplate.Height = 21;
            this.dgvPack.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvPack.Size = new System.Drawing.Size(971, 479);
            this.dgvPack.TabIndex = 9;
            this.dgvPack.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBoxId_CellContentClick);
            this.dgvPack.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPack_CellContentDoubleClick);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(78, 15);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(91, 12);
            this.label12.TabIndex = 6;
            this.label12.Text = "登记日期(起始): ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(80, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "捆扎ID: ";
            // 
            // txtPackId
            // 
            this.txtPackId.Location = new System.Drawing.Point(183, 46);
            this.txtPackId.Name = "txtPackId";
            this.txtPackId.Size = new System.Drawing.Size(110, 19);
            this.txtPackId.TabIndex = 1;
            // 
            // dtpRegsterDateFrom
            // 
            this.dtpRegsterDateFrom.CustomFormat = "yyyy/MM/dd";
            this.dtpRegsterDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegsterDateFrom.Location = new System.Drawing.Point(183, 12);
            this.dtpRegsterDateFrom.Name = "dtpRegsterDateFrom";
            this.dtpRegsterDateFrom.Size = new System.Drawing.Size(110, 19);
            this.dtpRegsterDateFrom.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(635, 81);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "登陆名: ";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(858, 119);
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
            this.label2.Location = new System.Drawing.Point(357, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "马达盆ID: ";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(358, 81);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "批(版本): ";
            // 
            // cbxHideCancel
            // 
            this.cbxHideCancel.AutoSize = true;
            this.cbxHideCancel.Location = new System.Drawing.Point(311, 117);
            this.cbxHideCancel.Name = "cbxHideCancel";
            this.cbxHideCancel.Size = new System.Drawing.Size(15, 14);
            this.cbxHideCancel.TabIndex = 50;
            this.cbxHideCancel.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(216, 119);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "隐藏取消: ";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(635, 50);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 12);
            this.label5.TabIndex = 6;
            this.label5.Text = "登陆部门: ";
            // 
            // txtLoginDept
            // 
            this.txtLoginDept.Enabled = false;
            this.txtLoginDept.Location = new System.Drawing.Point(734, 45);
            this.txtLoginDept.Name = "txtLoginDept";
            this.txtLoginDept.Size = new System.Drawing.Size(110, 19);
            this.txtLoginDept.TabIndex = 1;
            // 
            // txtLoginName
            // 
            this.txtLoginName.Enabled = false;
            this.txtLoginName.Location = new System.Drawing.Point(734, 78);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.Size = new System.Drawing.Size(110, 19);
            this.txtLoginName.TabIndex = 1;
            // 
            // cbxPackId
            // 
            this.cbxPackId.AutoSize = true;
            this.cbxPackId.Location = new System.Drawing.Point(311, 49);
            this.cbxPackId.Name = "cbxPackId";
            this.cbxPackId.Size = new System.Drawing.Size(15, 14);
            this.cbxPackId.TabIndex = 50;
            this.cbxPackId.UseVisualStyleBackColor = true;
            // 
            // cbxRegisterDateFrom
            // 
            this.cbxRegisterDateFrom.AutoSize = true;
            this.cbxRegisterDateFrom.Checked = true;
            this.cbxRegisterDateFrom.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbxRegisterDateFrom.Location = new System.Drawing.Point(311, 15);
            this.cbxRegisterDateFrom.Name = "cbxRegisterDateFrom";
            this.cbxRegisterDateFrom.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDateFrom.TabIndex = 50;
            this.cbxRegisterDateFrom.UseVisualStyleBackColor = true;
            // 
            // cbxTrayId
            // 
            this.cbxTrayId.AutoSize = true;
            this.cbxTrayId.Location = new System.Drawing.Point(584, 48);
            this.cbxTrayId.Name = "cbxTrayId";
            this.cbxTrayId.Size = new System.Drawing.Size(15, 14);
            this.cbxTrayId.TabIndex = 50;
            this.cbxTrayId.UseVisualStyleBackColor = true;
            // 
            // cbxBatch
            // 
            this.cbxBatch.AutoSize = true;
            this.cbxBatch.Location = new System.Drawing.Point(584, 82);
            this.cbxBatch.Name = "cbxBatch";
            this.cbxBatch.Size = new System.Drawing.Size(15, 14);
            this.cbxBatch.TabIndex = 50;
            this.cbxBatch.UseVisualStyleBackColor = true;
            // 
            // dtpRegisterDateTo
            // 
            this.dtpRegisterDateTo.CustomFormat = "yyyy/MM/dd";
            this.dtpRegisterDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegisterDateTo.Location = new System.Drawing.Point(456, 12);
            this.dtpRegisterDateTo.Name = "dtpRegisterDateTo";
            this.dtpRegisterDateTo.Size = new System.Drawing.Size(110, 19);
            this.dtpRegisterDateTo.TabIndex = 10;
            // 
            // cbxRegisterDateTo
            // 
            this.cbxRegisterDateTo.AutoSize = true;
            this.cbxRegisterDateTo.Location = new System.Drawing.Point(584, 15);
            this.cbxRegisterDateTo.Name = "cbxRegisterDateTo";
            this.cbxRegisterDateTo.Size = new System.Drawing.Size(15, 14);
            this.cbxRegisterDateTo.TabIndex = 50;
            this.cbxRegisterDateTo.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(357, 15);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(91, 12);
            this.label8.TabIndex = 6;
            this.label8.Text = "登记日期(截止): ";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(80, 81);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(35, 12);
            this.label13.TabIndex = 6;
            this.label13.Text = "批次: ";
            // 
            // txtLot
            // 
            this.txtLot.Location = new System.Drawing.Point(183, 79);
            this.txtLot.Name = "txtLot";
            this.txtLot.Size = new System.Drawing.Size(110, 19);
            this.txtLot.TabIndex = 1;
            // 
            // cbxLot
            // 
            this.cbxLot.AutoSize = true;
            this.cbxLot.Location = new System.Drawing.Point(311, 81);
            this.cbxLot.Name = "cbxLot";
            this.cbxLot.Size = new System.Drawing.Size(15, 14);
            this.cbxLot.TabIndex = 50;
            this.cbxLot.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(80, 119);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(59, 12);
            this.label14.TabIndex = 6;
            this.label14.Text = "多个批号: ";
            // 
            // cbxMultiLot
            // 
            this.cbxMultiLot.AutoSize = true;
            this.cbxMultiLot.Location = new System.Drawing.Point(175, 117);
            this.cbxMultiLot.Name = "cbxMultiLot";
            this.cbxMultiLot.Size = new System.Drawing.Size(15, 14);
            this.cbxMultiLot.TabIndex = 50;
            this.cbxMultiLot.UseVisualStyleBackColor = true;
            // 
            // txtTrayId
            // 
            this.txtTrayId.Location = new System.Drawing.Point(456, 45);
            this.txtTrayId.Name = "txtTrayId";
            this.txtTrayId.Size = new System.Drawing.Size(110, 19);
            this.txtTrayId.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(635, 15);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "纸箱ID: ";
            // 
            // txtCarton
            // 
            this.txtCarton.Enabled = false;
            this.txtCarton.Location = new System.Drawing.Point(734, 12);
            this.txtCarton.Name = "txtCarton";
            this.txtCarton.Size = new System.Drawing.Size(110, 19);
            this.txtCarton.TabIndex = 1;
            // 
            // cbxCarton
            // 
            this.cbxCarton.AutoSize = true;
            this.cbxCarton.Enabled = false;
            this.cbxCarton.Location = new System.Drawing.Point(862, 15);
            this.cbxCarton.Name = "cbxCarton";
            this.cbxCarton.Size = new System.Drawing.Size(15, 14);
            this.cbxCarton.TabIndex = 50;
            this.cbxCarton.UseVisualStyleBackColor = true;
            // 
            // btnPrint
            // 
            this.btnPrint.Location = new System.Drawing.Point(738, 119);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(100, 23);
            this.btnPrint.TabIndex = 6;
            this.btnPrint.Text = "打印捆扎标签";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnSetUpLabel
            // 
            this.btnSetUpLabel.Enabled = false;
            this.btnSetUpLabel.Location = new System.Drawing.Point(615, 119);
            this.btnSetUpLabel.Name = "btnSetUpLabel";
            this.btnSetUpLabel.Size = new System.Drawing.Size(100, 23);
            this.btnSetUpLabel.TabIndex = 6;
            this.btnSetUpLabel.Text = "建立标签";
            this.btnSetUpLabel.UseVisualStyleBackColor = true;
            this.btnSetUpLabel.Click += new System.EventHandler(this.btnSetUpLabel_Click);
            // 
            // txtBatch
            // 
            this.txtBatch.Enabled = false;
            this.txtBatch.Location = new System.Drawing.Point(456, 79);
            this.txtBatch.Name = "txtBatch";
            this.txtBatch.Size = new System.Drawing.Size(110, 19);
            this.txtBatch.TabIndex = 1;
            // 
            // frmPack
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(995, 650);
            this.Controls.Add(this.cbxBatch);
            this.Controls.Add(this.cbxCarton);
            this.Controls.Add(this.cbxTrayId);
            this.Controls.Add(this.cbxRegisterDateTo);
            this.Controls.Add(this.cbxRegisterDateFrom);
            this.Controls.Add(this.cbxLot);
            this.Controls.Add(this.cbxPackId);
            this.Controls.Add(this.cbxMultiLot);
            this.Controls.Add(this.cbxHideCancel);
            this.Controls.Add(this.dtpRegisterDateTo);
            this.Controls.Add(this.dtpRegsterDateFrom);
            this.Controls.Add(this.txtCarton);
            this.Controls.Add(this.txtBatch);
            this.Controls.Add(this.txtTrayId);
            this.Controls.Add(this.txtLot);
            this.Controls.Add(this.txtPackId);
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
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSetUpLabel);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnAddTray);
            this.Controls.Add(this.btnSearchPack);
            this.Controls.Add(this.dgvPack);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmPack";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "2-A. Pack";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmPack_FormClosed);
            this.Load += new System.EventHandler(this.frmPack_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPack)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dgvPack;
        private System.Windows.Forms.Button btnSearchPack;
        private System.Windows.Forms.Button btnAddTray;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPackId;
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
        private System.Windows.Forms.CheckBox cbxPackId;
        private System.Windows.Forms.CheckBox cbxRegisterDateFrom;
        private System.Windows.Forms.CheckBox cbxTrayId;
        private System.Windows.Forms.CheckBox cbxBatch;
        private System.Windows.Forms.DateTimePicker dtpRegisterDateTo;
        private System.Windows.Forms.CheckBox cbxRegisterDateTo;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtLot;
        private System.Windows.Forms.CheckBox cbxLot;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.CheckBox cbxMultiLot;
        private System.Windows.Forms.TextBox txtTrayId;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtCarton;
        private System.Windows.Forms.CheckBox cbxCarton;
        private System.Windows.Forms.Button btnPrint;
        private System.Windows.Forms.Button btnSetUpLabel;
        private System.Windows.Forms.TextBox txtBatch;
    }
}

