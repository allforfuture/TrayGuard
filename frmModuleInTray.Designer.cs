

namespace TrayGuard
{
    partial class frmModuleInTray
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmModuleInTray));
            this.btnClose = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.dgvModule = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.txtTrayId = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpRegisterDate = new System.Windows.Forms.DateTimePicker();
            this.txtModuleId = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnRegisterTray = new System.Windows.Forms.Button();
            this.txtOkCount = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnDeleteSelection = new System.Windows.Forms.Button();
            this.btnChangeCapacity = new System.Windows.Forms.Button();
            this.dgvLotSummary = new System.Windows.Forms.DataGridView();
            this.btnReplaceModule = new System.Windows.Forms.Button();
            this.btnReprintLabel = new System.Windows.Forms.Button();
            this.btnCancelTray = new System.Windows.Forms.Button();
            this.txtLoginDept = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbBinShift = new System.Windows.Forms.ComboBox();
            this.lblBin = new System.Windows.Forms.Label();
            this.btnRefillTray = new System.Windows.Forms.Button();
            this.btnExportModule = new System.Windows.Forms.Button();
            this.SptReceiveOrSend = new System.IO.Ports.SerialPort(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dgvModule)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLotSummary)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(667, 161);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(106, 22);
            this.btnClose.TabIndex = 11;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(565, 128);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(53, 12);
            this.label12.TabIndex = 6;
            this.label12.Text = "登录名: ";
            // 
            // txtLoginName
            // 
            this.txtLoginName.Enabled = false;
            this.txtLoginName.Location = new System.Drawing.Point(638, 125);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.Size = new System.Drawing.Size(105, 21);
            this.txtLoginName.TabIndex = 5;
            // 
            // dgvModule
            // 
            this.dgvModule.AllowUserToAddRows = false;
            this.dgvModule.AllowUserToDeleteRows = false;
            this.dgvModule.BackgroundColor = System.Drawing.Color.SkyBlue;
            this.dgvModule.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvModule.Location = new System.Drawing.Point(12, 230);
            this.dgvModule.Name = "dgvModule";
            this.dgvModule.ReadOnly = true;
            this.dgvModule.RowTemplate.Height = 21;
            this.dgvModule.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvModule.Size = new System.Drawing.Size(764, 459);
            this.dgvModule.TabIndex = 9;
            this.dgvModule.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvOverall_CellContentDoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 131);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "登记日期: ";
            // 
            // txtTrayId
            // 
            this.txtTrayId.Enabled = false;
            this.txtTrayId.Location = new System.Drawing.Point(117, 97);
            this.txtTrayId.Name = "txtTrayId";
            this.txtTrayId.Size = new System.Drawing.Size(160, 21);
            this.txtTrayId.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "马达盆ID: ";
            // 
            // dtpRegisterDate
            // 
            this.dtpRegisterDate.CustomFormat = "yyyy/MM/dd  HH:mm:ss";
            this.dtpRegisterDate.Enabled = false;
            this.dtpRegisterDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegisterDate.Location = new System.Drawing.Point(117, 128);
            this.dtpRegisterDate.Name = "dtpRegisterDate";
            this.dtpRegisterDate.Size = new System.Drawing.Size(160, 21);
            this.dtpRegisterDate.TabIndex = 12;
            // 
            // txtModuleId
            // 
            this.txtModuleId.Enabled = false;
            this.txtModuleId.Location = new System.Drawing.Point(117, 180);
            this.txtModuleId.Name = "txtModuleId";
            this.txtModuleId.Size = new System.Drawing.Size(160, 21);
            this.txtModuleId.TabIndex = 5;
            this.txtModuleId.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtModuleId_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 185);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "马达ID: ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "批号汇总:";
            // 
            // btnRegisterTray
            // 
            this.btnRegisterTray.Enabled = false;
            this.btnRegisterTray.Location = new System.Drawing.Point(292, 161);
            this.btnRegisterTray.Name = "btnRegisterTray";
            this.btnRegisterTray.Size = new System.Drawing.Size(106, 22);
            this.btnRegisterTray.TabIndex = 11;
            this.btnRegisterTray.Text = "登记马达盆";
            this.btnRegisterTray.UseVisualStyleBackColor = true;
            this.btnRegisterTray.Click += new System.EventHandler(this.btnRegisterTray_Click);
            // 
            // txtOkCount
            // 
            this.txtOkCount.Location = new System.Drawing.Point(403, 125);
            this.txtOkCount.Name = "txtOkCount";
            this.txtOkCount.ReadOnly = true;
            this.txtOkCount.Size = new System.Drawing.Size(105, 21);
            this.txtOkCount.TabIndex = 5;
            this.txtOkCount.DoubleClick += new System.EventHandler(this.txtOkCount_DoubleClick);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(330, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "OK 数量: ";
            // 
            // btnDeleteSelection
            // 
            this.btnDeleteSelection.Enabled = false;
            this.btnDeleteSelection.Location = new System.Drawing.Point(419, 161);
            this.btnDeleteSelection.Name = "btnDeleteSelection";
            this.btnDeleteSelection.Size = new System.Drawing.Size(106, 22);
            this.btnDeleteSelection.TabIndex = 11;
            this.btnDeleteSelection.Text = "删除";
            this.btnDeleteSelection.UseVisualStyleBackColor = true;
            this.btnDeleteSelection.Click += new System.EventHandler(this.btnDeleteSelection_Click);
            // 
            // btnChangeCapacity
            // 
            this.btnChangeCapacity.Enabled = false;
            this.btnChangeCapacity.Location = new System.Drawing.Point(545, 161);
            this.btnChangeCapacity.Name = "btnChangeCapacity";
            this.btnChangeCapacity.Size = new System.Drawing.Size(106, 22);
            this.btnChangeCapacity.TabIndex = 11;
            this.btnChangeCapacity.Text = "变更容量";
            this.btnChangeCapacity.UseVisualStyleBackColor = true;
            this.btnChangeCapacity.Click += new System.EventHandler(this.btnChangecapacity_Click);
            // 
            // dgvLotSummary
            // 
            this.dgvLotSummary.AllowUserToAddRows = false;
            this.dgvLotSummary.AllowUserToDeleteRows = false;
            this.dgvLotSummary.BackgroundColor = System.Drawing.Color.SkyBlue;
            this.dgvLotSummary.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dgvLotSummary.GridColor = System.Drawing.Color.Azure;
            this.dgvLotSummary.Location = new System.Drawing.Point(117, 12);
            this.dgvLotSummary.Name = "dgvLotSummary";
            this.dgvLotSummary.ReadOnly = true;
            this.dgvLotSummary.RowTemplate.Height = 21;
            this.dgvLotSummary.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvLotSummary.Size = new System.Drawing.Size(636, 61);
            this.dgvLotSummary.TabIndex = 9;
            // 
            // btnReplaceModule
            // 
            this.btnReplaceModule.Enabled = false;
            this.btnReplaceModule.Location = new System.Drawing.Point(545, 193);
            this.btnReplaceModule.Name = "btnReplaceModule";
            this.btnReplaceModule.Size = new System.Drawing.Size(106, 22);
            this.btnReplaceModule.TabIndex = 14;
            this.btnReplaceModule.Text = "替换马达";
            this.btnReplaceModule.UseVisualStyleBackColor = true;
            this.btnReplaceModule.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // btnReprintLabel
            // 
            this.btnReprintLabel.Enabled = false;
            this.btnReprintLabel.Location = new System.Drawing.Point(292, 193);
            this.btnReprintLabel.Name = "btnReprintLabel";
            this.btnReprintLabel.Size = new System.Drawing.Size(106, 22);
            this.btnReprintLabel.TabIndex = 42;
            this.btnReprintLabel.Text = "再次打印标签";
            this.btnReprintLabel.UseVisualStyleBackColor = true;
            this.btnReprintLabel.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancelTray
            // 
            this.btnCancelTray.Enabled = false;
            this.btnCancelTray.Location = new System.Drawing.Point(667, 193);
            this.btnCancelTray.Name = "btnCancelTray";
            this.btnCancelTray.Size = new System.Drawing.Size(106, 22);
            this.btnCancelTray.TabIndex = 43;
            this.btnCancelTray.Text = "取消马达盆";
            this.btnCancelTray.UseVisualStyleBackColor = true;
            this.btnCancelTray.Click += new System.EventHandler(this.btnCancelTray_Click);
            // 
            // txtLoginDept
            // 
            this.txtLoginDept.Enabled = false;
            this.txtLoginDept.Location = new System.Drawing.Point(637, 97);
            this.txtLoginDept.Name = "txtLoginDept";
            this.txtLoginDept.Size = new System.Drawing.Size(105, 21);
            this.txtLoginDept.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(564, 100);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "登陆部门: ";
            // 
            // cmbBinShift
            // 
            this.cmbBinShift.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBinShift.Enabled = false;
            this.cmbBinShift.FormattingEnabled = true;
            this.cmbBinShift.Items.AddRange(new object[] {
            "1",
            "2",
            "X"});
            this.cmbBinShift.Location = new System.Drawing.Point(403, 96);
            this.cmbBinShift.Name = "cmbBinShift";
            this.cmbBinShift.Size = new System.Drawing.Size(105, 20);
            this.cmbBinShift.TabIndex = 44;
            // 
            // lblBin
            // 
            this.lblBin.AutoSize = true;
            this.lblBin.Location = new System.Drawing.Point(327, 100);
            this.lblBin.Name = "lblBin";
            this.lblBin.Size = new System.Drawing.Size(41, 12);
            this.lblBin.TabIndex = 6;
            this.lblBin.Text = "班次: ";
            // 
            // btnRefillTray
            // 
            this.btnRefillTray.Enabled = false;
            this.btnRefillTray.Location = new System.Drawing.Point(419, 193);
            this.btnRefillTray.Name = "btnRefillTray";
            this.btnRefillTray.Size = new System.Drawing.Size(106, 22);
            this.btnRefillTray.TabIndex = 14;
            this.btnRefillTray.Text = "增加马达";
            this.btnRefillTray.UseVisualStyleBackColor = true;
            this.btnRefillTray.Click += new System.EventHandler(this.btnRefillTray_Click);
            // 
            // btnExportModule
            // 
            this.btnExportModule.Location = new System.Drawing.Point(16, 42);
            this.btnExportModule.Name = "btnExportModule";
            this.btnExportModule.Size = new System.Drawing.Size(70, 22);
            this.btnExportModule.TabIndex = 47;
            this.btnExportModule.Text = "输出马达";
            this.btnExportModule.UseVisualStyleBackColor = true;
            this.btnExportModule.Click += new System.EventHandler(this.btnExportModule_Click);
            // 
            // SptReceiveOrSend
            // 
            this.SptReceiveOrSend.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.SptReceiveOrSend_DataReceived);
            // 
            // frmModuleInTray
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(788, 701);
            this.Controls.Add(this.btnExportModule);
            this.Controls.Add(this.cmbBinShift);
            this.Controls.Add(this.dgvModule);
            this.Controls.Add(this.btnCancelTray);
            this.Controls.Add(this.btnReprintLabel);
            this.Controls.Add(this.btnRefillTray);
            this.Controls.Add(this.btnReplaceModule);
            this.Controls.Add(this.dgvLotSummary);
            this.Controls.Add(this.dtpRegisterDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.lblBin);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.txtModuleId);
            this.Controls.Add(this.txtTrayId);
            this.Controls.Add(this.txtOkCount);
            this.Controls.Add(this.txtLoginDept);
            this.Controls.Add(this.txtLoginName);
            this.Controls.Add(this.btnRegisterTray);
            this.Controls.Add(this.btnDeleteSelection);
            this.Controls.Add(this.btnChangeCapacity);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmModuleInTray";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "1-B. Module in Tray";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmModuleInTray_FormClosing);
            this.Load += new System.EventHandler(this.frmModule_Load);
            this.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.frmModuleInTray_MouseDoubleClick);
            ((System.ComponentModel.ISupportInitialize)(this.dgvModule)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLotSummary)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.DataGridView dgvModule;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtTrayId;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpRegisterDate;
        private System.Windows.Forms.TextBox txtModuleId;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnRegisterTray;
        private System.Windows.Forms.TextBox txtOkCount;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnDeleteSelection;
        private System.Windows.Forms.Button btnChangeCapacity;
        private System.Windows.Forms.DataGridView dgvLotSummary;
        private System.Windows.Forms.Button btnReplaceModule;
        private System.Windows.Forms.Button btnReprintLabel;
        private System.Windows.Forms.Button btnCancelTray;
        private System.Windows.Forms.TextBox txtLoginDept;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbBinShift;
        private System.Windows.Forms.Label lblBin;
        private System.Windows.Forms.Button btnRefillTray;
        private System.Windows.Forms.Button btnExportModule;
        private System.IO.Ports.SerialPort SptReceiveOrSend;
    }
}

