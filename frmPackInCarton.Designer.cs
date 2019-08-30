

namespace TrayGuard
{
    partial class frmPackInCarton
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPackInCarton));
            this.btnClose = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCarton = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpRegisterDate = new System.Windows.Forms.DateTimePicker();
            this.txtPack = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnRegisterCarton = new System.Windows.Forms.Button();
            this.txtOkCount = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnDeleteSelection = new System.Windows.Forms.Button();
            this.dgvLotSummary = new System.Windows.Forms.DataGridView();
            this.btnCancelCarton = new System.Windows.Forms.Button();
            this.txtLoginDept = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtBatch = new System.Windows.Forms.TextBox();
            this.dgvPack = new System.Windows.Forms.DataGridView();
            this.btnExportModule = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLotSummary)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPack)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(666, 175);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 22);
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
            this.label12.Size = new System.Drawing.Size(47, 12);
            this.label12.TabIndex = 6;
            this.label12.Text = "登录名: ";
            // 
            // txtLoginName
            // 
            this.txtLoginName.Enabled = false;
            this.txtLoginName.Location = new System.Drawing.Point(638, 125);
            this.txtLoginName.Name = "txtLoginName";
            this.txtLoginName.Size = new System.Drawing.Size(105, 19);
            this.txtLoginName.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 131);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "登记日期: ";
            // 
            // txtCarton
            // 
            this.txtCarton.Enabled = false;
            this.txtCarton.Location = new System.Drawing.Point(117, 97);
            this.txtCarton.Name = "txtCarton";
            this.txtCarton.Size = new System.Drawing.Size(160, 19);
            this.txtCarton.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "纸箱ID: ";
            // 
            // dtpRegisterDate
            // 
            this.dtpRegisterDate.CustomFormat = "yyyy/MM/dd  HH:mm:ss";
            this.dtpRegisterDate.Enabled = false;
            this.dtpRegisterDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegisterDate.Location = new System.Drawing.Point(117, 128);
            this.dtpRegisterDate.Name = "dtpRegisterDate";
            this.dtpRegisterDate.Size = new System.Drawing.Size(160, 19);
            this.dtpRegisterDate.TabIndex = 12;
            // 
            // txtPack
            // 
            this.txtPack.Enabled = false;
            this.txtPack.Location = new System.Drawing.Point(117, 177);
            this.txtPack.Name = "txtPack";
            this.txtPack.Size = new System.Drawing.Size(160, 19);
            this.txtPack.TabIndex = 5;
            this.txtPack.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtModuleId_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 182);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "捆扎ID: ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 12);
            this.label6.TabIndex = 6;
            this.label6.Text = "批号汇总";
            // 
            // btnRegisterCarton
            // 
            this.btnRegisterCarton.Enabled = false;
            this.btnRegisterCarton.Location = new System.Drawing.Point(295, 175);
            this.btnRegisterCarton.Name = "btnRegisterCarton";
            this.btnRegisterCarton.Size = new System.Drawing.Size(100, 22);
            this.btnRegisterCarton.TabIndex = 11;
            this.btnRegisterCarton.Text = "登记纸箱";
            this.btnRegisterCarton.UseVisualStyleBackColor = true;
            this.btnRegisterCarton.Click += new System.EventHandler(this.btnRegisterTray_Click);
            // 
            // txtOkCount
            // 
            this.txtOkCount.Location = new System.Drawing.Point(403, 125);
            this.txtOkCount.Name = "txtOkCount";
            this.txtOkCount.ReadOnly = true;
            this.txtOkCount.Size = new System.Drawing.Size(105, 19);
            this.txtOkCount.TabIndex = 5;
            this.txtOkCount.DoubleClick += new System.EventHandler(this.txtOkCount_DoubleClick);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(330, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(54, 12);
            this.label4.TabIndex = 6;
            this.label4.Text = "OK 数量: ";
            // 
            // btnDeleteSelection
            // 
            this.btnDeleteSelection.Enabled = false;
            this.btnDeleteSelection.Location = new System.Drawing.Point(417, 175);
            this.btnDeleteSelection.Name = "btnDeleteSelection";
            this.btnDeleteSelection.Size = new System.Drawing.Size(100, 22);
            this.btnDeleteSelection.TabIndex = 11;
            this.btnDeleteSelection.Text = "删除";
            this.btnDeleteSelection.UseVisualStyleBackColor = true;
            this.btnDeleteSelection.Click += new System.EventHandler(this.btnDeleteSelection_Click);
            // 
            // dgvLotSummary
            // 
            this.dgvLotSummary.AllowUserToAddRows = false;
            this.dgvLotSummary.AllowUserToDeleteRows = false;
            this.dgvLotSummary.BackgroundColor = System.Drawing.Color.DarkOrange;
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
            // btnCancelCarton
            // 
            this.btnCancelCarton.Enabled = false;
            this.btnCancelCarton.Location = new System.Drawing.Point(541, 175);
            this.btnCancelCarton.Name = "btnCancelCarton";
            this.btnCancelCarton.Size = new System.Drawing.Size(100, 22);
            this.btnCancelCarton.TabIndex = 43;
            this.btnCancelCarton.Text = "取消纸箱";
            this.btnCancelCarton.UseVisualStyleBackColor = true;
            this.btnCancelCarton.Click += new System.EventHandler(this.btnCancelCarton_Click);
            // 
            // txtLoginDept
            // 
            this.txtLoginDept.Enabled = false;
            this.txtLoginDept.Location = new System.Drawing.Point(637, 97);
            this.txtLoginDept.Name = "txtLoginDept";
            this.txtLoginDept.Size = new System.Drawing.Size(105, 19);
            this.txtLoginDept.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(564, 100);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "登陆部门: ";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(327, 100);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(55, 12);
            this.label8.TabIndex = 6;
            this.label8.Text = "批(版本): ";
            // 
            // txtBatch
            // 
            this.txtBatch.Location = new System.Drawing.Point(403, 97);
            this.txtBatch.Name = "txtBatch";
            this.txtBatch.ReadOnly = true;
            this.txtBatch.Size = new System.Drawing.Size(105, 19);
            this.txtBatch.TabIndex = 5;
            // 
            // dgvPack
            // 
            this.dgvPack.AllowUserToAddRows = false;
            this.dgvPack.AllowUserToDeleteRows = false;
            this.dgvPack.BackgroundColor = System.Drawing.Color.DarkOrange;
            this.dgvPack.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPack.Location = new System.Drawing.Point(12, 215);
            this.dgvPack.Name = "dgvPack";
            this.dgvPack.ReadOnly = true;
            this.dgvPack.RowTemplate.Height = 21;
            this.dgvPack.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvPack.Size = new System.Drawing.Size(764, 271);
            this.dgvPack.TabIndex = 9;
            this.dgvPack.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPack_CellContentClick);
            this.dgvPack.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPack_CellContentDoubleClick);
            // 
            // btnExportModule
            // 
            this.btnExportModule.Location = new System.Drawing.Point(16, 42);
            this.btnExportModule.Name = "btnExportModule";
            this.btnExportModule.Size = new System.Drawing.Size(70, 22);
            this.btnExportModule.TabIndex = 46;
            this.btnExportModule.Text = "输出马达";
            this.btnExportModule.UseVisualStyleBackColor = true;
            this.btnExportModule.Click += new System.EventHandler(this.btnExportModule_Click);
            // 
            // frmPackInCarton
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(788, 498);
            this.Controls.Add(this.btnExportModule);
            this.Controls.Add(this.dgvPack);
            this.Controls.Add(this.btnCancelCarton);
            this.Controls.Add(this.dgvLotSummary);
            this.Controls.Add(this.dtpRegisterDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.txtPack);
            this.Controls.Add(this.txtCarton);
            this.Controls.Add(this.txtBatch);
            this.Controls.Add(this.txtOkCount);
            this.Controls.Add(this.txtLoginDept);
            this.Controls.Add(this.txtLoginName);
            this.Controls.Add(this.btnRegisterCarton);
            this.Controls.Add(this.btnDeleteSelection);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmPackInCarton";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "3-B. Pack in Carton";
            this.Load += new System.EventHandler(this.frmPackInCarton_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLotSummary)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPack)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtCarton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpRegisterDate;
        private System.Windows.Forms.TextBox txtPack;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnRegisterCarton;
        private System.Windows.Forms.TextBox txtOkCount;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnDeleteSelection;
        private System.Windows.Forms.DataGridView dgvLotSummary;
        private System.Windows.Forms.Button btnCancelCarton;
        private System.Windows.Forms.TextBox txtLoginDept;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtBatch;
        private System.Windows.Forms.DataGridView dgvPack;
        private System.Windows.Forms.Button btnExportModule;
    }
}

