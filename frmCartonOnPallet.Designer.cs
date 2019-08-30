

namespace TrayGuard
{
    partial class frmCartonOnPallet
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCartonOnPallet));
            this.btnClose = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.txtLoginName = new System.Windows.Forms.TextBox();
            this.dgvCarton = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPallet = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpRegisterDate = new System.Windows.Forms.DateTimePicker();
            this.txtCarton = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnRegisterPallet = new System.Windows.Forms.Button();
            this.txtOkCount = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnDeleteSelection = new System.Windows.Forms.Button();
            this.dgvLotSummary = new System.Windows.Forms.DataGridView();
            this.btnCancelPallet = new System.Windows.Forms.Button();
            this.txtLoginDept = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtBatch = new System.Windows.Forms.TextBox();
            this.btnExportModule = new System.Windows.Forms.Button();
            this.btnReplaceCarton = new System.Windows.Forms.Button();
            this.btnDeleteCarton = new System.Windows.Forms.Button();
            this.btnAddCarton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCarton)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLotSummary)).BeginInit();
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
            // dgvCarton
            // 
            this.dgvCarton.AllowUserToAddRows = false;
            this.dgvCarton.AllowUserToDeleteRows = false;
            this.dgvCarton.BackgroundColor = System.Drawing.Color.MediumSpringGreen;
            this.dgvCarton.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCarton.Location = new System.Drawing.Point(12, 215);
            this.dgvCarton.Name = "dgvCarton";
            this.dgvCarton.ReadOnly = true;
            this.dgvCarton.RowTemplate.Height = 21;
            this.dgvCarton.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvCarton.Size = new System.Drawing.Size(764, 474);
            this.dgvCarton.TabIndex = 9;
            this.dgvCarton.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPack_CellContentClick);
            this.dgvCarton.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPack_CellContentDoubleClick);
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
            // txtPallet
            // 
            this.txtPallet.Enabled = false;
            this.txtPallet.Location = new System.Drawing.Point(117, 97);
            this.txtPallet.Name = "txtPallet";
            this.txtPallet.Size = new System.Drawing.Size(160, 19);
            this.txtPallet.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "卡板ID: ";
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
            // txtCarton
            // 
            this.txtCarton.Enabled = false;
            this.txtCarton.Location = new System.Drawing.Point(117, 177);
            this.txtCarton.Name = "txtCarton";
            this.txtCarton.Size = new System.Drawing.Size(160, 19);
            this.txtCarton.TabIndex = 5;
            this.txtCarton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtModuleId_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 182);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "纸箱ID: ";
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
            // btnRegisterPallet
            // 
            this.btnRegisterPallet.Enabled = false;
            this.btnRegisterPallet.Location = new System.Drawing.Point(295, 175);
            this.btnRegisterPallet.Name = "btnRegisterPallet";
            this.btnRegisterPallet.Size = new System.Drawing.Size(100, 22);
            this.btnRegisterPallet.TabIndex = 11;
            this.btnRegisterPallet.Text = "登记卡板";
            this.btnRegisterPallet.UseVisualStyleBackColor = true;
            this.btnRegisterPallet.Click += new System.EventHandler(this.btnRegisterTray_Click);
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
            this.dgvLotSummary.BackgroundColor = System.Drawing.Color.MediumSpringGreen;
            this.dgvLotSummary.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dgvLotSummary.GridColor = System.Drawing.Color.Azure;
            this.dgvLotSummary.Location = new System.Drawing.Point(117, 12);
            this.dgvLotSummary.Name = "dgvLotSummary";
            this.dgvLotSummary.ReadOnly = true;
            this.dgvLotSummary.RowTemplate.Height = 21;
            this.dgvLotSummary.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvLotSummary.Size = new System.Drawing.Size(636, 61);
            this.dgvLotSummary.TabIndex = 9;
            this.dgvLotSummary.ColumnAdded += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dgvLotSummary_ColumnAdded);
            // 
            // btnCancelPallet
            // 
            this.btnCancelPallet.Enabled = false;
            this.btnCancelPallet.Location = new System.Drawing.Point(541, 175);
            this.btnCancelPallet.Name = "btnCancelPallet";
            this.btnCancelPallet.Size = new System.Drawing.Size(100, 22);
            this.btnCancelPallet.TabIndex = 43;
            this.btnCancelPallet.Text = "取消卡板";
            this.btnCancelPallet.UseVisualStyleBackColor = true;
            this.btnCancelPallet.Click += new System.EventHandler(this.btnCancelCarton_Click);
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
            // btnExportModule
            // 
            this.btnExportModule.Location = new System.Drawing.Point(16, 42);
            this.btnExportModule.Name = "btnExportModule";
            this.btnExportModule.Size = new System.Drawing.Size(70, 22);
            this.btnExportModule.TabIndex = 43;
            this.btnExportModule.Text = "输出马达";
            this.btnExportModule.UseVisualStyleBackColor = true;
            this.btnExportModule.Click += new System.EventHandler(this.btnExportModule_Click);
            // 
            // btnReplaceCarton
            // 
            this.btnReplaceCarton.Enabled = false;
            this.btnReplaceCarton.Location = new System.Drawing.Point(295, 150);
            this.btnReplaceCarton.Name = "btnReplaceCarton";
            this.btnReplaceCarton.Size = new System.Drawing.Size(100, 22);
            this.btnReplaceCarton.TabIndex = 44;
            this.btnReplaceCarton.Text = "替换纸箱";
            this.btnReplaceCarton.UseVisualStyleBackColor = true;
            this.btnReplaceCarton.Visible = false;
            this.btnReplaceCarton.Click += new System.EventHandler(this.btnReplaceCarton_Click);
            // 
            // btnDeleteCarton
            // 
            this.btnDeleteCarton.Enabled = false;
            this.btnDeleteCarton.Location = new System.Drawing.Point(417, 150);
            this.btnDeleteCarton.Name = "btnDeleteCarton";
            this.btnDeleteCarton.Size = new System.Drawing.Size(100, 22);
            this.btnDeleteCarton.TabIndex = 44;
            this.btnDeleteCarton.Text = "删除纸箱";
            this.btnDeleteCarton.UseVisualStyleBackColor = true;
            this.btnDeleteCarton.Visible = false;
            this.btnDeleteCarton.Click += new System.EventHandler(this.btnDeleteCarton_Click);
            // 
            // btnAddCarton
            // 
            this.btnAddCarton.Enabled = false;
            this.btnAddCarton.Location = new System.Drawing.Point(541, 150);
            this.btnAddCarton.Name = "btnAddCarton";
            this.btnAddCarton.Size = new System.Drawing.Size(100, 22);
            this.btnAddCarton.TabIndex = 44;
            this.btnAddCarton.Text = "加纸箱";
            this.btnAddCarton.UseVisualStyleBackColor = true;
            this.btnAddCarton.Visible = false;
            this.btnAddCarton.Click += new System.EventHandler(this.btnAddCarton_Click);
            // 
            // frmCartonOnPallet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(788, 701);
            this.Controls.Add(this.btnAddCarton);
            this.Controls.Add(this.btnDeleteCarton);
            this.Controls.Add(this.btnReplaceCarton);
            this.Controls.Add(this.dgvCarton);
            this.Controls.Add(this.btnExportModule);
            this.Controls.Add(this.btnCancelPallet);
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
            this.Controls.Add(this.txtCarton);
            this.Controls.Add(this.txtPallet);
            this.Controls.Add(this.txtBatch);
            this.Controls.Add(this.txtOkCount);
            this.Controls.Add(this.txtLoginDept);
            this.Controls.Add(this.txtLoginName);
            this.Controls.Add(this.btnRegisterPallet);
            this.Controls.Add(this.btnDeleteSelection);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmCartonOnPallet";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "4-B. Carton on Pallet";
            this.Load += new System.EventHandler(this.frmCartonOnPallet_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCarton)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLotSummary)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtLoginName;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.DataGridView dgvCarton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPallet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpRegisterDate;
        private System.Windows.Forms.TextBox txtCarton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnRegisterPallet;
        private System.Windows.Forms.TextBox txtOkCount;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnDeleteSelection;
        private System.Windows.Forms.DataGridView dgvLotSummary;
        private System.Windows.Forms.Button btnCancelPallet;
        private System.Windows.Forms.TextBox txtLoginDept;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtBatch;
        private System.Windows.Forms.Button btnExportModule;
        private System.Windows.Forms.Button btnReplaceCarton;
        private System.Windows.Forms.Button btnDeleteCarton;
        private System.Windows.Forms.Button btnAddCarton;
    }
}

