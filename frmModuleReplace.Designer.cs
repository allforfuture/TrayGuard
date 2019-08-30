namespace TrayGuard
{
    partial class frmModuleReplace
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmModuleReplace));
            this.label2 = new System.Windows.Forms.Label();
            this.txtAfter = new System.Windows.Forms.TextBox();
            this.btnReplace = new System.Windows.Forms.Button();
            this.txtBefore = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtRow = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dgvModule = new System.Windows.Forms.DataGridView();
            this.txtBinSelect = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvModule)).BeginInit();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "替换后: ";
            // 
            // txtAfter
            // 
            this.txtAfter.Location = new System.Drawing.Point(82, 54);
            this.txtAfter.Name = "txtAfter";
            this.txtAfter.Size = new System.Drawing.Size(215, 19);
            this.txtAfter.TabIndex = 5;
            this.txtAfter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAfter_KeyDown);
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(331, 52);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(83, 22);
            this.btnReplace.TabIndex = 11;
            this.btnReplace.Text = "替换";
            this.btnReplace.UseVisualStyleBackColor = true;
            this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // txtBefore
            // 
            this.txtBefore.Enabled = false;
            this.txtBefore.Location = new System.Drawing.Point(82, 20);
            this.txtBefore.Name = "txtBefore";
            this.txtBefore.Size = new System.Drawing.Size(215, 19);
            this.txtBefore.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 12);
            this.label1.TabIndex = 6;
            this.label1.Text = "替换前: ";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(436, 52);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(83, 22);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "关闭";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtRow
            // 
            this.txtRow.Enabled = false;
            this.txtRow.Location = new System.Drawing.Point(417, 20);
            this.txtRow.Name = "txtRow";
            this.txtRow.Size = new System.Drawing.Size(75, 19);
            this.txtRow.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(332, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "目标行: ";
            // 
            // dgvModule
            // 
            this.dgvModule.AllowUserToAddRows = false;
            this.dgvModule.AllowUserToDeleteRows = false;
            this.dgvModule.BackgroundColor = System.Drawing.Color.SkyBlue;
            this.dgvModule.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvModule.Location = new System.Drawing.Point(10, 93);
            this.dgvModule.Name = "dgvModule";
            this.dgvModule.ReadOnly = true;
            this.dgvModule.RowTemplate.Height = 21;
            this.dgvModule.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvModule.Size = new System.Drawing.Size(707, 63);
            this.dgvModule.TabIndex = 12;
            // 
            // txtBinSelect
            // 
            this.txtBinSelect.Enabled = false;
            this.txtBinSelect.Location = new System.Drawing.Point(607, 12);
            this.txtBinSelect.Name = "txtBinSelect";
            this.txtBinSelect.ReadOnly = true;
            this.txtBinSelect.Size = new System.Drawing.Size(100, 19);
            this.txtBinSelect.TabIndex = 13;
            this.txtBinSelect.Visible = false;
            // 
            // frmModuleReplace
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(730, 168);
            this.Controls.Add(this.txtBinSelect);
            this.Controls.Add(this.dgvModule);
            this.Controls.Add(this.txtRow);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtBefore);
            this.Controls.Add(this.txtAfter);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnReplace);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmModuleReplace";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Replace Module";
            this.Load += new System.EventHandler(this.Form4_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvModule)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtAfter;
        private System.Windows.Forms.Button btnReplace;
        private System.Windows.Forms.TextBox txtBefore;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtRow;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dgvModule;
        private System.Windows.Forms.TextBox txtBinSelect;
    }
}

