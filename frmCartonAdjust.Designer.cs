namespace TrayGuard
{
    partial class frmCartonAdjust
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCartonAdjust));
            this.lblAfter = new System.Windows.Forms.Label();
            this.txtAfter = new System.Windows.Forms.TextBox();
            this.btnReplaceCarton = new System.Windows.Forms.Button();
            this.txtBefore = new System.Windows.Forms.TextBox();
            this.lblBefore = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtRow = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dgvCarton = new System.Windows.Forms.DataGridView();
            this.btnAddCarton = new System.Windows.Forms.Button();
            this.btnDeleteCarton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCarton)).BeginInit();
            this.SuspendLayout();
            // 
            // lblAfter
            // 
            this.lblAfter.AutoSize = true;
            this.lblAfter.Location = new System.Drawing.Point(16, 57);
            this.lblAfter.Name = "lblAfter";
            this.lblAfter.Size = new System.Drawing.Size(47, 12);
            this.lblAfter.TabIndex = 6;
            this.lblAfter.Text = "变更后: ";
            // 
            // txtAfter
            // 
            this.txtAfter.Location = new System.Drawing.Point(82, 54);
            this.txtAfter.Name = "txtAfter";
            this.txtAfter.Size = new System.Drawing.Size(215, 19);
            this.txtAfter.TabIndex = 5;
            this.txtAfter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAfter_KeyDown);
            // 
            // btnReplaceCarton
            // 
            this.btnReplaceCarton.Location = new System.Drawing.Point(331, 52);
            this.btnReplaceCarton.Name = "btnReplaceCarton";
            this.btnReplaceCarton.Size = new System.Drawing.Size(83, 22);
            this.btnReplaceCarton.TabIndex = 11;
            this.btnReplaceCarton.Text = "更换纸盒";
            this.btnReplaceCarton.UseVisualStyleBackColor = true;
            this.btnReplaceCarton.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // txtBefore
            // 
            this.txtBefore.Enabled = false;
            this.txtBefore.Location = new System.Drawing.Point(82, 20);
            this.txtBefore.Name = "txtBefore";
            this.txtBefore.Size = new System.Drawing.Size(215, 19);
            this.txtBefore.TabIndex = 5;
            // 
            // lblBefore
            // 
            this.lblBefore.AutoSize = true;
            this.lblBefore.Location = new System.Drawing.Point(16, 23);
            this.lblBefore.Name = "lblBefore";
            this.lblBefore.Size = new System.Drawing.Size(47, 12);
            this.lblBefore.TabIndex = 6;
            this.lblBefore.Text = "变更前: ";
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
            // dgvCarton
            // 
            this.dgvCarton.AllowUserToAddRows = false;
            this.dgvCarton.AllowUserToDeleteRows = false;
            this.dgvCarton.BackgroundColor = System.Drawing.Color.MediumSpringGreen;
            this.dgvCarton.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCarton.Location = new System.Drawing.Point(10, 93);
            this.dgvCarton.Name = "dgvCarton";
            this.dgvCarton.ReadOnly = true;
            this.dgvCarton.RowTemplate.Height = 21;
            this.dgvCarton.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvCarton.Size = new System.Drawing.Size(707, 63);
            this.dgvCarton.TabIndex = 12;
            // 
            // btnAddCarton
            // 
            this.btnAddCarton.Location = new System.Drawing.Point(331, 51);
            this.btnAddCarton.Name = "btnAddCarton";
            this.btnAddCarton.Size = new System.Drawing.Size(83, 22);
            this.btnAddCarton.TabIndex = 13;
            this.btnAddCarton.Text = "加纸箱";
            this.btnAddCarton.UseVisualStyleBackColor = true;
            this.btnAddCarton.Click += new System.EventHandler(this.btnAddCarton_Click);
            // 
            // btnDeleteCarton
            // 
            this.btnDeleteCarton.Location = new System.Drawing.Point(331, 52);
            this.btnDeleteCarton.Name = "btnDeleteCarton";
            this.btnDeleteCarton.Size = new System.Drawing.Size(83, 22);
            this.btnDeleteCarton.TabIndex = 14;
            this.btnDeleteCarton.Text = "删除纸箱";
            this.btnDeleteCarton.UseVisualStyleBackColor = true;
            this.btnDeleteCarton.Click += new System.EventHandler(this.btnDeleteCarton_Click);
            // 
            // frmCartonAdjust
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(730, 168);
            this.Controls.Add(this.btnDeleteCarton);
            this.Controls.Add(this.btnAddCarton);
            this.Controls.Add(this.dgvCarton);
            this.Controls.Add(this.txtRow);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblBefore);
            this.Controls.Add(this.lblAfter);
            this.Controls.Add(this.txtBefore);
            this.Controls.Add(this.txtAfter);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnReplaceCarton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmCartonAdjust";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Replace Carton";
            this.Load += new System.EventHandler(this.frmCartonAdjust_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvCarton)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblAfter;
        private System.Windows.Forms.TextBox txtAfter;
        private System.Windows.Forms.Button btnReplaceCarton;
        private System.Windows.Forms.TextBox txtBefore;
        private System.Windows.Forms.Label lblBefore;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtRow;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dgvCarton;
        private System.Windows.Forms.Button btnAddCarton;
        private System.Windows.Forms.Button btnDeleteCarton;
    }
}

