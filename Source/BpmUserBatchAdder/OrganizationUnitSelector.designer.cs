namespace BpmUserBatchAdder {
    partial class OrganizationUnitSelector {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.cbxOrganizationUnit = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // cbxOrganizationUnit
            // 
            this.cbxOrganizationUnit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cbxOrganizationUnit.FormattingEnabled = true;
            this.cbxOrganizationUnit.Location = new System.Drawing.Point(0, 0);
            this.cbxOrganizationUnit.Name = "cbxOrganizationUnit";
            this.cbxOrganizationUnit.Size = new System.Drawing.Size(490, 23);
            this.cbxOrganizationUnit.TabIndex = 0;
            this.cbxOrganizationUnit.KeyUp += new System.Windows.Forms.KeyEventHandler(this.cbxOrganizationUnit_KeyUp);
            // 
            // OrganizationUnitSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 26);
            this.Controls.Add(this.cbxOrganizationUnit);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OrganizationUnitSelector";
            this.ShowIcon = false;
            this.Text = "選擇所屬部門";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.OrganizationUnitSelector_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cbxOrganizationUnit;
    }
}