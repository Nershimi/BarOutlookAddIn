namespace BarOutlookAddIn
{
    partial class AttachmentSelectionDialog
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();

            // checkedListBox1
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 12);
            this.checkedListBox1.Size = new System.Drawing.Size(360, 244);
            this.checkedListBox1.TabIndex = 0;

            // btnOK
            this.btnOK.Location = new System.Drawing.Point(216, 270);
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "אישור";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);

            // btnCancel
            this.btnCancel.Location = new System.Drawing.Point(297, 270);
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "ביטול";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            // AttachmentSelectionDialog
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 311);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AttachmentSelectionDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "בחר קבצים לשמירה";
            this.ResumeLayout(false);
        }
    }
}
