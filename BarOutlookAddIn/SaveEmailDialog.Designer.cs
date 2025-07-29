using System;
using System.Windows.Forms;  // ← שורה חשובה שחסרה אצלך
namespace BarOutlookAddIn
{
    partial class SaveEmailDialog
    {
        private System.ComponentModel.IContainer components = null;

        private Button btnSaveEmail;
        private Button btnSaveAttachments;
        private Button btnCancel;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnSaveEmail = new System.Windows.Forms.Button();
            this.btnSaveAttachments = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSaveEmail
            // 
            this.btnSaveEmail.Location = new System.Drawing.Point(30, 30);
            this.btnSaveEmail.Name = "btnSaveEmail";
            this.btnSaveEmail.Size = new System.Drawing.Size(200, 30);
            this.btnSaveEmail.Text = "שמור את המייל כולו";
            this.btnSaveEmail.UseVisualStyleBackColor = true;
            this.btnSaveEmail.Click += new System.EventHandler(this.btnSaveEmail_Click);
            // 
            // btnSaveAttachments
            // 
            this.btnSaveAttachments.Location = new System.Drawing.Point(30, 70);
            this.btnSaveAttachments.Name = "btnSaveAttachments";
            this.btnSaveAttachments.Size = new System.Drawing.Size(200, 30);
            this.btnSaveAttachments.Text = "שמור רק קבצים מצורפים";
            this.btnSaveAttachments.UseVisualStyleBackColor = true;
            this.btnSaveAttachments.Click += new System.EventHandler(this.btnSaveAttachments_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(30, 110);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(200, 30);
            this.btnCancel.Text = "ביטול";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SaveEmailDialog
            // 
            this.ClientSize = new System.Drawing.Size(260, 170);
            this.Controls.Add(this.btnSaveEmail);
            this.Controls.Add(this.btnSaveAttachments);
            this.Controls.Add(this.btnCancel);
            this.Name = "SaveEmailDialog";
            this.Text = "שמירת מייל";
            this.ResumeLayout(false);
        }
    }
}
