using System.Windows.Forms;

namespace BarOutlookAddIn
{
    partial class SaveEmailDialog
    {
        private System.ComponentModel.IContainer components = null;

        private ComboBox comboBoxCategory;
        private ComboBox comboBoxEntity;
        private TextBox txtRequestNumber;
        private Label labelCategory;
        private Label labelEntity;
        private Label labelRequestNumber;
        private Button btnSaveEmail;
        private Button btnSaveAttachments;
        private Button btnCancel;
        private Button btnLoadSettings;
        private System.Windows.Forms.Label lblFolderPath;
        private System.Windows.Forms.CheckBox chkCustomFileName;
        private System.Windows.Forms.TextBox txtCustomFileName;



        // שדה עזר – לשמירת תאימות לקוד שמצפה לשם comboCategory
        private ComboBox comboCategory;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SaveEmailDialog));
            this.comboBoxEntity = new System.Windows.Forms.ComboBox();
            this.txtRequestNumber = new System.Windows.Forms.TextBox();
            this.labelCategory = new System.Windows.Forms.Label();
            this.labelEntity = new System.Windows.Forms.Label();
            this.labelRequestNumber = new System.Windows.Forms.Label();
            this.btnSaveEmail = new System.Windows.Forms.Button();
            this.btnSaveAttachments = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnLoadSettings = new System.Windows.Forms.Button();
            this.lblFolderPath = new System.Windows.Forms.Label();
            this.chkCustomFileName = new System.Windows.Forms.CheckBox();
            this.txtCustomFileName = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // comboBoxEntity
            // 
            this.comboBoxEntity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxEntity.Location = new System.Drawing.Point(17, 64);
            this.comboBoxEntity.Name = "comboBoxEntity";
            this.comboBoxEntity.Size = new System.Drawing.Size(215, 21);
            this.comboBoxEntity.TabIndex = 1;
            // 
            // txtRequestNumber
            // 
            this.txtRequestNumber.Location = new System.Drawing.Point(17, 110);
            this.txtRequestNumber.Name = "txtRequestNumber";
            this.txtRequestNumber.Size = new System.Drawing.Size(215, 20);
            this.txtRequestNumber.TabIndex = 2;
            // 
            // labelCategory
            // 
            this.labelCategory.Location = new System.Drawing.Point(0, 0);
            this.labelCategory.Name = "labelCategory";
            this.labelCategory.Size = new System.Drawing.Size(100, 23);
            this.labelCategory.TabIndex = 0;
            // 
            // labelEntity
            // 
            this.labelEntity.AutoSize = true;
            this.labelEntity.Location = new System.Drawing.Point(23, 43);
            this.labelEntity.Name = "labelEntity";
            this.labelEntity.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelEntity.Size = new System.Drawing.Size(57, 13);
            this.labelEntity.TabIndex = 0;
            this.labelEntity.Text = "סוג ישות:";
            this.labelEntity.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelRequestNumber
            // 
            this.labelRequestNumber.AutoSize = true;
            this.labelRequestNumber.Location = new System.Drawing.Point(23, 94);
            this.labelRequestNumber.Name = "labelRequestNumber";
            this.labelRequestNumber.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelRequestNumber.Size = new System.Drawing.Size(70, 13);
            this.labelRequestNumber.TabIndex = 0;
            this.labelRequestNumber.Text = "מספר בקשה:";
            this.labelRequestNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnSaveEmail
            // 
            this.btnSaveEmail.Location = new System.Drawing.Point(17, 190);
            this.btnSaveEmail.Name = "btnSaveEmail";
            this.btnSaveEmail.Size = new System.Drawing.Size(214, 26);
            this.btnSaveEmail.TabIndex = 5;
            this.btnSaveEmail.Text = "שמור את המייל כולו";
            this.btnSaveEmail.Click += new System.EventHandler(this.btnSaveEmail_Click);
            // 
            // btnSaveAttachments
            // 
            this.btnSaveAttachments.Location = new System.Drawing.Point(17, 225);
            this.btnSaveAttachments.Name = "btnSaveAttachments";
            this.btnSaveAttachments.Size = new System.Drawing.Size(214, 26);
            this.btnSaveAttachments.TabIndex = 6;
            this.btnSaveAttachments.Text = "שמור רק קבצים מצורפים";
            this.btnSaveAttachments.Click += new System.EventHandler(this.btnSaveAttachments_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(17, 260);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(214, 26);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "ביטול";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnLoadSettings
            // 
            this.btnLoadSettings.Location = new System.Drawing.Point(17, 13);
            this.btnLoadSettings.Name = "btnLoadSettings";
            this.btnLoadSettings.Size = new System.Drawing.Size(103, 23);
            this.btnLoadSettings.TabIndex = 8;
            this.btnLoadSettings.Text = "טעינת הגדרות";
            this.btnLoadSettings.UseVisualStyleBackColor = true;
            this.btnLoadSettings.Click += new System.EventHandler(this.btnLoadSettings_Click);
            // 
            // lblFolderPath
            // 
            this.lblFolderPath.AutoSize = true;
            this.lblFolderPath.Location = new System.Drawing.Point(131, 18);
            this.lblFolderPath.Name = "lblFolderPath";
            this.lblFolderPath.Size = new System.Drawing.Size(114, 13);
            this.lblFolderPath.TabIndex = 9;
            this.lblFolderPath.Text = "לא נבחר נתיב שמירה";
            // 
            // chkCustomFileName
            // 
            this.chkCustomFileName.AutoSize = true;
            this.chkCustomFileName.Location = new System.Drawing.Point(17, 136);
            this.chkCustomFileName.Name = "chkCustomFileName";
            this.chkCustomFileName.Size = new System.Drawing.Size(96, 17);
            this.chkCustomFileName.TabIndex = 3;
            this.chkCustomFileName.Text = "הצע שם קובץ";
            this.chkCustomFileName.UseVisualStyleBackColor = true;
            this.chkCustomFileName.CheckedChanged += new System.EventHandler(this.chkCustomFileName_CheckedChanged);
            // 
            // txtCustomFileName
            // 
            this.txtCustomFileName.Enabled = false;
            this.txtCustomFileName.Location = new System.Drawing.Point(37, 158);
            this.txtCustomFileName.Name = "txtCustomFileName";
            this.txtCustomFileName.Size = new System.Drawing.Size(195, 20);
            this.txtCustomFileName.TabIndex = 4;
            // 
            // SaveEmailDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(257, 300);
            this.Controls.Add(this.chkCustomFileName);
            this.Controls.Add(this.txtCustomFileName);
            this.Controls.Add(this.lblFolderPath);
            this.Controls.Add(this.btnLoadSettings);
            this.Controls.Add(this.labelEntity);
            this.Controls.Add(this.comboBoxEntity);
            this.Controls.Add(this.labelRequestNumber);
            this.Controls.Add(this.txtRequestNumber);
            this.Controls.Add(this.btnSaveEmail);
            this.Controls.Add(this.btnSaveAttachments);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SaveEmailDialog";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.Text = "שמירת מייל";
            this.Load += new System.EventHandler(this.SaveEmailDialog_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
