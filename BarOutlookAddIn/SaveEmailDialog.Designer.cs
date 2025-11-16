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
            this.comboBoxEntity.Location = new System.Drawing.Point(16, 89);
            this.comboBoxEntity.Name = "comboBoxEntity";
            this.comboBoxEntity.Size = new System.Drawing.Size(304, 21);
            this.comboBoxEntity.TabIndex = 1;
            // 
            // txtRequestNumber
            // 
            this.txtRequestNumber.Location = new System.Drawing.Point(16, 136);
            this.txtRequestNumber.Name = "txtRequestNumber";
            this.txtRequestNumber.Size = new System.Drawing.Size(304, 20);
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
            this.labelEntity.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelEntity.Location = new System.Drawing.Point(13, 60);
            this.labelEntity.Name = "labelEntity";
            this.labelEntity.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelEntity.Size = new System.Drawing.Size(69, 20);
            this.labelEntity.TabIndex = 0;
            this.labelEntity.Text = "סוג ישות:";
            this.labelEntity.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelRequestNumber
            // 
            this.labelRequestNumber.AutoSize = true;
            this.labelRequestNumber.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.labelRequestNumber.Location = new System.Drawing.Point(14, 113);
            this.labelRequestNumber.Name = "labelRequestNumber";
            this.labelRequestNumber.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.labelRequestNumber.Size = new System.Drawing.Size(91, 20);
            this.labelRequestNumber.TabIndex = 0;
            this.labelRequestNumber.Text = "מספר יישות:";
            this.labelRequestNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.labelRequestNumber.Click += new System.EventHandler(this.labelRequestNumber_Click);
            // 
            // btnSaveEmail
            // 
            this.btnSaveEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.btnSaveEmail.Location = new System.Drawing.Point(16, 218);
            this.btnSaveEmail.Name = "btnSaveEmail";
            this.btnSaveEmail.Size = new System.Drawing.Size(304, 40);
            this.btnSaveEmail.TabIndex = 5;
            this.btnSaveEmail.Text = "שמור את המייל כולו";
            this.btnSaveEmail.Click += new System.EventHandler(this.btnSaveEmail_Click);
            // 
            // btnSaveAttachments
            // 
            this.btnSaveAttachments.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.btnSaveAttachments.Location = new System.Drawing.Point(16, 264);
            this.btnSaveAttachments.Name = "btnSaveAttachments";
            this.btnSaveAttachments.Size = new System.Drawing.Size(304, 36);
            this.btnSaveAttachments.TabIndex = 6;
            this.btnSaveAttachments.Text = "שמור רק קבצים מצורפים";
            this.btnSaveAttachments.Click += new System.EventHandler(this.btnSaveAttachments_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.btnCancel.Location = new System.Drawing.Point(16, 306);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(304, 31);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "ביטול";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnLoadSettings
            // 
            this.btnLoadSettings.Location = new System.Drawing.Point(2, 4);
            this.btnLoadSettings.Name = "btnLoadSettings";
            this.btnLoadSettings.Size = new System.Drawing.Size(104, 23);
            this.btnLoadSettings.TabIndex = 8;
            this.btnLoadSettings.Text = "טעינת הגדרות";
            this.btnLoadSettings.UseVisualStyleBackColor = true;
            this.btnLoadSettings.Click += new System.EventHandler(this.btnLoadSettings_Click);
            // 
            // lblFolderPath
            // 
            this.lblFolderPath.AutoSize = true;
            this.lblFolderPath.Location = new System.Drawing.Point(207, 9);
            this.lblFolderPath.Name = "lblFolderPath";
            this.lblFolderPath.Size = new System.Drawing.Size(114, 13);
            this.lblFolderPath.TabIndex = 9;
            this.lblFolderPath.Text = "לא נבחר נתיב שמירה";
            // 
            // chkCustomFileName
            // 
            this.chkCustomFileName.AutoSize = true;
            this.chkCustomFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.chkCustomFileName.Location = new System.Drawing.Point(16, 162);
            this.chkCustomFileName.Name = "chkCustomFileName";
            this.chkCustomFileName.Size = new System.Drawing.Size(116, 24);
            this.chkCustomFileName.TabIndex = 3;
            this.chkCustomFileName.Text = "הצע שם קובץ";
            this.chkCustomFileName.UseVisualStyleBackColor = true;
            this.chkCustomFileName.CheckedChanged += new System.EventHandler(this.chkCustomFileName_CheckedChanged);
            // 
            // txtCustomFileName
            // 
            this.txtCustomFileName.Enabled = false;
            this.txtCustomFileName.Location = new System.Drawing.Point(16, 192);
            this.txtCustomFileName.Name = "txtCustomFileName";
            this.txtCustomFileName.Size = new System.Drawing.Size(304, 20);
            this.txtCustomFileName.TabIndex = 4;
            this.txtCustomFileName.TextChanged += new System.EventHandler(this.txtCustomFileName_TextChanged);
            // 
            // SaveEmailDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(333, 349);
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
