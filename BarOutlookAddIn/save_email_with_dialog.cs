// SaveEmailDialog.cs
using System;
using System.Windows.Forms;

namespace BarOutlookAddIn
{
    public partial class SaveEmailDialog : Form
    {
        public enum SaveOption { Cancel, SaveEmail, SaveAttachmentsOnly }
        public SaveOption SelectedOption { get; private set; } = SaveOption.Cancel;

        public SaveEmailDialog()
        {
            InitializeComponent();
        }

        private void btnSaveEmail_Click(object sender, EventArgs e)
        {
            SelectedOption = SaveOption.SaveEmail;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnSaveAttachments_Click(object sender, EventArgs e)
        {
            SelectedOption = SaveOption.SaveAttachmentsOnly;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
