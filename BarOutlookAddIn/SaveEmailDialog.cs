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
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnSaveAttachments_Click(object sender, EventArgs e)
        {
            SelectedOption = SaveOption.SaveAttachmentsOnly;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
