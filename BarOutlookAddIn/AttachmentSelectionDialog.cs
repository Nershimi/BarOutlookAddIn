using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BarOutlookAddIn
{
    public partial class AttachmentSelectionDialog : Form
    {
        public List<string> SelectedAttachments { get; private set; } = new List<string>();

        public AttachmentSelectionDialog(List<string> attachmentNames)
        {
            InitializeComponent();

            foreach (var name in attachmentNames)
            {
                checkedListBox1.Items.Add(name, true); // ברירת מחדל מסומן
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            foreach (var item in checkedListBox1.CheckedItems)
            {
                SelectedAttachments.Add(item.ToString());
            }
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
