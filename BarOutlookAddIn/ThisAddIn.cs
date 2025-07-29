using System;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BarOutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // You can initialize logic here if needed
            System.Windows.Forms.MessageBox.Show("התוסף עלה בהצלחה!");

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Optional cleanup logic
        }

        /// <summary>
        /// Connects the custom Ribbon UI to Outlook.
        /// </summary>
        /// <returns>Your ribbon class instance</returns>

        /// <summary>
        /// Registers the startup and shutdown handlers
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }


    }
}
