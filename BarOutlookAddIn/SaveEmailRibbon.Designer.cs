namespace BarOutlookAddIn
{
    partial class SaveEmailRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public SaveEmailRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            // ודא טעינה בכל ההקשרים הרלוונטיים של אאוטלוק
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Inspector";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SaveEmailRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSaveSelectedEmail = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail"; // טאב Home של ה-Explorer
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSaveSelectedEmail);
            this.group1.Label = "שמירה לבר";
            this.group1.Name = "group1";
            // 
            // btnSaveSelectedEmail
            // 
            this.btnSaveSelectedEmail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveSelectedEmail.Label = "שמירה לארכיב"; // שיהיה נראה גם בלי תמונה
            this.btnSaveSelectedEmail.Name = "btnSaveSelectedEmail";
            this.btnSaveSelectedEmail.ShowImage = true;
            try
            {
                this.btnSaveSelectedEmail.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveSelectedEmail.Image")));
            }
            catch { /* אם אין אייקון — לא להפיל */ }
            this.btnSaveSelectedEmail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveSelectedEmail_Click);
            // 
            // SaveEmailRibbon
            // 
            this.Name = "SaveEmailRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Inspector";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SaveEmailRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveSelectedEmail;
    }

    partial class ThisRibbonCollection
    {
        internal SaveEmailRibbon Ribbon1
        {
            get { return this.GetRibbon<SaveEmailRibbon>(); }
        }
    }
}
