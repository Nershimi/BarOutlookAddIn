using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using BarOutlookAddIn.Helpers;
// Removed `using BarOutlookAddIn.App_Code;` to avoid ambiguity if AddInConfig appears elsewhere.
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BarOutlookAddIn
{
    public partial class SaveEmailDialog : Form
    {
        // ===== Context passed from the ribbon =====
        private Outlook.MailItem _mailFromRibbon;
        private List<int> _preselectedAttachmentIndices = new List<int>();
        private List<string> _preselectedAttachmentFileNames = new List<string>();

        // Exposed properties for external callers (saving logic uses these)
        public Outlook.MailItem Mail { get { return _mailFromRibbon; } set { _mailFromRibbon = value; } }
        public Outlook.MailItem MailItem { get { return _mailFromRibbon; } set { _mailFromRibbon = value; } }
        public bool UseCustomFileName => chkCustomFileName.Checked;
        public string CustomFileName => txtCustomFileName.Text?.Trim();

        public IReadOnlyList<int> PreselectedAttachmentIndices
        {
            get { return _preselectedAttachmentIndices.AsReadOnly(); }
            set
            {
                _preselectedAttachmentIndices = value != null ? new List<int>(value) : new List<int>();
            }
        }

        // Alternative name kept for external compatibility
        public IReadOnlyList<int> SelectedAttachmentIndices
        {
            get { return _preselectedAttachmentIndices.AsReadOnly(); }
            set
            {
                _preselectedAttachmentIndices = value != null ? new List<int>(value) : new List<int>();
            }
        }

        public IReadOnlyList<string> PreselectedAttachmentFileNames
        {
            get { return _preselectedAttachmentFileNames.AsReadOnly(); }
            set
            {
                _preselectedAttachmentFileNames = value != null ? new List<string>(value) : new List<string>();
            }
        }

        // Enum the ribbon expects
        public enum SaveOption { None, SaveEmail, SaveAttachments, SaveAttachmentsOnly = SaveAttachments }
        public SaveOption SelectedOption { get; private set; } = SaveOption.None;

        // Configuration loaded from XML (implementation in Helpers.AddInConfig)
        private global::BarOutlookAddIn.Helpers.AddInConfig _config =
            new global::BarOutlookAddIn.Helpers.AddInConfig();

        // Entities loaded from DB for the combo box
        private List<EntityInfo> _entities;

        // Expose the selected EntityInfo for DB insert calls
        public EntityInfo SelectedEntityInfo
        {
            get { return comboBoxEntity != null ? comboBoxEntity.SelectedItem as EntityInfo : null; }
        }

        // Convenience: name only (keeps compatibility with older code)
        public string SelectedEntityName
        {
            get { return SelectedEntityInfo != null ? SelectedEntityInfo.Name : string.Empty; }
        }

        // Unified access to category combo (handles possible designer alias)
        private ComboBox CategoryCombo
        {
            get { return comboBoxCategory ?? comboCategory; }
        }

        public string SelectedCategory
        {
            get { return CategoryCombo != null && CategoryCombo.SelectedItem != null ? CategoryCombo.SelectedItem.ToString() : ""; }
        }

        public string RequestNumber
        {
            get { return txtRequestNumber != null ? txtRequestNumber.Text.Trim() : ""; }
        }

        // ===== Constructor (functionality unchanged) =====
        public SaveEmailDialog()
        {
            DevDiag.Log("Dialog: ctor ENTER");
            InitializeComponent();

            if (comboBoxEntity != null)
                comboBoxEntity.SelectedIndexChanged += comboBoxEntity_SelectedIndexChanged;

            try
            {
                TryLoadConfigFromSavedPath();
                ApplyConfigToUI();
                RestoreLastCategory();
                ApplyDefaultEntityToUI();
                DevDiag.Log("Dialog: about to LoadEntitiesFromDb");
                LoadEntitiesFromDb();
                UpdateFolderLabel();
                DevDiag.Log("Dialog: ctor EXIT OK");
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: ctor EX " + ex.Message);
            }
        }

        // Additional ctor that accepts MailItem and preselected attachment indices
        public SaveEmailDialog(Outlook.MailItem mail, IReadOnlyList<int> preselectedAttachmentIndices)
        {
            DevDiag.Log("Dialog: ctor(mail,indices) ENTER");
            InitializeComponent();

            _mailFromRibbon = mail;
            _preselectedAttachmentIndices = preselectedAttachmentIndices != null
                ? new List<int>(preselectedAttachmentIndices)
                : new List<int>();

            if (comboBoxEntity != null)
                comboBoxEntity.SelectedIndexChanged += comboBoxEntity_SelectedIndexChanged;

            try
            {
                TryLoadConfigFromSavedPath();
                ApplyConfigToUI();
                RestoreLastCategory();
                ApplyDefaultEntityToUI();
                DevDiag.Log("Dialog: about to LoadEntitiesFromDb");
                LoadEntitiesFromDb();
                UpdateFolderLabel();
                DevDiag.Log("Dialog: ctor(mail,indices) EXIT OK");
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: ctor(mail,indices) EX " + ex.Message);
            }
        }

        // Methods for setting context when using the default ctor
        public void SetContext(Outlook.MailItem mail, IReadOnlyList<int> indices)
        {
            _mailFromRibbon = mail;
            _preselectedAttachmentIndices = indices != null ? new List<int>(indices) : new List<int>();
            DevDiag.Log("Dialog: SetContext(mail,indices) set; indices=" + _preselectedAttachmentIndices.Count);
        }

        public void SetContext(Outlook.MailItem mail, IReadOnlyList<int> indices, IReadOnlyList<string> fileNames)
        {
            SetContext(mail, indices);
            _preselectedAttachmentFileNames = fileNames != null ? new List<string>(fileNames) : new List<string>();
            DevDiag.Log("Dialog: SetContext(mail,indices,fileNames) set; files=" + _preselectedAttachmentFileNames.Count);
        }

        public void InitializeWithMailAndAttachments(Outlook.MailItem mail, IReadOnlyList<int> indices)
        {
            SetContext(mail, indices);
        }

        public void InitializeWithMailAndAttachments(Outlook.MailItem mail, IReadOnlyList<int> indices, IReadOnlyList<string> fileNames)
        {
            SetContext(mail, indices, fileNames);
        }

        // ---------------- Internal logic ----------------

        private void TryLoadConfigFromSavedPath()
        {
            try
            {
                string path = Properties.Settings.Default.ConfigPath;
                DevDiag.Log("Dialog: TryLoadConfigFromSavedPath path=" + (path ?? "<null>")
                            + " exists=" + (!string.IsNullOrWhiteSpace(path) && File.Exists(path)));
                if (string.IsNullOrWhiteSpace(path)) return;

                try
                {
                    _config = global::BarOutlookAddIn.Helpers.AddInConfig.Load(path);
                    DevDiag.Log("Dialog: Config loaded. ArchivePath=" + (_config.ArchivePath ?? "<null>")
                                + ", DefaultEntity=" + (_config.DefaultEntity ?? "<null>"));
                }
                catch (Exception ex)
                {
                    DevDiag.Log("Dialog: Config load FAILED " + ex.Message);
                    _config = new global::BarOutlookAddIn.Helpers.AddInConfig();
                }
            }
            catch (Exception exOuter)
            {
                DevDiag.Log("Dialog: TryLoadConfigFromSavedPath EX " + exOuter.Message);
            }
        }

        // Fill categories combo (will be empty if XML has none). Safe for compilation.
        private void ApplyConfigToUI()
        {
            try
            {
                if (CategoryCombo == null)
                {
                    DevDiag.Log("Dialog: ApplyConfigToUI CategoryCombo is null");
                    return;
                }

                CategoryCombo.BeginUpdate();
                CategoryCombo.Items.Clear();

                var cats = _config.Categories ?? new List<string>();
                for (int i = 0; i < cats.Count; i++)
                    CategoryCombo.Items.Add(cats[i]);

                DevDiag.Log("Dialog: ApplyConfigToUI categories count=" + cats.Count);

                if (!string.IsNullOrEmpty(_config.DefaultCategory) &&
                    cats.Contains(_config.DefaultCategory))
                {
                    CategoryCombo.SelectedItem = _config.DefaultCategory;
                    DevDiag.Log("Dialog: ApplyConfigToUI selected DefaultCategory=" + _config.DefaultCategory);
                }
                else if (CategoryCombo.Items.Count > 0)
                {
                    CategoryCombo.SelectedIndex = 0;
                    DevDiag.Log("Dialog: ApplyConfigToUI selected index 0");
                }

                CategoryCombo.EndUpdate();
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: ApplyConfigToUI EX " + ex.Message);
            }
        }

        // Set default entity to UI from XML before DB load (keeps compatibility)
        private void ApplyDefaultEntityToUI()
        {
            try
            {
                string defEntity = _config.DefaultEntity;
                if (string.IsNullOrWhiteSpace(defEntity) || comboBoxEntity == null)
                {
                    DevDiag.Log("Dialog: ApplyDefaultEntityToUI skip (defEntity empty OR combo null)");
                    return;
                }

                bool found = false;
                for (int i = 0; i < comboBoxEntity.Items.Count; i++)
                {
                    string it = comboBoxEntity.Items[i] != null ? comboBoxEntity.Items[i].ToString() : "";
                    if (string.Equals(it, defEntity, StringComparison.OrdinalIgnoreCase))
                    {
                        comboBoxEntity.SelectedIndex = i;
                        found = true;
                        DevDiag.Log("Dialog: ApplyDefaultEntityToUI matched default '" + defEntity + "' at index " + i);
                        break;
                    }
                }
                if (!found)
                {
                    // add a temporary EntityInfo item instead of raw string,
                    // so SelectedEntityInfo remains usable
                    var tmp = new EntityInfo
                    {
                        Name = defEntity,
                        Definement = 0,
                        SystemType = MapDefaultHebrewToSystemType(defEntity) ?? string.Empty
                    };
                    comboBoxEntity.Items.Add(tmp);
                    comboBoxEntity.SelectedItem = tmp;
                    DevDiag.Log("Dialog: ApplyDefaultEntityToUI default '" + defEntity + "' not found; added temporary EntityInfo and selected");
                }
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: ApplyDefaultEntityToUI EX " + ex.Message);
            }
        }

        private void RestoreLastCategory()
        {
            try
            {
                string last = Properties.Settings.Default.LastCategory;
                var cats = _config.Categories ?? new List<string>();
                if (!string.IsNullOrEmpty(last) && cats.Contains(last) && CategoryCombo != null)
                {
                    CategoryCombo.SelectedItem = last;
                    DevDiag.Log("Dialog: RestoreLastCategory -> '" + last + "'");
                }
                else
                {
                    DevDiag.Log("Dialog: RestoreLastCategory skip (last='" + (last ?? "<null>") + "')");
                }
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: RestoreLastCategory EX " + ex.Message);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            TryPersistLastCategory();
        }

        private void TryPersistLastCategory()
        {
            try
            {
                Properties.Settings.Default.LastCategory = SelectedCategory;
                Properties.Settings.Default.Save();
                DevDiag.Log("Dialog: PersistLastCategory -> '" + SelectedCategory + "'");
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: PersistLastCategory EX " + ex.Message);
            }
        }

        // ---------------- Event Handlers (wired in Designer) ----------------

        // Button: "Save whole email"
        private void btnSaveEmail_Click(object sender, EventArgs e)
        {
            DevDiag.Log($"Dialog: btnSaveEmail_Click ENTER - UseCustom?={chkCustomFileName.Checked}, RawName='{txtCustomFileName.Text}'");

            if (UseCustomFileName)
            {
                var raw = CustomFileName;
                if (string.IsNullOrWhiteSpace(raw))
                {
                    MessageBox.Show("Please enter a file name.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DevDiag.Log("Dialog: validation fail - empty custom file name");
                    return;
                }

                // Remove extension if provided and sanitize invalid characters
                var baseName = System.IO.Path.GetFileNameWithoutExtension(raw);
                baseName = CleanFileName(baseName);

                if (string.IsNullOrWhiteSpace(baseName))
                {
                    MessageBox.Show("The entered file name is invalid after sanitization. Try another name.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DevDiag.Log("Dialog: sanitized name empty");
                    return;
                }

                if (!string.Equals(txtCustomFileName.Text, baseName, StringComparison.Ordinal))
                {
                    txtCustomFileName.Text = baseName;
                    DevDiag.Log($"Dialog: sanitized custom name -> '{baseName}'");
                }
            }

            var ent = this.SelectedEntityInfo;
            DevDiag.Log("Dialog: btnSaveEmail_Click with entity -> " +
                (ent != null ? (ent.Name + " | Def=" + ent.Definement + " | Sys=" + ent.SystemType) : "<null>"));

            SelectedOption = SaveOption.SaveEmail;
            TryPersistLastCategory();

            DevDiag.Log($"Dialog: btnSaveEmail_Click EXIT - SelectedOption={SelectedOption}, UseCustom?={chkCustomFileName.Checked}, FinalName='{txtCustomFileName.Text}'");

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // Button: "Save attachments"
        private void btnSaveAttachments_Click(object sender, EventArgs e)
        {
            DevDiag.Log($"Dialog: btnSaveAttachments_Click ENTER - UseCustom?={chkCustomFileName.Checked}, RawName='{txtCustomFileName.Text}', preselectedCount={_preselectedAttachmentIndices?.Count ?? 0}");

            if (UseCustomFileName)
            {
                var raw = CustomFileName;
                if (string.IsNullOrWhiteSpace(raw))
                {
                    MessageBox.Show("Please enter a file name.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DevDiag.Log("Dialog: validation fail - empty custom file name (attachments)");
                    return;
                }

                // Remove extension if provided and sanitize invalid characters
                var baseName = System.IO.Path.GetFileNameWithoutExtension(raw);
                baseName = CleanFileName(baseName);

                if (string.IsNullOrWhiteSpace(baseName))
                {
                    MessageBox.Show("The entered file name is invalid after sanitization. Try another name.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DevDiag.Log("Dialog: sanitized name empty (attachments)");
                    return;
                }

                if (!string.Equals(txtCustomFileName.Text, baseName, StringComparison.Ordinal))
                {
                    txtCustomFileName.Text = baseName;
                    DevDiag.Log($"Dialog: sanitized custom name (attachments) -> '{baseName}'");
                }
            }

            var ent = this.SelectedEntityInfo;
            DevDiag.Log("Dialog: btnSaveAttachments_Click with entity -> " +
                (ent != null ? (ent.Name + " | Def=" + ent.Definement + " | Sys=" + ent.SystemType) : "<null>") +
                $" | preselectedCount={_preselectedAttachmentIndices?.Count ?? 0}");

            SelectedOption = SaveOption.SaveAttachments; // matches ribbon logic (SaveAttachments / SaveAttachmentsOnly)
            TryPersistLastCategory();

            DevDiag.Log($"Dialog: btnSaveAttachments_Click EXIT - SelectedOption={SelectedOption}, UseCustom?={chkCustomFileName.Checked}, FinalName='{txtCustomFileName.Text}'");

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // Cancel button
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DevDiag.Log("Dialog: btnCancel_Click");
            SelectedOption = SaveOption.None;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // Load settings button
        private void btnLoadSettings_Click(object sender, EventArgs e)
        {
            DevDiag.Log("Dialog: btnLoadSettings_Click ENTER");
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select configuration file (XML)";
                ofd.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                ofd.CheckFileExists = true;
                ofd.Multiselect = false;

                string savedPath = Properties.Settings.Default.ConfigPath;
                DevDiag.Log("Dialog: btnLoadSettings savedPath=" + (savedPath ?? "<null>"));
                if (!string.IsNullOrWhiteSpace(savedPath))
                {
                    try
                    {
                        ofd.InitialDirectory = Path.GetDirectoryName(savedPath);
                        ofd.FileName = Path.GetFileName(savedPath);
                    }
                    catch { }
                }

                if (ofd.ShowDialog(this) != DialogResult.OK)
                {
                    DevDiag.Log("Dialog: btnLoadSettings canceled by user");
                    return;
                }

                try
                {
                    var newCfg = global::BarOutlookAddIn.Helpers.AddInConfig.Load(ofd.FileName);
                    DevDiag.Log("Dialog: btnLoadSettings loaded cfg from " + ofd.FileName + " | ArchivePath=" + (newCfg.ArchivePath ?? "<null>"));

                    Properties.Settings.Default.ConfigPath = ofd.FileName;

                    // If XML includes archive path, store it into SaveBaseFolder (existing behavior)
                    if (!string.IsNullOrWhiteSpace(newCfg.ArchivePath))
                    {
                        try
                        {
                            var props = Properties.Settings.Default.Properties;
                            if (props != null && props["SaveBaseFolder"] != null)
                            {
                                Properties.Settings.Default["SaveBaseFolder"] = newCfg.ArchivePath;
                                DevDiag.Log("Dialog: btnLoadSettings SaveBaseFolder set -> " + newCfg.ArchivePath);
                            }
                        }
                        catch (Exception exSet)
                        {
                            DevDiag.Log("Dialog: btnLoadSettings set SaveBaseFolder EX " + exSet.Message);
                        }
                    }

                    // NEW: If XML includes SQL settings, build and persist a ConnectionString so DB calls work immediately
                    try
                    {
                        bool setConn = false;
                        var props = Properties.Settings.Default.Properties;
                        // If config has server+db, create connection string
                        if (!string.IsNullOrWhiteSpace(newCfg.SqlServerAddress) && !string.IsNullOrWhiteSpace(newCfg.SqlDbName))
                        {
                            try
                            {
                                var builder = new System.Data.SqlClient.SqlConnectionStringBuilder();
                                builder.DataSource = newCfg.SqlServerAddress;
                                builder.InitialCatalog = newCfg.SqlDbName;

                                if (!string.IsNullOrWhiteSpace(newCfg.SqlUserName))
                                {
                                    builder.UserID = newCfg.SqlUserName;
                                    builder.Password = newCfg.SqlPassword ?? "";
                                    builder.IntegratedSecurity = false;
                                }
                                else
                                {
                                    // When user name not set assume integrated security
                                    builder.IntegratedSecurity = true;
                                }

                                if (props != null && props["ConnectionString"] != null)
                                {
                                    Properties.Settings.Default["ConnectionString"] = builder.ConnectionString;
                                    DevDiag.Log("Dialog: btnLoadSettings ConnectionString set -> " + builder.ConnectionString);
                                }
                                setConn = true;
                            }
                            catch (Exception exCs)
                            {
                                DevDiag.Log("Dialog: btnLoadSettings build ConnectionString EX " + exCs.Message);
                            }
                        }

                        // Also persist individual SQL settings used elsewhere (best-effort)
                        try
                        {
                            if (props != null)
                            {
                                if (!string.IsNullOrWhiteSpace(newCfg.SqlServerAddress) && props["ServerAddressNB"] != null)
                                    Properties.Settings.Default["ServerAddressNB"] = newCfg.SqlServerAddress;
                                if (!string.IsNullOrWhiteSpace(newCfg.SqlDbName) && props["NBDBName"] != null)
                                    Properties.Settings.Default["NBDBName"] = newCfg.SqlDbName;
                                if (!string.IsNullOrWhiteSpace(newCfg.SqlUserName) && props["SqlUserName"] != null)
                                    Properties.Settings.Default["SqlUserName"] = newCfg.SqlUserName;
                                if (!string.IsNullOrWhiteSpace(newCfg.SqlPassword) && props["SqlPassword"] != null)
                                    Properties.Settings.Default["SqlPassword"] = newCfg.SqlPassword;
                                if (setConn)
                                {
                                    DevDiag.Log("Dialog: btnLoadSettings persisted individual SQL settings (best-effort)");
                                }
                            }
                        }
                        catch (Exception exSetSql)
                        {
                            DevDiag.Log("Dialog: btnLoadSettings persist individual SQL settings EX " + exSetSql.Message);
                        }
                    }
                    catch (Exception) { }

                    try { Properties.Settings.Default.Save(); DevDiag.Log("Dialog: btnLoadSettings Settings saved"); } catch { }

                    _config = newCfg;
                    ApplyConfigToUI();
                    RestoreLastCategory();
                    ApplyDefaultEntityToUI();
                    LoadEntitiesFromDb();
                    UpdateFolderLabel();

                    DevDiag.Log("Dialog: btnLoadSettings EXIT OK");
                    MessageBox.Show("Settings loaded and saved successfully.",
                                    "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    DevDiag.Log("Dialog: btnLoadSettings EX " + ex.Message);
                    MessageBox.Show("Error loading file:\r\n" + ex.Message,
                                    "Load settings failed",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UpdateFolderLabel()
        {
            try
            {
                if (this.lblFolderPath == null)
                {
                    DevDiag.Log("Dialog: UpdateFolderLabel skip (lblFolderPath is null)");
                    return;
                }

                string folderPath = null;

                try
                {
                    var props = Properties.Settings.Default.Properties;
                    if (props["SaveBaseFolder"] != null)
                    {
                        var val = Properties.Settings.Default["SaveBaseFolder"];
                        if (val != null) folderPath = val as string;
                    }
                }
                catch { }

                if (string.IsNullOrWhiteSpace(folderPath) && _config != null && !string.IsNullOrWhiteSpace(_config.ArchivePath))
                    folderPath = _config.ArchivePath;

                this.lblFolderPath.Text = string.IsNullOrWhiteSpace(folderPath)
                    ? "לא נבחר נתיב לשמירה"
                    : "נתיב שמירה תקין";

                DevDiag.Log("Dialog: UpdateFolderLabel -> '" + this.lblFolderPath.Text + "'");
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: UpdateFolderLabel EX " + ex.Message);
                if (this.lblFolderPath != null)
                    this.lblFolderPath.Text = "תקלה בקריאת נתיב שמירה";
            }
        }

        private static string MapDefaultHebrewToSystemType(string heb)
        {
            if (string.IsNullOrWhiteSpace(heb)) return null;
            heb = heb.Trim();
            if (heb == "בקשות") return "ב";
            if (heb == "פיקוח") return "פ";
            if (heb == "ישות תכנונית") return "ת";
            if (heb == "ישות כללית") return "כ";
            if (heb == "תביעה") return "ע";
            return null;
        }

        private void LoadEntitiesFromDb()
        {
            try
            {
                var cs = global::BarOutlookAddIn.Properties.Settings.Default.ConnectionString;
                DevDiag.Log("Dialog: LoadEntitiesFromDb cs exists? " + (!string.IsNullOrWhiteSpace(cs)));

                // Remember previous selection (by Name + Definement + SystemType)
                var prev = this.SelectedEntityInfo;

                var repo = new EntityRepository(cs);
                _entities = repo.GetEntities();
                DevDiag.Log("Dialog: LoadEntitiesFromDb entities count=" + (_entities != null ? _entities.Count : -1));

                if (comboBoxEntity == null) return;

                comboBoxEntity.BeginUpdate();
                comboBoxEntity.DataSource = null;
                comboBoxEntity.Items.Clear();

                if (_entities != null && _entities.Count > 0)
                {
                    comboBoxEntity.DisplayMember = "DisplayText";
                    comboBoxEntity.ValueMember = null;
                    comboBoxEntity.DataSource = _entities;

                    // Try to restore previous selection first
                    bool restored = false;
                    if (prev != null)
                    {
                        int idxPrev = _entities.FindIndex(e =>
                            string.Equals((e.Name ?? "").Trim(), (prev.Name ?? "").Trim(), StringComparison.OrdinalIgnoreCase) &&
                            e.Definement == prev.Definement &&
                            string.Equals(e.SystemType, prev.SystemType, StringComparison.OrdinalIgnoreCase));

                        if (idxPrev >= 0)
                        {
                            comboBoxEntity.SelectedIndex = idxPrev;
                            restored = true;
                            DevDiag.Log("Dialog: LoadEntitiesFromDb restored previous selection -> index " + idxPrev);
                        }
                    }

                    if (!restored)
                    {
                        // default mapping by XML "DefaultEntity" (e.g., "בקשות"→"ב")
                        string sys = MapDefaultHebrewToSystemType(_config.DefaultEntity);
                        int idx = -1;

                        if (!string.IsNullOrWhiteSpace(_config.DefaultEntity))
                        {
                            // 1) exact by Name
                            idx = _entities.FindIndex(e =>
                                string.Equals((e.Name ?? "").Trim(), _config.DefaultEntity.Trim(), StringComparison.OrdinalIgnoreCase));
                        }

                        // 2) fallback by SystemType
                        if (idx < 0 && !string.IsNullOrWhiteSpace(sys))
                            idx = _entities.FindIndex(e => string.Equals(e.SystemType, sys, StringComparison.OrdinalIgnoreCase));

                        if (idx >= 0)
                        {
                            comboBoxEntity.SelectedIndex = idx;
                            DevDiag.Log("Dialog: LoadEntitiesFromDb matched default '" + _config.DefaultEntity + "' -> index " + idx +
                                        " (Name=" + _entities[idx].Name + ", SystemType=" + _entities[idx].SystemType + ")");
                        }
                        else
                        {
                            comboBoxEntity.SelectedIndex = 0;
                            DevDiag.Log("Dialog: LoadEntitiesFromDb default not found; selected index 0");
                        }
                    }
                }

                comboBoxEntity.EndUpdate();
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: LoadEntitiesFromDb EX " + ex.Message);
            }
        }

        private void comboBoxEntity_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var ent = this.SelectedEntityInfo;
                DevDiag.Log("Dialog: Entity changed -> " +
                    (ent != null ? (ent.Name + " | Def=" + ent.Definement + " | Sys=" + ent.SystemType) : "<null>")
                    + " (index=" + comboBoxEntity.SelectedIndex + ")");
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: Entity changed EX " + ex.Message);
            }
        }
        private void chkCustomFileName_CheckedChanged(object sender, EventArgs e)
        {
            txtCustomFileName.Enabled = chkCustomFileName.Checked;
            if (!chkCustomFileName.Checked)
                txtCustomFileName.Text = "";
        }

        // Helper to sanitize file name (returns empty for blank input so caller treats it as invalid)
        private static string CleanFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return string.Empty;
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim();
        }

        private static string Sanitize(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim();
        }


        private void labelEntity_Click(object sender, EventArgs e)
        {

        }

        private void SaveEmailDialog_Load(object sender, EventArgs e)
        {

        }

        private void txtCustomFileName_TextChanged(object sender, EventArgs e)
        {

        }

        private void labelRequestNumber_Click(object sender, EventArgs e)
        {

        }
    }
}
