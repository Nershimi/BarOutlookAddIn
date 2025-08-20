using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using BarOutlookAddIn.Helpers;
// הסר את using BarOutlookAddIn.App_Code; כדי למנוע דו-משמעות אם יש גם שם AddInConfig

namespace BarOutlookAddIn
{
    public partial class SaveEmailDialog : Form
    {
        // enum שהריבון מצפה לו
        public enum SaveOption { None, SaveEmail, SaveAttachments, SaveAttachmentsOnly = SaveAttachments }
        public SaveOption SelectedOption { get; private set; } = SaveOption.None;

        // קונפיג מ־XML (משתמש במימוש שב-Helpers)
        private global::BarOutlookAddIn.Helpers.AddInConfig _config =
            new global::BarOutlookAddIn.Helpers.AddInConfig();

        // Holds the entities loaded from DB for the combo
        private List<EntityInfo> _entities;

        // Expose the selected entity object for DB insert
        public EntityInfo SelectedEntityInfo
        {
            get { return comboBoxEntity != null ? comboBoxEntity.SelectedItem as EntityInfo : null; }
        }

        // Convenience: name only (used sometimes by older code)
        public string SelectedEntityName
        {
            get { return SelectedEntityInfo != null ? SelectedEntityInfo.Name : string.Empty; }
        }

        // גישה אחידה לקומבו קטגוריה (בין comboBoxCategory ל-comboCategory אם נוצר alias ב-Designer)
        private ComboBox CategoryCombo
        {
            get { return (object)comboBoxCategory != null ? comboBoxCategory : comboCategory; }
        }

        public string SelectedCategory
        {
            get { return CategoryCombo != null && CategoryCombo.SelectedItem != null ? CategoryCombo.SelectedItem.ToString() : ""; }
        }

        public string RequestNumber
        {
            get { return txtRequestNumber != null ? txtRequestNumber.Text.Trim() : ""; }
        }

        public SaveEmailDialog()
        {
            DevDiag.Log("Dialog: ctor ENTER");
            InitializeComponent();

            // עקוב אחרי שינוי ישות
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

        // ---------------- לוגיקה פנימית ----------------

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

        // ממלא את הקומבו של הקטגוריות (יישאר ריק אם אין ב-XML), בטוח לקומפילציה
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

        // קובע ישות ברירת מחדל מתוך ה-XML אם קיימת (ל-comboBoxEntity) — לפני טעינת DB (ייצור זמני)
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
                    comboBoxEntity.Items.Add(defEntity);
                    comboBoxEntity.SelectedItem = defEntity;
                    DevDiag.Log("Dialog: ApplyDefaultEntityToUI default '" + defEntity + "' not found; added and selected");
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

        // ---------------- Event Handlers (מחוברים מה-Designer) ----------------

        // כפתור: "שמור את המייל כולו"
        private void btnSaveEmail_Click(object sender, EventArgs e)
        {
            var ent = this.SelectedEntityInfo;
            DevDiag.Log("Dialog: btnSaveEmail_Click with entity -> " +
                (ent != null ? (ent.Name + " | Def=" + ent.Definement + " | Sys=" + ent.SystemType) : "<null>"));

            SelectedOption = SaveOption.SaveEmail;
            TryPersistLastCategory();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // כפתור: "שמור רק קבצים מצורפים"
        private void btnSaveAttachments_Click(object sender, EventArgs e)
        {
            var ent = this.SelectedEntityInfo;
            DevDiag.Log("Dialog: btnSaveAttachments_Click with entity -> " +
                (ent != null ? (ent.Name + " | Def=" + ent.Definement + " | Sys=" + ent.SystemType) : "<null>"));

            SelectedOption = SaveOption.SaveAttachments; // alias ל-SaveAttachmentsOnly
            TryPersistLastCategory();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // כפתור: "ביטול"
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DevDiag.Log("Dialog: btnCancel_Click");
            SelectedOption = SaveOption.None;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // כפתור: "טעינת הגדרות"
        private void btnLoadSettings_Click(object sender, EventArgs e)
        {
            DevDiag.Log("Dialog: btnLoadSettings_Click ENTER");
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "בחר/י קובץ הגדרות (XML)";
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

                    try { Properties.Settings.Default.Save(); DevDiag.Log("Dialog: btnLoadSettings Settings saved"); } catch { }

                    _config = newCfg;
                    ApplyConfigToUI();
                    RestoreLastCategory();
                    ApplyDefaultEntityToUI();
                    LoadEntitiesFromDb();
                    UpdateFolderLabel();

                    DevDiag.Log("Dialog: btnLoadSettings EXIT OK");
                    MessageBox.Show("ההגדרות נטענו ונשמרו בהצלחה.",
                                    "טעינת הגדרות", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    DevDiag.Log("Dialog: btnLoadSettings EX " + ex.Message);
                    MessageBox.Show("שגיאה בטעינת הקובץ:\r\n" + ex.Message,
                                    "טעינת הגדרות נכשלה",
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
                    ? "לא נבחר נתיב שמירה"
                    : "נתיב שמירה: " + folderPath;

                DevDiag.Log("Dialog: UpdateFolderLabel -> '" + this.lblFolderPath.Text + "'");
            }
            catch (Exception ex)
            {
                DevDiag.Log("Dialog: UpdateFolderLabel EX " + ex.Message);
                if (this.lblFolderPath != null)
                    this.lblFolderPath.Text = "שגיאה בקריאת נתיב השמירה";
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

        private void labelEntity_Click(object sender, EventArgs e)
        {

        }
    }
}
