namespace BarOutlookAddIn.Helpers
{
    // Minimal model representing an entity row from DB.
    public class EntityInfo
    {
        public string Name;        // entity_name / Description_Code
        public int Definement;  // entity_type / definement_entity_type (int)
        public string SystemType;  // system_entity_type / Code_Identification ("ת", "ב", "פ", ...)

        // Display text for ComboBox (name + type description)
        public string DisplayText
        {
            get
            {
                var d = EntityTypeCatalog.GetDescription(SystemType);
                return string.IsNullOrEmpty(d) ? (Name ?? "") : (Name + " (" + d + ")");
            }
        }

        public override string ToString() { return Name ?? ""; }
    }
}
