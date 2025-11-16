// allocate numeric name for attachment (DB-first, then fallback)
int allocatedNumber;
string filePath;
try
{
    allocatedNumber = BarOutlookAddIn.Helpers.NumeratorService.GetNextArchiveNumber();
    filePath = Path.Combine(categoryFolder, "ol" + allocatedNumber + ext);
    DevDiag.Log($"Ribbon: allocated numeric attachment name from DB ol{allocatedNumber}{ext} (raw='{rawName}')");
}
catch (Exception dbEx)
{
    DevDiag.Log("Ribbon: NumeratorService failed for attachment, falling back to FileNameAllocator: " + dbEx.Message);
    filePath = FileNameAllocator.AllocatePath(categoryFolder, ext, out allocatedNumber);
    DevDiag.Log($"Ribbon: allocated numeric attachment name FS ol{allocatedNumber}{ext} (raw='{rawName}')");
}