// allocate numeric attachment name (DB-first, then fallback)
int allocated;
string filePath;
try
{
    allocated = BarOutlookAddIn.Helpers.NumeratorService.GetNextArchiveNumber();
    filePath = Path.Combine(categoryFolder, "ol" + allocated + ext);
    DevDiag.Log($"CtxBtn: allocated numeric attachment name from DB ol{allocated}{ext} (raw='{rawName}')");
}
catch (Exception dbEx)
{
    DevDiag.Log("CtxBtn: NumeratorService failed for attachment, falling back to FileNameAllocator: " + dbEx.Message);
    filePath = FileNameAllocator.AllocatePath(categoryFolder, ext, out allocated);
    DevDiag.Log($"CtxBtn: allocated numeric attachment name FS ol{allocated}{ext} (raw='{rawName}')");
}