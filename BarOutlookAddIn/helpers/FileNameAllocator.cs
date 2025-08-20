// Helpers/FileNameAllocator.cs
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace BarOutlookAddIn.Helpers
{
    public static class FileNameAllocator
    {
        // FS-only allocator: scans folder for "ol{N}.*" and picks next number
        public static string AllocatePath(string folder, string ext, out int number)
        {
            DevDiag.Log($"FNAlloc: start folder='{folder}', ext='{ext}'");
            number = 0;

            try { Directory.CreateDirectory(folder); }
            catch (Exception ex)
            {
                DevDiag.Log("FNAlloc: CreateDirectory EX: " + ex.Message);
                throw;
            }

            // ext normalize
            if (string.IsNullOrWhiteSpace(ext)) ext = ".bin";
            if (!ext.StartsWith(".")) ext = "." + ext;

            // pattern on the *file name without extension*: ^ol{digits}$
            var re = new Regex(@"^ol(?<n>\d+)$", RegexOptions.IgnoreCase);
            int max = 0;

            try
            {
                foreach (var f in Directory.GetFiles(folder))
                {
                    var nameNoExt = Path.GetFileNameWithoutExtension(f);
                    var m = re.Match(nameNoExt);
                    if (m.Success)
                    {
                        int n;
                        if (int.TryParse(m.Groups["n"].Value, out n))
                            if (n > max) max = n;
                    }
                }
            }
            catch (Exception ex)
            {
                DevDiag.Log("FNAlloc: scan EX: " + ex.Message);
                // not fatal – continue with max=0
            }

            number = max + 1;
            string file = Path.Combine(folder, "ol" + number + ext);
            DevDiag.Log($"FNAlloc: result number={number}, path='{file}'");
            return file;
        }
    }
}
