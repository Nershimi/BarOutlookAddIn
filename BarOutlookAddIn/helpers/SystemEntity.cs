using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using BarOutlookAddIn.Properties;

namespace BarOutlookAddIn.App_Code
{
    /// <summary>
    /// DEPRECATED: legacy SystemEntity helper.
    /// This class used to contain database lookup and write logic. The project now uses
    /// BarOutlookAddIn.Helpers.EntityRepository and ArchiveWriter instead.
    ///
    /// Keep this small shim temporarily to avoid breaking consumers while you remove the legacy file.
    /// After verifying no runtime usage remains, delete this file entirely.
    /// </summary>
    [Obsolete("SystemEntity is deprecated. Use BarOutlookAddIn.Helpers.EntityRepository and ArchiveWriter instead. Remove this file after verification.")]
    internal sealed class SystemEntity
    {
        // Prevent accidental instantiation — fail fast so callers are easy to find.
        public SystemEntity()
        {
            throw new NotSupportedException("SystemEntity is deprecated. Use BarOutlookAddIn.Helpers.EntityRepository and ArchiveWriter instead.");
        }
    }
}



