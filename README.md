# BarOutlookAddIn — Project README

Summary
-------
BarOutlookAddIn is an Outlook VSTO add-in that lets users save emails and attachments into a configured archive folder and optionally register those files in a backend (DB). Core responsibilities:
- Provide a Ribbon button and context-menu action for saving mail/attachments.
- Present a dialog to choose entity, category and filename options.
- Allocate stable numeric file names (ol{N}.*) using a DB numerator with a filesystem fallback.
- Write files to disk and (optionally) insert archive metadata into a DB stored-procedure.

Target platform: C# 7.3, .NET Framework 4.7.2.

Files and responsibilities
--------------------------

- README.md
  - This file.

- Properties/AssemblyInfo.cs
  - Assembly metadata (Title, Company, Version, GUID, COM visibility). No runtime logic.

- ThisAddIn.cs
  - Add-in lifecycle (Startup/Shutdown).
  - Configures logging early (DevDiag.ConfigureLogFolder).
  - Provides CreateRibbonExtensibilityObject to register the Ribbon implementation (AttachmentContextMenuRibbon).
  - Reads/persists basic connection settings from __ConfigPath__ and project __Settings__.
  - Helpers to build a DB connection string from XML and to inspect stored-proc signature (DevDiag + DbInspector calls).

- AttachmentContextMenuRibbon.cs
  - Implements Office.IRibbonExtensibility for the custom Ribbon XML used for the context menu and Home tab button.
  - Supplies fallback inline Ribbon XML (GetInlineXmlFallback) and provides LoadImage for custom images (from Resources).
  - Handles context-menu save action (OnSaveAttachmentToArchive) and the Home tab button action (OnHomeSaveButton).
  - Performs attachment selection, inline-attachment filtering (IsInlineImage / GetContentId), file-name allocation logic (NumeratorService → FileNameAllocator → GetUniquePath), SaveAs calls and DB insert via ArchiveWriter.
  - Utility helpers: ResolveArchiveRoot, EnsureCategoryFolder, EnsureWritableFolder, GetUniquePath, CleanFileName, TrySaveMsgWithDiag, and safe COM release patterns.
  - Important: protects against COMExceptions and logs details via DevDiag.

- SaveEmailRibbon.cs
  - Designer-backed Ribbon class (the visual Ribbon created with the designer).
  - Handles designer button click (btnSaveSelectedEmail_Click) for saving the active/selected MailItem.
  - Mirrors the same allocation / save / DB-insert flow used by AttachmentContextMenuRibbon (NumeratorService → FileNameAllocator → GetUniquePath → mailItem.SaveAs / att.SaveAsFile).
  - Provides helper functions duplicated with AttachmentContextMenuRibbon: ResolveArchiveRoot, EnsureCategoryFolder, EnsureWritableFolder, GetUniquePath, CleanFileName, IsInlineImage, GetContentId, TrySaveMsgWithDiag.
  - Logs actions and surfaces user-facing MessageBoxes on errors/success.

- SaveEmailDialog.cs
  - Modal WinForms dialog used to choose SaveOption (SaveEmail, SaveAttachments, SaveAttachmentsOnly), category, entity and optional custom filename.
  - Loads configuration from the file at __ConfigPath__ (via Helpers.AddInConfig.Load) and persists selected __LastCategory__ to Settings.
  - Loads entities from DB via EntityRepository, supports preselected attachment indices (for context-menu flow), validates custom filename input and builds/persists ConnectionString when XML contains SQL settings.
  - Exposes public properties required by ribbons: SelectedOption, SelectedEntityInfo, SelectedEntityName, SelectedCategory, RequestNumber, UseCustomFileName, CustomFileName.

- helpers/SaveEngine.cs
  - Centralized programmatic save logic (internal static class) that can be reused by other callers.
  - Implements SaveWholeMail and SaveAttachmentsOnly with the same allocation strategy (NumeratorService → FileNameAllocator → GetUniquePath).
  - Uses ArchiveWriter to insert DB records and handles inline detection and sanitization.
  - Provides helper utilities parallel to the UI code: ResolveArchiveRoot, EnsureCategoryFolder, GetUniquePath, Sanitize and IsInline.

- helpers/FileNameAllocator.cs
  - Simple filesystem-only allocator that scans a folder for files named "ol{N}.*" and returns the next number and candidate path.
  - Used as a fallback when the DB numerator (NumeratorService) is unavailable.

Referenced components (helpers / DB / resources)
-----------------------------------------------
(The files for these components may be in the project but not shown in this snapshot. Descriptions explain how they integrate.)

- Helpers.AddInConfig
  - Loads XML configuration (ArchivePath, categories, default entity, SQL settings). Used by SaveEmailDialog and ThisAddIn.

- ArchiveWriter
  - Responsible for inserting archive metadata into the backend (stored-proc SP_Insert_Archive).
  - API used: TryInsertRecord(EntityInfo ent, string dspEntityNum, string fullPath, string fileDesc) or overload with entityName.

- EntityRepository / EntityInfo
  - EntityRepository obtains list of entities from DB for the dialog combo.
  - EntityInfo is a simple model with Name, Definement, SystemType and DisplayText.

- NumeratorService
  - Returns next numeric archive number (atomic, DB-backed). First choice for generating ol{N} names.

- DevDiag
  - Project logging helper. This project configures DevDiag early in ThisAddIn_Startup; many code paths call DevDiag.Log for diagnostics.
  - Default log folder is configured in ThisAddIn (example path in code: C:\bar\m9).

- Helpers.DbInspector
  - Utility used from ThisAddIn to check existence and parameters of stored procedures.

- Project Resources (Properties.Resources)
  - Ribbon images are loaded with ResourceManager.GetObject(imageId) — supply an icon named (e.g.) BarAddin_icon in Resources to get the custom Ribbon image.

Important settings (project Settings)
------------------------------------
These keys are read / written by the add-in. They live under the project __Settings__ (Properties → Settings):
- __SaveBaseFolder__ — preferred archive folder path (overrides XML).
- __ConfigPath__ — path to the XML configuration file the dialog can load.
- __ConnectionString__ — persisted SQL connection string (built from XML when available).
- __LastCategory__ — the last-selected category persisted by the dialog.

Runtime flow (high level)
-------------------------
1. ThisAddIn_Startup — configure logging, optionally show DB toast.
2. Ribbon(s) load:
   - Designer Ribbon (SaveEmailRibbon) for the Home tab
   - AttachmentContextMenuRibbon via CreateRibbonExtensibilityObject for context menu and custom image
3. User triggers save (Home button or context-menu):
   - Acquire MailItem (Inspector or Explorer selection).
   - Show SaveEmailDialog to pick entity/category/filename option.
   - Resolve archive root (Settings __SaveBaseFolder__ → config XML ArchivePath → default Documents\SavedMails).
   - Ensure folder is writable (creates a probe temp file).
   - Allocate filename: call NumeratorService.GetNextArchiveNumber() (DB) → fallback to FileNameAllocator → fallback to humanized GetUniquePath.
   - Save file(s) using Outlook API: mailItem.SaveAs or attachment.SaveAsFile.
   - Insert DB record via ArchiveWriter.TryInsertRecord (if entity info or name provided).
   - Log and show MessageBox about success/failure.

Troubleshooting / common issues
-------------------------------
- If Save fails with COM exceptions, logs capture HResult and message (TrySaveMsgWithDiag and COM catch blocks).
- If images do not appear on the custom Ribbon, make sure a resource named matching image id (e.g., BarAddin_icon) is present in __Resources__ and the AddIn returns embedded XML or uses GetCustomUI correctly.
- If numeric allocation fails, code falls back to filesystem allocator (FileNameAllocator) and then unique name generation; examine logs for NumeratorService errors.
- If DB calls do not work, ensure __ConnectionString__ or XML SQL settings are correct and DevDiag logs the connection attempts (ThisAddIn TryBuildConnectionStringFromXml / SaveEmailDialog btnLoadSettings).

Extending the add-in
--------------------
- Add/modify Ribbon UI: edit the designer (SaveEmailRibbon) or change GetInlineXmlFallback in AttachmentContextMenuRibbon.
- Change allocation strategy: adjust NumeratorService usage or modify helpers/FileNameAllocator.cs.
- Add new metadata to DB insert: extend ArchiveWriter.TryInsertRecord and dialog to collect extra fields.
- Add resources: put icons in project Properties → Resources and reference by name in Ribbon XML / LoadImage.

Notes for maintenance
---------------------
- Many code paths interact with COM objects (Outlook). Always follow existing pattern: release COM objects via Marshal.ReleaseComObject in finally blocks to avoid leaks.
- Logging is centralized via DevDiag — use it to capture issues in production.
- Key user-editable configuration is the __ConfigPath__ XML and the __SaveBaseFolder__ project setting.

If you want, I can:
- generate a shorter quick-start README for non-developers,
- add a diagram of runtime flow,
- or create a checklist for testing save/DB scenarios.
