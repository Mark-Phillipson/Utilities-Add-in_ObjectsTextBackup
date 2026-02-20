# Utilities Add-in (Access Object Text Backup)

This repository is a **text backup/export** of a Microsoft Access utilities add-in database.

It stores Access objects as text files so they can be versioned in Git:
- Forms: `*.form`
- Modules: `*.bas`
- Macros: `*.macro`
- Reports: `*.report`
- Query SQL: `*.SQL`

## What this repository is

- A source-control-friendly snapshot of Access objects exported from an add-in database.
- A working area for reviewing and editing VBA/form/macro logic as text.
- A deployment source for importing objects into other Access databases.

## What this repository is not

- Not a conventional compiled app (no standard build/test pipeline).
- Not directly runnable from VS Code alone; validation happens inside Access.

## How to use it

### 1) Review or edit objects as text

- Edit VBA in `*.bas` files.
- Edit form/report code-behind in `*.form`/`*.report` files (look below `CodeBehindForm` for form code).
- Avoid hand-editing Access-generated `_AXL:` blocks in `*.macro` unless necessary.

### 2) Load objects into Access

Use Access import flows already present in the project, for example:
- `FormInstall.form` for installing/exporting selected objects via `DoCmd.TransferDatabase`.
- `basAddInProcedures.GetLatestExtrasMacro()` for refreshing `ExtrasRibbon` and `AutoKeys` from the add-in file.

### 3) Run and validate in Access

Primary runtime entry points:
- `ExtrasRibbon.macro`
- `Autokeys.macro`
- `frmMain.form` switchboard actions

Validate behavior by opening the Access database and running these actions directly.

### 4) Export updated objects back to text

Use `BackupDatabaseObjects.exportAllObjectsAsText()` to write forms/modules/macros/reports/query SQL back to disk.

## Key implementation patterns

- DAO is the default for local Access object/query operations (`CurrentDb`, `DAO.Recordset`, `DAO.QueryDef`).
- ADODB is used for external database operations (notably SQL Server linking/metadata work).
- Connection/session state is frequently passed via `Application.TempVars`.
- Many workflows depend on stable object names referenced in `DoCmd.OpenForm`, `DoCmd.RunMacro`, and `Application.Run` calls.

## Contributing guidance

- Prefer small, surgical edits in existing modules/forms used by current macros.
- Preserve the existing error-handler style (`ExitHere` / `HandleErr` blocks).
- Keep object names in sync across forms, macros, and switchboard references when renaming.
