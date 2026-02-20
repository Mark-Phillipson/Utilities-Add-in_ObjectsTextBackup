# Copilot Instructions for Utilities Add-in (Access Object Text Backup)

## Project shape and architecture
- This repo is a text export of Access objects (`*.form`, `*.bas`, `*.macro`, `*.report`, `*.SQL`), not a conventional compiled app.
- Runtime entry points are macro/table driven: `ExtrasRibbon.macro`, `Autokeys.macro`, `basAddInProcedures.bas`, and switchboard dispatch in `frmMain.form` via `SwitchboardItems_AddIn`.

## Core workflows
- Backup/export: `BackupDatabaseObjects.exportAllObjectsAsText()` writes all objects/query SQL to disk.
- Install/deploy: `FormInstall.form` (`InstallObjects`, `RefreshObjects`, `ResetInstall`) stages through `temporaryObjectList` + `subformObjectList.form` and uses `DoCmd.TransferDatabase`.
- Macro refresh: `basAddInProcedures.GetLatestExtrasMacro()` imports `ExtrasRibbon` and `AutoKeys` from `C:\MSOffice\access\Utilities Add-in.accda`.
- Object text round-trip helpers live in `basAddInProcedures.bas` (`SaveFormAsText`, `LoadFormFromText`, `SaveReportAsText`, `LoadReportFromText`).

## Data access and integration patterns
- DAO is primary for local Access objects (`CurrentDb`, `DAO.Recordset`, `DAO.QueryDef`).
- ADODB is used for external server operations (`SQLServer.bas`, `Create SQL Server Table.form`, `frmManageLinkedTables.form`).
- SQL Server/ODBC connection state is often passed through `Application.TempVars` (e.g. `SQLServerConnectionString`, `ODBCConnectionString`, `CurrentDefaultSchema`).
- Linked-table relinking/metadata updates are concentrated in `basLinkedTables.bas`, `SQLServer.bas`, and `frmManageLinkedTables.form`.

## Conventions
- Keep `Option Compare Database` + `Option Explicit` in VBA modules.
- Preserve naming style: modules/forms `bas*`, `frm*`, `sfm*`; temp objects/tables `tmp*`, `temporary*`; variables `string*`, `str*`, `int*`, `long*`, `boolean*`, `db*`, `rs*`.
- Preserve established error-handler structure (`ExitHere` / `HandleErr`, `Select Case Err.Number`, message-box diagnostics).
- Prefer modifying procedures already wired by macros/forms instead of adding disconnected entry points.
- If a macro `RunCode` target changes, update both the macro argument and the VBA procedure name.

## Editing Access text exports safely
- In `*.form`, code-behind starts at `CodeBehindForm`; design metadata above it is verbose and fragile.
- In `*.macro`, `_AXL:` XML comment fragments are Access-generated; do not hand-edit unless absolutely necessary.
- Keep object names stable: many routines open forms/macros by literal name (`DoCmd.OpenForm "..."`, `DoCmd.RunMacro`, `Application.Run`).
- Keep related artifacts synchronized when changing UI/object names (macros, switchboard records, form references).

## Build/test/debug
- There is no reliable CLI compile/test pipeline in this export repo.
- Validate in Access by running `ExtrasRibbon` actions, `Autokeys` shortcuts, and `frmMain` switchboard options.
- VS Code `dotnet build` is not representative of Access VBA runtime behavior.

## Practical guidance for AI agents
- Make small, surgical edits in existing modules/forms already used by macros.
- Prefer discoverable existing helpers (`codeDB`, `IsLoaded`, `GetObjectName`, Transfer/Load/Save helpers) before introducing new utilities.
- For backup/install features, maintain the existing object-text round-trip design rather than introducing external packaging/build tools.
