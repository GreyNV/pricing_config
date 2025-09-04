# Contribution Guidelines

This repository contains a VBA module for managing pricing configuration data in Excel workbooks. The code is stored in `.bas` modules and targets the Windows VBA environment.

## Code Style
- Use `Option Explicit` in every module.
- Group user-editable constants in a `USER CONFIG` section at the top of the file.
- Use `Private Const` for configuration values, `Private`/`Public` scope for subs and functions as appropriate.
- Name subs and functions using `PascalCase` (e.g., `Btn_ClearPricingData`).
- Prefix button entry points with `Btn_` and keep business logic in private helpers.
- Maintain comment banners using the form `' ========= SECTION NAME =========` to delineate major areas.
- Keep mapping definitions in `MAPPING_PAIRS()` using two-element arrays (`Array("A","O"), _`). When adding mappings, preserve order and comment intent.
- For reusable utilities, place them near the end of the file under a `UTILITIES` section.

## Error Handling & Performance
- Wrap external operations in `On Error GoTo EH` blocks and finish with `Finally`/`OptimizeEnd` to restore application settings.
- Use `OptimizeStart`/`OptimizeEnd` when performing workbook-wide operations to disable screen updating and calculation.
- Centralize error messages through `MsgBox` with clear user-facing text.

## Notes and Logging
- Respect the `DISABLE_NOTES` constant. When enabling notes, use `AddCellNote`/`NoteReplace` helpers for consistency.
- Keep helper functions like `IsSkipValue` for standardized checks.

## Feature Additions
- Reuse existing patterns for import/export of pricing data.
- When adding columns, update `MAPPING_PAIRS`, header mapping (`EnsureMappedHeadersFromTool`), and any relevant documentation.
- Ensure new logic handles workbook bounds via `UsedBounds` and column conversions via `ColLetterToNum`.

## Testing
- Automated tests are not yet available. After changes, open the workbook in Excel and manually run `Btn_ClearPricingData` and `Btn_UploadAndProcess` to confirm behavior.
- If automated tests are added in the future (e.g., via `pytest` or VBA unit frameworks), run them before committing.

## Git Workflow
1. Verify the code compiles in the VBA editor (`Debug → Compile`).
2. Run available tests or manual checks.
3. Commit with a concise message prefixed by the area of change (e.g., `feat:`, `fix:`, `refactor:`, `docs:`).

