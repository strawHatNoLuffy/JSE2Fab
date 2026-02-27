# WindowsFormTool (SymbolicLink)

## Summary
Windows desktop application for processing TSK wafer map files. The app focuses on
batch merge and INK (re-bin) operations for TSK data, with supporting converters
and utilities for related mapping formats.

## Product scope
- Primary use cases
  - Merge: copy Fail die from a template TSK into a target TSK.
  - INK: apply INK rules to update bin codes in a TSK file.
  - Batch processing over folders of TSK files.
- Supporting utilities (library-level)
  - Convert TSK to TMA.
  - Convert TSK to CMD TXT.
  - Restore/update TSK data from Excel tables.

## Tech stack
- Language/runtime: C# on .NET Framework 4.7.2
- UI: Windows Forms
- Build: MSBuild / Visual Studio; output type WinExe
- Dependencies
  - MiniExcel 1.32.0 (NuGet, packages.config)
  - Microsoft.Office.Core COM reference
- Target platform: Windows (AnyCPU)

## Repository layout
- `Form1.cs`, `Form1.Designer.cs`, `Form1.resx`: main WinForms UI
- `Program.cs`: app entry point (single-instance guard, optional test runner)
- `TskUtil/`: processing layer (merge, INK, data processing, file helpers)
- `TskUtil/strategy/`: processor strategies (ITskProcessor, TskMergeProcessor, TskInkProcessor)
- `TskUtil/InkRules/`: INK rule definitions and manager
- `File/`: mapping file formats and converters (TSK, TMA, DAT, etc.)
- `Util/`: converters and helper utilities
- `Tests/`: manual test harness (INK rule tests)

## Conventions and patterns
- Naming/style
  - Public types and methods: PascalCase
  - Locals: camelCase
  - Private fields: leading underscore (for new code)
- Namespaces
  - App and mapping code: `DataToExcel`
  - TSK utilities: `WindowsFormTool.TskUtil`
- UI conventions
  - WinForms partial classes; avoid manual edits in `*.Designer.cs`
  - User-facing status goes to RichTextBox; errors via MessageBox
  - UI text is primarily Chinese; keep language consistent
- Processing architecture
  - `TskProcessor` dispatches to `ITskProcessor` strategies by operation type
  - INK rules implement `IInkRule` and are registered via `InkRuleManager`

## Domain notes (TSK)
- Spec compliance: follows `TSK_spec_2013.pdf`
  - Spec location: `D:\HT\JSE2Fab\WindowsFormTool-HT\docs\TSK_spec_2013.pdf`
- Encoding/endianness
  - TSK files are binary; strings are ASCII
  - Numeric fields are big-endian; conversion uses byte-order reverse helpers
- File structure (high level)
  - Header (~200 bytes)
  - Die data: `Rows × Cols × 6 bytes`
  - Optional: ExtendHeadFlag (172 bytes)
  - Optional: ExtendFlag die data `Rows × Cols × 4 bytes`
  - Optional: ExtendFlag2 trailing bytes (keep for round-trip integrity)
- Version compatibility
  - MapVersion 2/4/7 supported
  - ExtendFlag byte order differs by MapVersion (2 normal, 4/7 reversed)
- Coordinates and limits
  - X/Y range 0–511 with sign bits for negatives
  - X/Y growth directions: X (1=left-negative, 2=right-positive), Y (1=front-positive, 2=back-negative)
  - Base Bin range 0–63; extended Bin range 0–255
  - PassDie Bin usually 1; FailDie Bin > 1
- Merge semantics
  - Only merge same mapping type
  - Source FailDie overrides target, then recompute Pass/Fail/Total

## Key data structures
- `DieCategory` enum and `DieData` model for die attributes/bin/site/coords
- `DieMatrix` provides indexing, rotation, offset, expand/collapse, paint, stats
- `ConvertConfig` loads XML field mappings and rotation/trim settings for converters

## Data flow (high level)
- Read: `Tsk.Read()` → header → die matrix → optional extension data
- Merge: `Tsk.Merge()` → overlay FailDie → recompute stats → save new TSK
- Convert: `CMDTskToTxt` / `TskToTma` via `ConvertConfig` mappings

## Configuration
- Default output folder: `D:\New-Tsk\` (see `TskUtil/TskFileHelper.cs`)
- Save path can be overridden by the UI input

## Testing
- `Tests/InkRuleTests.cs` provides a manual test harness
- `Program.RunTests()` can be used to run tests (currently not wired by default)
