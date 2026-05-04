![NanoXLSX](NanoXLSX.png)

# NanoXLSX Demo

Comprehensive demo applications showcasing **NanoXLSX v3.x** and **PicoXLSX v4.x** functionality with practical, ready-to-run examples.

## 📋 Overview

This repository contains executable use cases demonstrating the core features of NanoXLSX and PicoXLSX, .NET libraries for creating (and reading) Microsoft Excel (XLSX) files. Each use case is a standalone example that generates or reads Excel files, making it easy to learn and understand the libraries' capabilities.

**Current Implementation:**

- ✅ **NanoXLSX** (.NET 8.0) - Fully implemented with 28 use cases
- ✅ **PicoXLSX** (.NET 8.0) - Fully implemented with 25 use cases (writer-only)
- ⏳ **NanoXLSX4j** (Java >11) - Planned

## 🚀 Getting Started

### Prerequisites

- .NET 8.0 SDK or .NET 9.0 SDK (for .NET 8.0 target)
- .NET Framework 4.5 or higher (for .NET Framework 4.5 target)
- Any IDE supporting .NET (Visual Studio, VS Code, Rider)

### Running the NanoXLSX Demos

**Interactive Mode** (shows menu of all use cases):

```bash
cd NanoXLSX/Demo
dotnet run
```

To run a specific framework:

```bash
dotnet run --framework net8.0    # Run with .NET 8.0
dotnet run --framework net45      # Run with .NET Framework 4.5
```

**Run All Use Cases**:

```bash
dotnet run all
```

**Run Specific Use Cases** (by name or number):

```bash
dotnet run "BasicDemo,Read,StyleBasics"    # By name
dotnet run "1,2,12"                        # By number
```

### Running the PicoXLSX Demos

```bash
cd PicoXLSX/Demo
dotnet run                                      # Interactive mode
dotnet run all                                   # Run all use cases
dotnet run "BasicDemo,StyleBasics,Formulas"      # By name
dotnet run "1,2,10"                              # By number
```

## 📚 Available Use Cases

### Basic Operations

| # | Use Case | Description |
|---|----------|-------------|
| 1 | **BasicDemo** | Create a basic workbook with text, number, and datetime values |
| 2 | **Read** | Load and read an existing Excel file with cell information |
| 9 | **DataTypes** | Working with various data types (strings, numbers, dates, booleans) |
| 10 | **CellDirectionsAndValues** | Cell navigation and value handling |

### Advanced Cell Operations

| # | Use Case | Description |
|---|----------|-------------|
| 4 | **CellAndWorksheetSelection** | Usage of cell and worksheet selection |
| 8 | **AddNextCell** | Using AddNextCell for sequential cell addition |
| 15 | **CellRanges** | Working with cell ranges and range operations |
| 17 | **MergeCells** | Merging cells for headers and layout |

### Styling & Formatting

| # | Use Case | Description |
|---|----------|-------------|
| 12 | **StyleBasics** | Basic usage of styles (fonts, fills, alignment) |
| 13 | **StyleAppending** | Style appending and method chaining |
| 14 | **ActiveAndSetStyle** | Applying styles to individual cells and ranges |

### Inline Formatting (Rich Text)

| # | Use Case | Description |
|---|----------|-------------|
| 25 | **InlineFormattingBasics** | Multiple text runs in a single cell with line breaks |
| 26 | **InlineFormattingStyles** | Styled inline formatting (bold, italic, colors, sizes) |
| 27 | **InlineFormattingPhonetic** | Phonetic runs for East Asian languages (Japanese) |
| 28 | **InlineFormattingRead** | Reading cells with inline formatting from saved files |

### Layout & Dimensions

| # | Use Case | Description |
|---|----------|-------------|
| 11 | **ColumnWidthsRowHeights** | Setting column widths and row heights |
| 19 | **HidingRowsAndColumns** | Hiding rows and columns |
| 23 | **PaneSplitAndFreeze** | Worksheet splitting and freezing panes |
| 24 | **HidingWorkbooksAndWorksheets** | Hiding workbooks and worksheets from visibility |

### Data Management

| # | Use Case | Description |
|---|----------|-------------|
| 20 | **AutoFilter** | Auto-filtering for data tables |
| 22 | **Formulas** | Excel formulas usage (SUM, AVERAGE, VLOOKUP, etc.) |
| 16 | **Metadata** | Assigning workbook metadata (title, subject, author) |
| 21 | **SanitizingWorksheetNames** | Worksheet name sanitization and validation |

### Security & Protection

| # | Use Case | Description |
|---|----------|-------------|
| 18 | **ProtectionAndPasswords** | Sheet protection and password protection |

### I/O Operations

| # | Use Case      | Description                                        |
|---|---------------|----------------------------------------------------|
| 5 | **Stream**    | Save workbooks to memory streams and file streams  |
| 6 | **SaveAsync** | Asynchronous file saving for better performance    |
| 7 | **ReadAsync** | Asynchronous workbook loading from file and stream |

### Syntax Shortcuts

| # | Use Case | Description |
|---|----------|-------------|
| 3 | **Shortener** | Demonstrate shortened syntax for writing cells |

## 🏗️ Project Structure

```
Demo/
├── NanoXLSX/
│   ├── Demo/                   # Multi-targeting project (net8.0 + net45)
│   │   ├── Program.cs          # Interactive demo runner
│   │   ├── UseCases/           # 28 individual use case files
│   │   └── NanoXLSX.Demo.csproj  # Project file
│   └── Demo.sln                # Visual Studio solution
├── PicoXLSX/
│   ├── Demo/                   # Multi-targeting project (net8.0 + net45)
│   │   ├── Program.cs          # Interactive demo runner
│   │   ├── UseCases/           # 25 individual use case files (writer-only)
│   │   └── PicoXLSX.Demo.csproj  # Project file
│   └── Demo.sln                # Visual Studio solution
├── NanoXLSX4j/                 # (Planned)
└── global.json                 # .NET SDK version configuration
```

## 📖 Documentation & Resources

### NanoXLSX

- **Main Repository**: [github.com/rabanti-github/NanoXLSX](https://github.com/rabanti-github/NanoXLSX)
- **API Documentation**: [rabanti-github.github.io/NanoXLSX](https://rabanti-github.github.io/NanoXLSX/)
- **NuGet Package**: [nuget.org/packages/NanoXLSX](https://www.nuget.org/packages/NanoXLSX)
- **Current Demo Version**: NanoXLSX v3.0.0-rc.9

### PicoXLSX

- **Main Repository**: [github.com/rabanti-github/PicoXLSX](https://github.com/rabanti-github/PicoXLSX)
- **API Documentation**: [rabanti-github.github.io/PicoXLSX](https://rabanti-github.github.io/PicoXLSX/)
- **NuGet Package**: [nuget.org/packages/PicoXLSX](https://www.nuget.org/packages/PicoXLSX)
- **Current Demo Version**: PicoXLSX v4.0.0-rc.9
- **Migration Guide**: [PicoXLSX v3.x to v4.0.0](https://github.com/rabanti-github/PicoXLSX/blob/master/MigrationGuide.md)

### NanoXLSX4j (Java)

- **Main Repository**: [github.com/rabanti-github/NanoXLSX4j](https://github.com/rabanti-github/NanoXLSX4j)
- **Javadoc**: [rabanti-github.github.io/NanoXLSX4j](https://rabanti-github.github.io/NanoXLSX4j/)

## 💡 Example Usage

Here's a quick example from the `BasicDemo` use case:

```csharp
using NanoXLSX;

// Create a new workbook
Workbook workbook = new Workbook("BasicDemo.xlsx", "Sheet1");

// Add cells with different data types
workbook.CurrentWorksheet.AddNextCell("Hello World");    // String
workbook.CurrentWorksheet.AddNextCell(42);               // Number
workbook.CurrentWorksheet.AddNextCell(DateTime.Now);     // Date

// Save the workbook
workbook.Save();
```

For reading Excel files:

```csharp
using NanoXLSX;
using NanoXLSX.Extensions;

// Load an existing workbook
Workbook workbook = WorkbookReader.Load("BasicDemo.xlsx");

// Access cells
foreach (var cell in workbook.CurrentWorksheet.Cells)
{
    Console.WriteLine($"Cell {cell.Key}: {cell.Value.Value}");
}
```

## 🎯 Multi-Targeting Support

The demo projects target both:

- **.NET Framework 4.5** - For legacy Windows applications
- **.NET 8.0** - For modern cross-platform applications

This allows you to see how NanoXLSX and PicoXLSX work across different .NET implementations using a single codebase.

## 🔄 Migration Guides

### NanoXLSX v2.x to v3.0.0

All NanoXLSX demos have been migrated to v3.0.0. Key changes include:

- `Workbook.Load()` → `WorkbookReader.Load()` (requires `using NanoXLSX.Extensions;`)
- `SetSelectedCells()` → `ClearSelectedCells()` + `AddSelectedCells()`
- Enum values now use PascalCase (e.g., `fillColor` → `FillColor`)

For complete migration details, see the [NanoXLSX Migration Guide](https://github.com/rabanti-github/NanoXLSX/blob/master/MigrationGuide.md).

### PicoXLSX v3.x to v4.0.0

All PicoXLSX demos have been migrated to v4.0.0. Key changes include:

- Namespace changed from `PicoXLSX` to `NanoXLSX`
- Sub-classes like `Cell.Address`, `Style.Font` are now independent classes (`Address`, `Font`)
- All constants and enum values renamed to PascalCase
- PicoXLSX can be extended with NanoXLSX packages (e.g., add `NanoXLSX.Reader` for read support)

For complete migration details, see the [PicoXLSX Migration Guide](https://github.com/rabanti-github/PicoXLSX/blob/master/MigrationGuide.md).

## 📜 License

This demo project follows the same license as the main NanoXLSX library - MIT License.

## 🤝 Contributing

This is a demo repository. For contributions to the main library, please visit:

- [NanoXLSX Issues](https://github.com/rabanti-github/NanoXLSX/issues)
- [PicoXLSX Issues](https://github.com/rabanti-github/PicoXLSX/issues)
- [NanoXLSX4j Issues](https://github.com/rabanti-github/NanoXLSX4j/issues)

---

**Note**: For library requirements, roadmap, and detailed feature documentation, please refer to the main repository READMEs linked above.
