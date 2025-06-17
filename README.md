# WYSIWYG PDF, XLSX, and Print Preview for WinForms

A C# WinForms project demonstrating a powerful technique for generating reports with a **guaranteed identical appearance** across Print Preview, exported PDF files, and exported XLSX spreadsheets. What you see in the print preview is *exactly* what you get in the final documents.

The application renders data from multiple `DataGridView` controls into paginated reports, showcasing a real-world use case.


*(A demo GIF showing the application generating a print preview, a PDF, and an XLSX file with identical layouts would be placed here.)*

## ✨ Key Features

*   **True WYSIWYG:** The layout, fonts, and positioning are pixel-perfectly consistent across all outputs.
*   **Multiple Export Formats:**
    *   Print Preview
    *   PDF (`.pdf`)
    *   Excel Spreadsheet (`.xlsx`)
*   **Zero External Dependencies:** PDF and XLSX files are generated **from scratch** without any third-party libraries (like iTextSharp, EPPlus, ClosedXML, etc.).
*   **PDF/A Compliant:** The generated PDFs are built to be compliant with the PDF/A standard for long-term archiving.
*   **Precise Excel Layout:** Replicates the report layout in Excel by calculating column widths, row heights, and embedding printer settings to ensure the print layout in Excel matches the source.

## ⚙️ How It Works

Instead of using separate rendering logic for each output format, this project uses a unified pipeline built around an intermediate representation. This is the key to its consistency.

> **1. Capture Drawing Commands:** A custom `GraphicsRecorder` class intercepts all GDI+ drawing calls (`DrawString`, `DrawRectangle`, etc.) during a "dry run" of the print logic. It stores these calls as a simple, format-agnostic list of commands. This defines the entire report layout *once*.

> **2. Render for Different Targets:**
> *   **Print Preview:** Executes the original GDI+ drawing logic directly to the screen for an immediate preview.
> *   **PDF Export:** A custom PDF writer parses the recorded command list and translates each command into low-level PDF objects and streams.
> *   **XLSX Export:** A custom XLSX writer also parses the same command list, constructing an Office Open XML (`.xlsx`) file from scratch.

This approach ensures that any change to the report's appearance is automatically and consistently reflected in all three output formats.

## 🛠️ Technical Deep Dive

This project is a deep dive into file format specifications and Windows APIs.

### The `GraphicsRecorder`
The heart of the solution. It acts as a "tape recorder" for `System.Drawing` commands, creating a reusable blueprint of the report.

```csharp
public class GraphicsRecorder
{
    public List<string> Commands { get; private set; } = new List<string>();

    public void DrawRectangle(Pen pen, float x, float y, float width, float height)
    {
        // Records the command and its parameters into a simple string format
        Commands.Add(string.Format(CultureInfo.InvariantCulture, "DrawRectangle|{0}|{1}|{2}|{3}|{4}", x, y, width, height, currentPage));
    }

    public void DrawString(string text, Font font, Brush brush, float x, float y)
    {
        // Records text, font info, and position
        Commands.Add(string.Format(CultureInfo.InvariantCulture, "DrawString|{0}|{1}|{2}|{3}|{4}|{5}|{6}", text, font.Name, font.Size, font.Style, x, y + font.Size, currentPage));
    }
    // ...
}
```

### Zero-Dependency PDF Generation
*   Builds the entire PDF object structure (catalog, pages, fonts, streams) manually in C#.
*   Uses **P/Invoke calls to `gdi32.dll`** (`GetFontData`, `GetOutlineTextMetrics`) to extract font metrics and embed the actual font files into the PDF. This ensures text is rendered with the correct font, not a substitute.
*   Compresses font and content streams using the built-in `DeflateStream`.

### Zero-Dependency XLSX Generation
*   Creates a valid `.xlsx` package (which is a ZIP archive) using `System.IO.Compression.ZipArchive`.
*   Generates all necessary XML files (`[Content_Types].xml`, `workbook.xml`, `styles.xml`, worksheet files, etc.) by hand according to the Office Open XML specification.
*   **Crucially**, it replicates the exact page setup (margins, orientation, paper size) by using P/Invoke to extract the printer's **`DEVMODE`** structure and embedding it as a `printerSettings.bin` part within the XLSX file. This is how Excel knows how to lay out the page for printing exactly as intended.

## 🚀 Getting Started

1.  Clone the repository.
2.  Open the `.sln` file in Visual Studio.
3.  Build and run the project (`F5`).
4.  Use the **File** menu to access **Page Setup**, **Print Preview**, **Export to PDF**, and **Export to XLSX**.
