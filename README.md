## C# Unified Rendering: Identical PDF, XLSX & WinForms PrintPreview from Scratch

This project solves a common and frustrating challenge in application development: making exported documents look exactly like the on-screen print preview. This C# WinForms application demonstrates a powerful technique to generate reports with a **guaranteed identical appearance** across 🖼️ **WinForms PrintPreview**, 📄 **Adobe Acrobat PDF**, and 📊 **Microsoft Office Excel XLSX** outputs.

The entire rendering pipeline is **built from scratch with zero third-party libraries**. This project serves as a powerful, educational example of how to take full control over your document generation process by creating a custom rendering engine.

## ✨ Key Features

*   **Consistent Output:** The layout, fonts, and positioning are pixel-perfect and identical across all three output formats. What you see in the **WinForms PrintPreview** is *truly* what you get in the final PDF and XLSX files.
*   **Zero Dependencies:** No Nuget packages, no external libraries. Just pure .NET Framework and P/Invoke calls. The PDF and XLSX writers are built from the ground up.
*   **Unified Codebase:** A single drawing method (`DrawContent`) defines the report's appearance. No need to write separate, format-specific code for PDF, Excel, and printing.
*   **Shared Page Setup:** Page settings (margins, orientation, paper size) from a single `PageSetupDialog` are respected and replicated by all export formats.
*   **Structurally Compliant PDF/A-1a:** The generated PDF is designed for long-term archiving. It achieves this by manually building the required document structure (`StructTreeRoot`), embedding fonts, and including all necessary metadata.
*   **Native XLSX Output:** The generated Excel file is a native Office Open XML (`.xlsx`) document, so it opens without "Compatibility Mode" warnings and perfectly preserves the layout.

## 📸 Screenshots

<table>
  <tr>
    <td align="center">
      <b>Main Application UI</b><br>
      <img src="Screenshots/UI.png" alt="Main Application UI">
    </td>
  </tr>
  <tr>
    <td align="center">
      <b>WinForms PrintPreview</b><br>
      <img src="Screenshots/print-preview.png" alt="WinForms Print Preview">
    </td>
  </tr>
  <tr>
    <td align="center">
      <b>Identical PDF/A-1a Compliant PDF File Output</b><br>
      <img src="Screenshots/pdf-a1a-compliant-output.png" alt="PDF/A-1a Compliant Output">
    </td>
  </tr>
  <tr>
    <td align="center">
      <b>Identical XLSX Native Format Output</b><br>
      <img src="Screenshots/xlsx-native-format-output.png" alt="XLSX Native Format Output">
    </td>
  </tr>
</table>

## 🤔 Core Concept: The "Capture & Replay" Engine

Instead of using different logic for each format, this project uses a unified rendering pipeline built around an intermediate representation of the document.

```
[Your Data] -> [GraphicsRecorder] -> [Replay on Specific Renderer]
```

1.  **⚙️ Capture Phase:** A custom `GraphicsRecorder` class intercepts all GDI+ drawing commands (`DrawString`, `DrawRectangle`, etc.) during a preliminary "dry run" of the print logic. These commands are stored as a simple, serializable list of instructions that acts as a "blueprint" for the report.

    ```csharp
    // Instead of drawing directly to a Graphics object...
    // e.Graphics.DrawString("Hello", ...);

    // We record the command:
    graphicsRecorder.DrawString("Hello", ...);
    // Command is stored as: "DrawString|Hello|Arial|9|Bold|100|150|1"
    ```

2.  **🎨 Replay Phase:** This "blueprint" is then passed to different renderers, each translating the abstract commands into a specific format:
    *   **WinForms PrintPreview:** Replays the commands onto the `PrintPageEventArgs.Graphics` object for a live on-screen preview.
    *   **PDF Exporter:** Parses the command list and manually constructs a PDF document from scratch. It builds all necessary PDF objects (`/Catalog`, `/Pages`, `/Font`), embeds font data (fetched using P/Invoke calls to `gdi32.dll`), and writes the file stream byte-by-byte.
    *   **XLSX Exporter:** Also parses the command list. It creates a `.zip` archive and generates the required XML files (`workbook.xml`, `sheet.xml`, etc.) to build a valid `.xlsx` file. It meticulously translates command coordinates into Excel row heights, column widths, and cell positions to replicate the original layout.

## ✅ Feature Support

This implementation focuses on the core features needed for structured reports. It is intentionally kept simple to serve as a clear example.

| Feature                 | 📄 PDF         | 📊 XLSX        | 🖼️ WinForms PrintPreview | Status                                             |
| ----------------------- | :------------: | :------------: | :---------------------: | -------------------------------------------------- |
| **Text**                |       ✅       |       ✅       |           ✅            | Fully supported.                                   |
| **Lines & Rectangles**  |       ✅       |       ✅       |           ✅            | Used for creating tables.                          |
| **Identical Page Setup**|       ✅       |       ✅       |           ✅            | Margins, paper size, and orientation are identical. |
| **PDF/A-1a Compliance** |       ✅       |       --       |           --            | Structural & metadata compliance is implemented.   |
| **Images**              |       ❌       |       ❌       |           ❌            | Not implemented.                                   |

## 💡 Philosophy & Target Audience

This project is for you if:
*   You are a **student or developer** who wants to learn how file formats like PDF and XLSX are structured at a binary level.
*   You want to understand how to bridge the gap between **high-level GDI+ drawing and low-level file formats**.
*   You need **absolute, pixel-perfect control** over a simple report's output and find existing libraries too abstract.
*   You are working in a legacy or restricted environment where **adding external dependencies is not an option**.

While powerful libraries like **QuestPDF**, **iText**, and **ClosedXML** are the right choice for most production applications, this project offers a lightweight, transparent, and educational alternative.

## 🚀 Future Roadmap & Potential Improvements

This project provides a solid foundation, but there are many exciting ways it can be extended. Here are some of the planned and potential features for the future:

#### 🎨 Core Drawing Features
*   **Implement Image Support:** Add a `DrawImage` command to the `GraphicsRecorder` and implement the logic to embed image data (JPEG, PNG) into both PDF and XLSX files and embedding necessary color profiles to generated PDF files.
*   **Support for Additional Shapes:** Extend the system to handle other GDI+ primitives like `DrawEllipse`, `DrawArc`, and `DrawPolygon` for more complex diagrams and charts.
*   **Improve Text Handling:** Add support for more advanced text rendering, such as text alignment (center/right), rotated text, and styles like underline/strikethrough.

#### ⚡ Performance & Optimization
*   **Introduce Font Subsetting:** A high-priority task. Currently, the entire font file is embedded in the PDF, leading to larger file sizes. Implementing font subsetting (embedding only the characters actually used) will drastically reduce the PDF file size while maintaining compliance.
*   **Background Processing:** For large reports, the export process can freeze the UI. Refactor the export logic to run on a background thread (`async/await`) with progress updates to keep the application responsive.

#### 🏗️ Code Architecture & Quality
*   **Modularize the Code:** Refactor the core `GraphicsRecorder`, PDF writer, and XLSX writer into a separate, reusable class library. This would allow other projects to easily consume this functionality.
*   **Expand Unit Tests:** Create a comprehensive test suite that verifies the output for a given set of intermediate commands, ensuring that bug fixes or new features don't cause regressions.

**Contributions are welcome!** If any of these features interest you, feel free to fork the repository and submit a pull request.

## 🛠️ How to Run

1.  Clone the repository: `git clone https://github.com/samialtas/CSharp-PDF-XLSX-PrintPreview-Export-and-Generation.git`
2.  Open the solution file (`.sln`) in Visual Studio.
3.  Build and run the project (F5).
