WYSIWYG PDF, XLSX, and Print Preview Export for WinForms
This C# WinForms project demonstrates a powerful technique for generating reports with a guaranteed identical appearance across Print Preview, exported PDF files, and exported XLSX spreadsheets. What you see in the print preview is exactly what you get in the PDF and Excel files, down to the layout, fonts, and positioning.

The application renders data from multiple DataGridView controls into paginated reports.

How It Works
Instead of using separate logic for each output format, this project uses a unified rendering pipeline:

Intermediate Representation: A custom GraphicsRecorder class captures all GDI+ drawing commands (DrawString, DrawRectangle, etc.) into a simple, format-agnostic list of commands. This step defines the entire report layout once.

Print Preview: Renders the report directly to the screen by executing the drawing logic in real-time.

PDF Export: Interprets the recorded command list and translates each command into low-level PDF objects and streams. This module builds a PDF/A-compliant document from scratch without any external libraries, including manual font embedding via P/Invoke and gdi32.dll to get precise font metrics.

XLSX Export: Also interprets the same command list. It constructs an .xlsx file (Office Open XML) from scratch using ZipArchive. It translates the coordinates and sizes of drawing commands into corresponding Excel row heights, column widths, and merged cells to replicate the visual layout perfectly. It even preserves page setup details (margins, orientation, paper size) by embedding the printer's DEVMODE structure into the file.
