# Security Policy

## Supported Versions

The following versions of the `C# Unified Rendering: Identical PDF, XLSX & WinForms PrintPreview from Scratch` application are currently supported with security updates:

| Version | Supported          |
|---------|--------------------|
| Pre-Release     | ✅                 |

## Security Considerations

This application is built using .NET Framework 4.8 and includes functionalities for generating PDF files and exporting data to XLSX format. Below are key security considerations for users and contributors:

### Native Windows API Calls
- The application uses Windows API functions (`gdi32.dll`, `winspool.Drv`, `kernel32.dll`) for font handling and printer settings. These calls are performed using `DllImport` with proper error handling to prevent crashes or undefined behavior.
- Memory management for native resources (e.g., `GlobalLock`, `GlobalUnlock`, `GlobalFree`) is implemented to avoid memory leaks or unauthorized access.
- Users should ensure the application runs in a trusted environment, as native API calls may interact with system-level resources.

### File Handling
- The application writes to PDF and XLSX files using user-specified paths via `SaveFileDialog`. File paths are validated to ensure they have the correct extensions (`.pdf`, `.xlsx`) to prevent unintended file overwrites.
- No external file dependencies (e.g., image files) are included in PDF generation, reducing the risk of unauthorized file access.
- File operations use `FileStream` and `ZipArchive` with proper disposal to prevent resource leaks.

### Data Processing
- Input data for PDF and XLSX generation is derived from in-memory `DataGridView` components and does not directly process user-provided input, minimizing risks of injection attacks.
- String escaping is implemented (e.g., `EscapeString`, `EscapeXml`, `EscapeXmlAttribute`) to prevent injection of malicious content into PDF or XLSX outputs.
- Random sample data is generated internally using `Random` for testing purposes. This data is not user-controlled and poses no security risk.

### Compression
- The application uses `DeflateStream` for zlib compression of font data and ICC profiles in PDF generation. The compression process is performed in-memory and does not involve external libraries prone to vulnerabilities.
- Adler-32 checksums are calculated to ensure data integrity during compression.

### Third-Party Dependencies
- This application does not rely on external NuGet packages or third-party libraries, reducing the risk of supply chain attacks.
- The embedded sRGB ICC profile is sourced from application resources and compressed securely.

## Best Practices for Secure Usage
- **Run in a Trusted Environment**: Execute the application on a trusted system to prevent unauthorized access to system resources via native API calls.
- **Validate Output Files**: Ensure output file paths are in secure, user-controlled directories to avoid overwriting critical system files.
- **Keep .NET Framework Updated**: Use the latest patched version of .NET Framework 4.8 to mitigate known vulnerabilities in the framework. Also it is highly recommended to upgrade this project to a newer version of .NET Framework in order to execute more secure and faster code.
- **Limit Permissions**: Run the application with least-privilege permissions to minimize the impact of potential exploits.

## Reporting a Vulnerability
If you discover a security vulnerability in this project, please report it responsibly by following these steps:
1. **Do Not Open a Public Issue**: To protect users, do not disclose vulnerabilities in public GitHub issues or discussions.
2. **Contact the Maintainer**: Email the vulnerability details to samialtas@gmail.com with the subject "Security Vulnerability in C# Unified Rendering: Identical PDF, XLSX & WinForms PrintPreview from Scratch".
3. **Provide Details**: Include a detailed description of the vulnerability, steps to reproduce, and potential impact.
4. **Response Time**: Expect an acknowledgment within 48 hours. We aim to address and resolve reported vulnerabilities promptly.

## Vulnerability Handling
- Reported vulnerabilities will be investigated and prioritized based on their severity.
- Patches or mitigations will be released in a timely manner, and affected versions will be updated in the supported versions table above.
- We will credit reporters (if desired) in release notes, unless anonymity is requested.

Thank you for helping keep this project secure!
