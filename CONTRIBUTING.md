# Contributing to PDF Print Preview & XLSX Export

First off, thank you for considering contributing! This project is a community effort, and every contribution, from a small typo fix to a major new feature, is greatly appreciated.

This document provides a set of guidelines for contributing to the project. These are mostly guidelines, not strict rules. Use your best judgment, and feel free to propose changes to this document in a pull request.

## Code of Conduct

This project and everyone participating in it is governed by the [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected to uphold this code. Please report unacceptable behavior to the project maintainers.

## How Can I Contribute?

There are many ways to contribute, and not all of them involve writing code.

### Reporting Bugs

*   If you encounter a bug, please check the [issues list](../../issues) to see if it has already been reported.
*   If you don't see an open issue addressing the problem, [open a new one](../../issues/new). Be sure to include a **title and clear description**, as much relevant information as possible, and a **code sample or an executable test case** demonstrating the expected behavior that is not occurring.
*   Follow the bug report template provided on GitHub. See our [SUPPORT.md](SUPPORT.md) file for more details.

### Suggesting Enhancements

*   If you have an idea for an improvement or a new feature, we'd love to hear it! Please check the [issues list](../../issues) and [discussions](../../discussions) to see if it has been suggested before.
*   If not, open a new issue and use the "Feature Request" template to describe your idea. Explain why this enhancement would be useful and what problem it solves.

### Your First Code Contribution

Unsure where to begin contributing to the project? You can start by looking through these `good first issue` and `help wanted` issues:

*   [Good first issues][good-first-issue] - issues which should only require a few lines of code, and a test or two.
*   [Help wanted issues][help-wanted] - issues which should be a bit more involved than `good first issue`s.

## Getting Started: Development Setup

Ready to write some code? Here’s how to get your development environment set up.

1.  **Fork the repository** on GitHub.
2.  **Clone your fork** to your local machine:
    ```bash
    git clone https://github.com/YOUR_USERNAME/PDF-PrintPreview-XLSX-Export.git
    ```
3.  **Prerequisites:**
    *   You will need **Visual Studio 2022** (or a recent version).
    *   Make sure you have the **".NET desktop development"** workload installed via the Visual Studio Installer.
    *   The project targets the **.NET Framework**. The exact version is specified in the `.csproj` file.
4.  **Open the Solution:**
    *   Navigate to the cloned directory and open the `PDF_PrintPreview_XLSX_Export.sln` file in Visual Studio.
5.  **Build the Project:**
    *   Build the solution by pressing `Ctrl+Shift+B` or from the `Build` menu. Visual Studio should automatically restore any required NuGet packages.
    *   Run the project (`F5`) to ensure everything is working correctly before you make any changes.

## Pull Request Process

When you are ready to submit your changes, please follow this process:

1.  **Create a New Branch:**
    *   Create a new branch from `main` for your changes. Please give it a descriptive name.
    ```bash
    git checkout -b fix/pdf-font-encoding-bug
    # or
    git checkout -b feat/add-xlsx-cell-merging
    ```
2.  **Make Your Changes:**
    *   Write your code, and please adhere to the coding style outlined below.
    *   Add comments to your code where the logic is complex or not immediately obvious. This is especially important for the `P/Invoke` sections and low-level PDF/XLSX generation code.
3.  **Commit Your Changes:**
    *   Use clear and descriptive commit messages. A good commit message explains the "what" and the "why" of the change.
    ```bash
    git commit -m "Fix: Correctly handle character encoding for non-ANSI fonts in PDF"
    ```
4.  **Push to Your Fork:**
    *   Push your branch to your forked repository on GitHub.
    ```bash
    git push origin fix/pdf-font-encoding-bug
    ```
5.  **Open a Pull Request (PR):**
    *   Go to the original repository on GitHub and you will see a prompt to create a Pull Request from your new branch.
    *   Provide a clear title and a detailed description of the changes.
    *   **Reference any related issues** in your PR description (e.g., `Closes #34`). This helps automatically link the PR to the issue.
    *   Wait for a project maintainer to review your PR. We will do our best to provide feedback in a timely manner. Be prepared to make changes based on the feedback.

## Coding Style and Conventions

To keep the codebase consistent and easy to read, please follow these conventions:

*   **Follow existing style:** When in doubt, look at the existing code and try to match its style.
*   **C# Naming Conventions:** Use standard Microsoft C# Naming Conventions.
    *   `PascalCase` for classes, methods, properties, and events.
    *   `camelCase` for local variables and method parameters.
    *   Use `_camelCase` for private instance fields.
*   **Bracing:** Use the Allman style for braces, where each brace is on a new line.
    ```csharp
    public void MyMethod()
    {
        // code
    }
    ```
*   **Comments:** Use XML documentation comments (`///`) for public methods and properties. Use `//` for internal implementation comments.
*   **File Organization:** Do not change the existing project structure without discussing it first.

Thank you again for your interest in contributing!

[good-first-issue]: https://github.com/mustafasami/PDF-PrintPreview-XLSX-Export/labels/good%20first%20issue
[help-wanted]: https://github.com/mustafasami/PDF-PrintPreview-XLSX-Export/labels/help%20wanted
