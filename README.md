# Programa-para-convertir-archivos-XML-de-CFDI
Software for converting CFDI (Digital Tax Receipt via the Internet) XML files into Excel (.xlsx) files with a specific format for analysis and reporting.

# CFDI XML to Excel Converter

## Overview

CFDI XML to Excel Converter is a desktop application designed to transform Mexican electronic invoice files (CFDI - Comprobante Fiscal Digital por Internet) from XML format into structured, professionally formatted Excel reports. This tool is particularly useful for accountants, tax consultants, and businesses that need to consolidate and analyze large volumes of CFDI documents for auditing, financial analysis, and regulatory compliance.

The application automatically extracts over fifty relevant fields from CFDI files, including fiscal information, financial details, tax data, and payment information, organizing everything into a clean, easy-to-analyze spreadsheet format.

## Supported Formats

The application supports multiple CFDI-related standards including CFDI versions 3.3 and 4.0 for basic invoices and receipts, the Timbre Fiscal Digital (TFD) for digital authentication and fiscal validity, Complemento de Pagos 2.0 and 1.0 for payment transaction information, and Nómina 1.2 for payroll complements. This comprehensive support ensures compatibility with nearly all types of CFDI documents currently in use within Mexico's fiscal system.

## Key Features

The application provides flexible input options, allowing users to select an entire folder containing XML files (which are automatically scanned recursively through subfolders) or to manually select individual files. The processing is intelligent and non-blocking, maintaining a responsive user interface with real-time progress updates even when handling large batches of documents.

Data extraction is comprehensive and intelligent, capturing comprobante details such as series, folio, date, and UUID, information about all parties involved including RFC, business name, fiscal regime, and address, complete economic details including subtotals, discounts, amounts, currency, and exchange rates, fully delineated tax information including IVA at multiple rates (16%, 8%, 0%), IEPS, ISR, retentions and local taxes, comprehensive payment data including payment date, amount, payment method, and installment information, and relationship tracking including related CFDI UUIDs, cancellation reasons, and document modifications.

The application offers significant customization through its column selection panel, where users can choose exactly which fields to include in the generated Excel file. Quick action buttons allow selecting or deselecting all columns at once, making it easy to generate reports tailored to specific needs. The resulting Excel files feature professionally formatted headers with appropriate styling and automatic column width adjustment based on content.

Upon conversion completion, users receive detailed reporting including a real-time progress bar showing conversion status, a comprehensive list of any files that failed to process with specific error messages explaining what went wrong, and Excel files automatically named with timestamps to prevent overwriting previous reports.

## System Requirements

The application requires Windows 7 or later for the executable version, at least 100 MB of RAM for normal operation, and approximately 50 MB of disk space for the executable and its dependencies. For those running directly from Python source code, Python 3.8 or later is required with the openpyxl library as the only runtime dependency.

## Getting Started

For most users, the simplest approach is to locate and run the executable file "Convertidor XML a Excel.exe" in the Aplicación_Version_2 folder. Once launched, users should click either "Examinar carpeta" to select an entire folder of XML files or "Elegir archivos" to manually select specific files. Next, users specify the output folder where the generated Excel report will be saved. Optionally, the right-side panel allows customization of which columns appear in the final Excel file. Finally, clicking "Convertir a Excel" initiates the conversion process, which displays progress in real-time and produces a timestamped Excel file upon completion.

Alternatively, for users with Python installed, the application can be run directly from source code by first installing the required openpyxl library using pip, then executing the converter.py file. This approach is useful for development or when modifications to the processing logic are needed.

## Data Structure and Fields

The generated Excel reports can contain up to fifty-seven columns organized across several logical categories. The identification section includes period, CFDI version, UUID, series, folio, issuance date, and comprobante type. Party information contains RFC and business names for both issuer and receiver, their respective fiscal regimes, and the receiver's fiscal address. Economic data covers subtotal, discounts, total amounts, currency designation, and applicable exchange rates. Payment information includes payment date, amount, payment method, and details about partial payments and outstanding balances. Tax information is comprehensively broken down into IVA categories (16%, 8%, 0%), IEPS, ISR, and local taxes, with separate columns for both transferred and retained amounts. Relationship data tracks related CFDI UUIDs, relationship types, and cancellation information. Finally, complement-specific data includes digital stamp information, certifying authority details, and any payroll or other specialized complement information.

## Building the Executable

The project includes build scripts that automate the creation of a standalone Windows executable using PyInstaller. Users can either execute the provided build.bat script in the Aplicación_Version_2 folder or manually run PyInstaller with appropriate parameters. The build process automatically handles dependency bundling and produces a single executable file with no external dependencies required. The resulting exe file is self-contained and can be distributed to end users without requiring Python or any additional setup.

## Error Handling and Reporting

The application is designed to be robust and forgiving. If some XML files fail to process, the application continues processing all remaining files rather than stopping completely. A detailed list of failed files appears at the bottom of the application window after conversion, showing the specific filename and the reason for failure. The Excel report is still generated with all successfully processed records, allowing users to access valid data even when some files encounter issues. Common errors reported include XML structural problems, empty or corrupted files, and CFDI documents with missing or malformed elements.

## Development and Customization

The entire application is contained within a single Python file, converter.py, making it straightforward to understand and modify. The code is organized into three main functional layers: a parser that extracts data from CFDI XML files and returns structured data or error messages, an Excel writer that handles spreadsheet creation and formatting, and a Tkinter-based graphical user interface that manages user interaction and displays progress.

For developers wishing to add new fields to the extraction process, the workflow is straightforward. First, extract the desired value from the XML in the parsear_xml_cfdi function, then add the field name to the HEADERS list to ensure it appears as a column option in the interface and in the Excel output, and finally ensure that every field defined in HEADERS receives a value in the registro dictionary, using an empty string or zero as a default for optional fields. Special attention should be paid to tax retention and transfer fields, which must be extracted only from the global cfdi:Impuestos block to avoid double-counting the same values that appear at the individual item level.

## Technical Architecture

The parser function reads XML files and safely extracts attributes using helper functions that prevent errors from missing or malformed elements. It automatically detects the CFDI version and selects the appropriate XML namespace, ensuring compatibility across different document formats. The parser never raises exceptions to calling code; instead, all errors are caught, converted to human-readable Spanish descriptions, and returned as part of the result tuple.

The Excel writer function iterates through all provided XML files, calling the parser for each one, and writes successful results to an Excel workbook row by row. It applies professional formatting to the header row, automatically adjusts column widths, and returns comprehensive statistics about the conversion process. If no files process successfully, the writer correctly returns null to indicate that no file should be saved.

The user interface is built with Tkinter and implements a responsive design that never blocks on file I/O. Long-running operations like reading and converting files execute on background threads, with progress updates marshalled back to the main thread using the root.after mechanism. This ensures the window remains interactive and responsive throughout the conversion process.

## Usage Notes

The application is fully functional as distributed and requires no configuration or setup beyond execution. The Excel output uses automatic formatting with an appropriate color scheme and readable fonts. Column widths adjust automatically to fit content while maintaining reasonable maximum widths to prevent unwieldy spreadsheets. All processing is logged and made visible to the user, with clear indication of success and failure for each file processed.

For large batches of files, the application processes them sequentially with progress feedback, allowing users to accurately estimate completion time. The timestamp in the output filename ensures that multiple conversions never overwrite each other, allowing users to maintain a history of reports generated over time.
