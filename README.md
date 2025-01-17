# ExcelToSQL

## Overview
ExcelToSQL is a powerful WPF application designed to convert Excel or CSV files into SQL queries. It supports creating or updating database tables for popular SQL databases such as MSSQL, MySQL, and PostgreSQL.

## Features
- Drag-and-drop file support.
- Auto-detection of file delimiters.
- Sheet selection for Excel files with multiple sheets.
- SQL type selection (MSSQL, MySQL, PostgreSQL) with MSSQL as the default.
- Options for "Create Table" or "Update Table" operations.
- Automatic saving of user preferences.
- Localization support with languages like English and Spanish.

## Installation

### Prerequisites
- .NET 9 Runtime
- Windows OS (for WPF support)

### Steps
1. Clone the repository:
    ```bash
    git clone https://github.com/jonavarro22/ExcelToSQL.git
    ```
2. Open the project in Visual Studio.
3. Restore NuGet packages.
4. Build and run the application.

## Usage
1. Launch the application.
2. Drag and drop a supported file (Excel or CSV) into the designated area or use the upload button.
3. If the file is an Excel workbook with multiple sheets, select the desired sheet when prompted.
4. Choose the desired operation: `Create Table` or `Update Table` using the ComboBox.
5. Select the target SQL type (default: MSSQL).
6. Click on `Generate SQL` to create the SQL query.
7. Save the generated SQL file to your preferred location.

## Localization
- The app supports multiple languages. Switch languages from the settings.
- Currently available languages:
  - English
  - Spanish

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE.txt) file for details.

## Contributing
We welcome contributions! Feel free to fork the repository, submit pull requests, or open issues for suggestions and bugs.

## Acknowledgments
- [CsvHelper](https://joshclose.github.io/CsvHelper/) for handling CSV parsing.
- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) for Excel file processing.

## Author
- [jonavarro22](https://github.com/jonavarro22)

---

**Note:** Ensure you comply with the licensing of third-party libraries used in the project.

