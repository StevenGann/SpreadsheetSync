# Obsidian Spreadsheet Sync Plugin

This plugin allows you to display parts of Excel, CSV, and other spreadsheet files directly in your Obsidian notes. It supports XLSX, XLS, and CSV files, and allows you to specify which cells to display using standard Excel range notation.

## Features

- Display spreadsheet data as tables in your notes
- Support for XLSX, XLS, and CSV files
- Specify cell ranges using standard Excel notation (e.g., A1:B2)
- Support for named sheets in Excel files
- Relative file paths from your notes

## Usage

Add a code block to your note with the `spreadsheet` language identifier, followed by the filename and cell range:

    ```spreadsheet
    example.xlsx(A1:B2)
    ```

To specify a particular sheet in an Excel file:

    ```spreadsheet
    example.xlsx(Sheet2!A1:B2)
    ```

The plugin will look for the spreadsheet file in the same directory as your note. You can also use relative paths:

    ```spreadsheet
    ../data/example.xlsx(A1:B2)
    ```

## Installation

### Using BRAT (Recommended)

1. Install the [BRAT](https://github.com/TfTHacker/obsidian42-brat) plugin in Obsidian
2. Open BRAT settings
3. Click "Add Beta plugin"
4. Enter the repository URL: `https://github.com/YOUR_GITHUB_USERNAME/SpreadsheetSync`
5. Enable the "Spreadsheet Sync" plugin in Obsidian's Community Plugins settings

### Manual Installation

1. Download the latest release
2. Extract the files into your vault's `.obsidian/plugins/obsidian-spreadsheet-sync/` directory
3. Enable the plugin in Obsidian's Community Plugins settings

## Examples

### Basic Example
```spreadsheet
data.xlsx(A1:B5)
```
This will display cells A1 through B5 from the first sheet of `data.xlsx`.

### Specific Sheet Example
```spreadsheet
budget.xlsx(Monthly!A1:D10)
```
This will display cells A1 through D10 from the "Monthly" sheet of `budget.xlsx`.

### CSV Example
```spreadsheet
data.csv(A1:C10)
```
This will display the specified range from a CSV file.

## Troubleshooting

### File Not Found
- Make sure the spreadsheet file is in the same directory as your note, or that the relative path is correct
- Check that the filename matches exactly (case-sensitive on some systems)
- Verify that the file extension is correct (.xlsx, .xls, or .csv)

### Invalid Range
- Ensure the range follows Excel notation (e.g., A1:B2)
- For specific sheets, use the format SheetName!A1:B2
- Check that the specified sheet exists in your Excel file

## Support

If you encounter any issues or have feature requests, please:
1. Check the [GitHub Issues](https://github.com/YOUR_GITHUB_USERNAME/SpreadsheetSync/issues) for existing reports
2. Create a new issue if needed, including:
   - A description of the problem
   - Steps to reproduce
   - Expected vs actual behavior
   - Sample files (if possible)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
