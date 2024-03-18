# PowerBIAnalyzer

## Description
PowerBIAnalyzer is a script that helps to identify unused fields (columns, calculated columns, or measures) in Power BI files. It works on Power BI Template (.pbit) files.

## Usage
1. Save your Power BI report as a template in .pbit extension
1. Download the PowerShell or Python script
1. Open a terminal and navigate to the directory where the script is located.
### PowerShell
1. Run the script using the following command:
    ```powershell
    .\PowerBIAnalyzer.ps1 -TemplatePath "path_to_pbit_file"
    ```
    Replace "path_to_pbit_file" with the actual path to your Power BI file.

### Python
1. Run the script using the following command:
    ```python
    TBD
    ```

## Script Logic
The PowerBIAnalyzer script performs the following steps:

1. Takes the input .pbit file, extracts it to a remporary location under the user's `%home%` directory (.pbit files are just .zip files with a different extension), reads the contents of the relevant files inside (DataModelSchema and Report\Layout), then deletes all the temporary content
2. Extracts all the columns, calculated columns, and measures from DataModelSchema
3. Checks all fields one by one against the relevant contents of Layout to identify fields which are directly referenced (as inputs or formatting for visuals)
4. Recursively checks all the unused fields if they are used in DAX expressions anywhere in the used fields
5. Prints out all the field references that are not used anywhere

## Requirements
- PowerShell 5.1 or later

## Contributing
Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License
This project is licensed under the [MIT License](https://opensource.org/license/mit).

## Acknowledgements
- [PowerShell](https://github.com/PowerShell/PowerShell)

## Disclaimer
This script is provided as-is without any warranty. Use it at your own risk.
