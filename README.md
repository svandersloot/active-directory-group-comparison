# AD Group Comparison Script

## Description

This PowerShell script compares users in specified Active Directory (AD) groups and outputs the results to an Excel file. It identifies users that are common to two AD groups and logs the comparison results, including timestamps, into a designated worksheet within the same Excel file. The script handles the creation of the worksheet and headers if they do not exist, and appends new results without overwriting existing data.

## Requirements

- Windows PowerShell
- Active Directory PowerShell Module
- ImportExcel PowerShell Module

## Installation

1. **Install Active Directory Module** (if not already installed):

    ```powershell
    Install-WindowsFeature -Name RSAT-AD-PowerShell
    ```

2. **Install ImportExcel Module**:

    ```powershell
    Install-Module -Name ImportExcel -Scope CurrentUser
    ```

## Usage

1. **Prepare the Excel File**:

    Ensure you have an Excel file named `AD_Group_Comparison.xlsx` in the same directory as the script. The file should have a worksheet named `Sheet1` with the following headers:

    - `AD Group 1`
    - `AD Group 2`

    Example content of `Sheet1`:

    | AD Group 1               | AD Group 2                  |
    |--------------------------|-----------------------------|
    | GrpGIT_Mobilitas_Write   | GrpGWGCC_SM_Prod_Admins     |
    | GroupA                   | GroupB                      |
    | GroupC                   | GroupD                      |

2. **Run the Script**:

    Execute the script from PowerShell:

    ```powershell
    .\Find_Users_In_Both_Groups.ps1
    ```

## Script Details

### Import Modules

The script imports the necessary PowerShell modules for Active Directory and Excel operations.

### Define Paths and Initialize Variables

The script determines its directory, defines the path to the Excel file, and initializes error logging and timestamps.

### Read Group Pairs

It reads the AD group pairs from `Sheet1` of the Excel file, ensuring both `AD Group 1` and `AD Group 2` columns have values.

### Compare AD Groups

For each pair of AD groups, the script:
- Retrieves the members of each group.
- Compares the two groups to find common users.
- Formats the results to display each user's full name and `samaccountname`.

### Write Results to Excel

The script ensures the `Comparison Results` worksheet exists and writes headers if they do not exist. It reads existing results, combines them with new results, and writes the combined results back to the Excel file, starting from row 2 to preserve headers.

### Error Handling

Any errors encountered during the process are logged to `Error_Log.txt` in the script's directory. The script provides a success or failure message upon completion.

## Logging

The `Comparison Results` worksheet will have the following columns:
- `AD Group 1`
- `AD Group 2`
- `Comparison Results`
- `Timestamp`

## Example Output

| AD Group 1               | AD Group 2                  | Comparison Results                                         | Timestamp           |
|--------------------------|-----------------------------|------------------------------------------------------------|---------------------|
| GrpGIT_Mobilitas_Write   | GrpGWGCC_SM_Prod_Admins     | Vandersloot, Shayne (gzxvand); Saravanan, Jayakumar (g2l1sar); Chi, Daniel (gfochi); Munaganuri, Hareesh Kumar (gc0xmun) | 2024-06-13 12:34:56 |
| GroupA                   | GroupB                      | No users found in both groups.                             | 2024-06-13 12:34:56 |
| GroupC                   | GroupD                      | No users found in both groups.                             | 2024-06-13 12:34:56 |

## Notes

- Ensure the specified AD groups exist and are accessible from the machine running the script.
- The script requires appropriate permissions to read AD group memberships and modify the Excel file.

## Troubleshooting

- **Common Errors**:
  - `A parameter cannot be found that matches parameter name 'WorksheetName'`: Ensure the ImportExcel module is correctly installed and imported.
  - `No column headers found on top row '1'`: Verify that the `Sheet1` worksheet contains the correct headers (`AD Group 1` and `AD Group 2`).

- **Log File**:
  - Check `Error_Log.txt` for detailed error messages and troubleshooting information.

## License

This script is provided "as-is" without any warranties or guarantees. Use at your own risk.
