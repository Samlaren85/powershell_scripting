# Import the required module
Import-Module AzureAD

# Authenticate to Azure AD
Connect-AzureAD



# Function to check if a user account is active and return the status
function CheckUserAccountStatus {
    param (
        [string]$UserPrincipalName
    )
    #$Create @usrAttr object
    $userObject = New-Object PSObject

    foreach ($attribute in $userAttributes) {
        # Adding properties to the user object (initialize with empty values)
        $userObject | Add-Member -MemberType NoteProperty -Name $attribute -Value $null
    }
    foreach ($attribute in $managerAttributes) {
        # Adding properties to the user object (initialize with empty values)
        $mngAttr = "Manager"+$attribute
        $userObject | Add-Member -MemberType NoteProperty -Name $mngAttr -Value $null
    }
    
    # Get user details
    try {
        $usr = Get-AzureADUser -objectid $UserPrincipalName | Select-Object $userAttributes
    }
    catch {
        foreach ($attr in $userObject){
            $attr = ""
        }
        return $userObject
    }

    # Get Manager details
    $manager = Get-AzureADUserManager -objectid $usr.UserPrincipalName | Select-Object $managerAttributes
    # Check if the account is enabled or disabled

    foreach ($attr in $userObject.PSObject.Properties){
        if ($attr.Name.StartsWith("Manager")) {
            $mngAttr = $attr.Name
            $mngAttr = $mngAttr.Replace("Manager","")
            $attr.Value = $manager.PSObject.Properties[$mngAttr].Value

        } 
        else {
            $attr.Value = $usr.PSObject.Properties[$attr.Name].Value
        }
    }
    return $userObject
    #return the selected attributes to caller
    
}

function UserLookUp {
    # Read the input excel file
    Write-Host = "Processing information!"
    $ErrorActionPreference = "SilentlyContinue"
    $excel = New-Object -Com Excel.Application
    $excel.Visible = $true
    #Open the importfile for edit.
    $wb = $excel.Workbooks.Open($InputExcelPath)
    $ws = $wb.sheets.item(1)
    $col = $ws.UsedRange.columns.count
    $rows = $ws.UsedRange.rows.count
    
    $i = 0
    # Setting up columns to collect the results
    foreach ($attr in $userAttributes) {
        $ws.cells.item(1, $col+$i+1).value = $attr
        $i++
    }  
    $j = 0
    foreach ($attr in $managerAttributes) {
        $ws.cells.item(1, $col+$i+$j+1).value = "Managers "+$attr
        $j++
    }  

    for ($y = 2; $y -le $rows; $y++) {
        # Searching for value in column set in settings.ColumnNumber
        $UserPrincipalName = $ws.cells.item($y, $columnNumber).text.Trim()+$mailDomain

        
        # Check user account status and collect the result
        $status = CheckUserAccountStatus -UserPrincipalName $UserPrincipalName
        $i = 1
        foreach ($value in $status.PSObject.Properties){
            $ws.cells.item($y, $col+$i).value = $value.Value.ToString()
            $i++
        }
    }
    
    
    $ErrorActionPreference = "Continue"
    
    # Export the results to the output CSV file
    $wb.Save()
    $wb.Close()
    $excel.Quit()

    Write-Host "User account statuses have been saved to $InputExcelPath"
}

# Prompt user for input CSV file path
#$InputExcelPath = Read-Host "Enter the full path to your input CSV file"
# Get the path of the script's directory
$scriptPath = $PSScriptRoot

# Define a dynamic path to a file in the same directory as the script
$settingsFile = Join-Path -Path $scriptPath -ChildPath "settings.json"
$inputExcelPath = ""

if (Test-Path $settingsFile) {
    # Import the JSON file content
    $settings = Get-Content -Path $settingsFile | ConvertFrom-Json

    # Fetching the settings from the JSON
    $mailDomain = $settings.MailDomain
    $userAttributes = $settings.UserAttributes
    $managerAttributes = $settings.ManagerAttributes
    $columnNumber = $settings.ColumnNumber
    $inputExcelPath = Join-Path -Path $scriptPath -ChildPath $settings.FilePath

}else {
    Write-Host "Settings file not found. Creating a new settings.json file..."

    #Prompt user for mailDomain
    $mailDomain = Read-Host "What are the mail domain (e.g., @example.com)"
    
    #Prompt user for userAttributes
    $userAttributesInput = Read-Host "What attributes on the user are you looking for (comma-separated, e.g., FirstName, LastName, AccountEnabled)"
    $userAttributes = $userAttributesInput -split ",\s*"  # Split input by commas and trim whitespace

    #Prompt user for managerAttributes
    $managerAttributesInput = Read-Host "What attributes on the manager are you looking for (comma-separated, e.g., DisplayName, Email)"
    $managerAttributes = $managerAttributesInput -split ",\s*"  # Split input by commas and trim whitespace
    
    #Prompt user for column
    $columnNumber = Read-Host "In what column are the userID (e.g., 1)"

    # Create a hashtable for default settings
     $defaultSettings = @{
        MailDomain      = $mailDomain
        UserAttributes  = $userAttributes
        ManagerAttributes = $managerAttributes
        ColumnNumber    = $columnNumber
        FilePath = ""
    }

    # Convert hashtable to JSON format
    $jsonContent = $defaultSettings | ConvertTo-Json -Depth 3

    # Write the JSON content to a file
    Set-Content -Path $settingsFile -Value $jsonContent

    Write-Host "New settings.json file created successfully."
}

if (-Not (Test-Path $inputExcelPath)) {
    # Load the Windows Forms assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Create a new OpenFileDialog object
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')  # Default folder (optional)
    $openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"  # Filter for Excel files
    $openFileDialog.Title = "Select an Excel (.xlsx) file"

    # Show the dialog and capture the result
    $dialogResult = $openFileDialog.ShowDialog()

    # If the user selected a file, $dialogResult will be 'OK'
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $excelFile = $openFileDialog.FileName

        # Check if the file exists
        if (Test-Path $excelFile) {
            Write-Host "Excel file selected: $excelFile"
            $inputExcelPath = -Path $excelFile
            UserLookUp
        }
    } else {
        Write-Host "No file selected."
    }
} else {
    UserLookUp
} 


