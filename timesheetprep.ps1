<#
    Robert Stepp, robert@robertstepp.ninja
    Functionality -
#>

# Import the required .NET assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Global variables
$startTime = $null
$selectedButton = $null

<# Debug settings
    No Debug output = SilentlyContinue
    Debug output = Continue
#>
$DebugPreference = 'Continue'

# Start the transcript
$transcriptPath = "DebugLogs\" + (Get-Date -Format 'yyyyMMdd_HHmm') + "_debug.log"
Start-Transcript -Path $transcriptPath -Append

# Get the path to the parent folder
function Get-ParentScriptFolder {
    $thisScriptPath = $MyInvocation.PSCommandPath
    $myParentFolder = Split-Path -Path $thisScriptPath
    Write-Debug "Parent Folder: $($myParentFolder)"
    return $myParentFolder
}

# Read the config file
function Read-ConfigFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    # Read the contents of the config file
    $configContent = Get-Content -Raw $FilePath

    # Split the content by line breaks
    $lines = $configContent -split "`r?`n"

    # Initialize an empty hashtable to store the configuration data
    $configData = @{}

    # Process each line in the config file
    foreach ($line in $lines) {
        # Skip empty lines or lines starting with a semicolon (comments)
        if (-not [string]::IsNullOrWhiteSpace($line) -and -not $line.TrimStart().StartsWith(";")) {
            $trimmedLine = $line.Trim()

            # Create key-value pair for the item
            $item = $trimmedLine
            if (-not $configData.ContainsKey($item)) {
                $configData[$item] = 0
            }
        }
    }

    # Return the configuration data hashtable
    return $configData
}

# Show the item form
function Show-ItemForm {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Items
    )

    # Create a new form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Time Entries"
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = 'CenterScreen'
    $form.AutoSize = $true

    # Create a hashtable to store the selected item values
    $selectedItems = @{}

    # Calculate the maximum button width
    $maxButtonWidth = ($Items.Keys | Measure-Object -Maximum -Property Length).Maximum * 10

    # Create buttons for each item
    $paddingTop = 20
    $buttonHeight = 30
    $buttonMargin = 10
    $index = 0
    $buttons = @{}
    foreach ($item in $Items.Keys | Sort-Object) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $item
        $buttonTop = $paddingTop + ($buttonHeight + $buttonMargin) * $index
        $button.Location = New-Object System.Drawing.Point(0, $buttonTop)
        $button.Size = New-Object System.Drawing.Size($maxButtonWidth, $buttonHeight)

        # Create a closure to capture the current button
        $handler = {
            param($sender)
            # Update the selected item and button states
            $selectedItems.Keys | ForEach-Object {
                $buttons[$_].BackColor = [System.Drawing.Color]::Gray
                $buttons[$_].ForeColor = [System.Drawing.Color]::Black
            }
            $selectedItems.Clear()
            $selectedItems[$sender.Text] = $Items[$sender.Text]
            $sender.BackColor = [System.Drawing.Color]::White
            $sender.ForeColor = [System.Drawing.Color]::Red
            Write-Debug $sender
            Save-ButtonPress -ButtonText $sender.Text
        }

        # Assign the closure as the button's event handler
        $button.Add_Click($handler)

        # Add the button to the form
        $form.Controls.Add($button)
        $buttons[$item] = $button
        $index++
    }

    # Create a "Done" button
    $doneButton = New-Object System.Windows.Forms.Button
    $doneButton.Text = "Done"
    $doneButton.Top = $paddingTop + ($buttonHeight + $buttonMargin) * $index
    $doneButton.Size = New-Object System.Drawing.Size($maxButtonWidth, $buttonHeight)
    $doneButton.BackColor = [System.Drawing.Color]::Gray
    $doneButton.Add_Click({
        $form.Close()
        Export-Timesheet
        # Delete the raw_timesheet CSV file
        Remove-Item -Path "raw_timesheet.csv" -Force -ErrorAction SilentlyContinue

    })

    # Add the "Done" button to the form
    $form.Controls.Add($doneButton)

    # Set the initial button states
    $buttons.Values | ForEach-Object {
        $_.BackColor = [System.Drawing.Color]::Gray
        $_.ForeColor = [System.Drawing.Color]::Black
    }

    # Show the form on top of all windows
    $form.TopMost = $true
    [void]$form.ShowDialog()
}

# Save the button press information
function Save-ButtonPress {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ButtonText
    )

    # Check if the button is the currently selected button
    if ($ButtonText -eq $global:selectedButton) {
        Write-Debug "Button already selected."
        return
    }

    # Check if the start time is already set
    if ($null -eq $global:startTime) {
        $global:startTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Debug "Start time: $($global:startTime)"
    } else {
        # Get the end time and calculate the duration
        $endTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $duration = New-TimeSpan -Start $global:startTime -End $endTime
        Write-Debug "End time: $endTime"
        Write-Debug "Duration: $($duration.TotalSeconds) seconds"

        # Write the data to the CSV file
        $data = [PSCustomObject]@{
            Button = $ButtonText
            Start = $global:startTime
            End = $endTime
            Duration = $duration.TotalSeconds
        }
        $data | Export-Csv -Path "raw_timesheet.csv" -Append -NoTypeInformation

        # Store the selected button
        $global:selectedButton = $ButtonText

        Write-Debug "Button selected: $ButtonText"
        $global:startTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Debug "Start time: $($global:startTime)"
    }
}

# Generate the timesheet
function Export-Timesheet {
    # Read the raw_timesheet CSV file
    $rawTimesheet = Import-Csv -Path "raw_timesheet.csv"

    # Group the raw timesheet data by button and calculate the total duration
    $groupedTimesheet = $rawTimesheet |
        Group-Object -Property Button |
        Select-Object @{Name = 'Button'; Expression = {$_.Name}}, @{Name = 'TotalDuration'; Expression = {($_.Group | Measure-Object -Property Duration -Sum).Sum}}

    # Create a new Excel workbook
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Add()

    # Set the worksheet name to "Today's Time"
    $date = Get-Date -Format 'yyyyMMdd'
    $worksheet.Name = "Today's Time"

    # Write the timesheet data to the worksheet
    $worksheet.Cells.Item(1, 1) = "Button"
    $worksheet.Cells.Item(1, 2) = "Total Duration (Minutes)"
    $worksheet.Cells.Item(1, 3) = "Total Duration (Tenths per Hour)"
    $row = 2
    foreach ($item in $groupedTimesheet) {
        $durationMinutes = [Math]::Ceiling($item.TotalDuration / 60)
        $durationTenthsPerHour = [Math]::Ceiling(($item.TotalDuration / 60) * 10)
        $worksheet.Cells.Item($row, 1) = $item.Button
        $worksheet.Cells.Item($row, 2) = $durationMinutes
        $worksheet.Cells.Item($row, 3) = $durationTenthsPerHour
        $row++
    }

    # Save the workbook to an Excel file
    $scriptDirectory = (Get-ParentScriptFolder)
    $excelFile = Join-Path -Path $scriptDirectory -ChildPath "Timesheet-$date.xlsx"
    $workbook.SaveAs($excelFile)

    Write-Debug "Timesheet generated: $excelFile"

    # Remove the "Sheet1" from the workbook, if it exists
    $sheetName = "Sheet1"
    $sheet = $workbook.Sheets | Where-Object { $_.Name -eq $sheetName }
    if ($sheet) {
        $sheet.Delete()
    }

    # Add a new sheet for the raw timesheet data
    $rawTimesheetSheet = $workbook.Worksheets.Add()
    $rawTimesheetSheet.Name = "raw-timesheet"

    # Write the raw timesheet data to the worksheet
    $row = 1
    foreach ($item in $rawTimesheet) {
        $column = 1
        foreach ($property in $item.PSObject.Properties) {
            $rawTimesheetSheet.Cells.Item($row, $column) = $property.Value
            $column++
        }
        $row++
    }

    # Reorder the sheets in the workbook
    $worksheet = $workbook.Worksheets | Where-Object { $_.Name -eq "Today's Time" }
    $worksheet.Move($workbook.Sheets.Item(1))

    # Save the workbook to include the raw timesheet data
    $workbook.Save()

    # Close the workbook and Excel application
    $workbook.Close()
    $excel.Quit()

    # Delete the raw_timesheet CSV file
    Remove-Item -Path "raw_timesheet.csv" -Force -Confirm:$false -ErrorAction SilentlyContinue
}

function Test-ExistingTimesheet {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExcelFile
    )

    # Check if the Excel file exists
    if (Test-Path $ExcelFile) {
        # Prompt the user for confirmation to proceed
        $result = [System.Windows.Forms.MessageBox]::Show("An Excel spreadsheet for today already exists. Do you want to continue and overwrite it?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Debug "Operation canceled. Script will exit."
            exit
        }
    }
}

$scriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$excelFile = Join-Path -Path $scriptDirectory -ChildPath "Timesheet-$(Get-Date -Format 'yyyyMMdd').xlsx"
Test-ExistingTimesheet -ExcelFile $excelFile

$configFile = "config.ini"
$configData = Read-ConfigFile -FilePath $configFile
Show-ItemForm -Items $configData

# Stop the transcript
Stop-Transcript
