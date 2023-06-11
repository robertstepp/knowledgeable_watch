<#
    Robert Stepp, robert@robertstepp.ninja
    Functionality -

#>

# Import the required .NET assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

<# Debug settings
    No Debug output = SilentlyContinue
    Debug output = Continue
#>
$DebugPreference = 'Continue'

# Start the transcript
if ($DebugPreference -eq "Continue") {
    # Create a new StringBuilder object
    $logFileName = New-Object System.Text.StringBuilder
    
    # Append strings
    [void]$logFileName.Append("DebugLogs\")
    [void]$logFileName.Append((Get-Date -Format yyyyMMdd_HHmm))
    [void]$logFileName.Append("_debug.log")
    $logFile = Join-Path -Path (Get-ParentScriptFolder) -ChildPath $logFileName
    Start-Transcript -Path $logFile -Append
}
Write-Debug "Debug Preference: $($DebugPreference)"

# Get the path to the parent folder
function Get-ParentScriptFolder {
    $scriptPath = $MyInvocation.PSCommandPath
    $myParentFolder = Split-Path -Path $scriptPath
    Write-Debug "Parent Folder: $($myParentFolder)"
    return $myParentFolder
}

######################################

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
        }

        # Assign the closure as the button's event handler
        $button.Add_Click({ $handler.Invoke($this) })

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
    })

    # Add the "Done" button to the form
    $form.Controls.Add($doneButton)

    # Set the initial button states
    $buttons.Values | ForEach-Object {
        $_.BackColor = [System.Drawing.Color]::Gray
        $_.ForeColor = [System.Drawing.Color]::Black
    }

    # Show the form
    [void]$form.ShowDialog()

    # Return the selected items
    return $selectedItem
}

# Example usage:
$configFile = "config.ini"

Write-Debug "Reading config file..."
$configData = Read-ConfigFile -FilePath $configFile

Write-Debug "Showing item form..."
$selectedItems = Show-ItemForm -Items $configData

# Calculate totals or perform further operations with the selected items
Write-Debug "Selected items:"
Write-Host $selectedItems

######################################

if ($DebugPreference -eq "Continue") {
    # Stop the transcript
    Stop-Transcript
}