<#
    Robert Stepp, robert@robertstepp.ninja
    Functionality -
        Check hashes for files against a pre-existing file
        Can be run twice, once to build the initial file and again to compare
            the final against initial hashes.
#>

<# Debug settings
    No Debug output = SilentlyContinue
    Debug output = Continue
#>
$DebugPreference = 'SilentlyContinue'

# Start the transcript
if ($DebugPreference -eq "Continue") {
    $logFile = Join-Path -Path (Get-ParentScriptFolder) -ChildPath "debug.log"
    Start-Transcript -Path $logFile -Append
}
Write-Debug "Debug Preference: $($DebugPreference)"




if ($DebugPreference -eq "Continue") {
    # Stop the transcript
    Stop-Transcript
}