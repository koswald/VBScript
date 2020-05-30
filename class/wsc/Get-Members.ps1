# PowerShell script
# List Windows Scripting Components (.wsc) members
# This script can be launched with Get-Members-launch.bat

@(
    'VBScripting.EventExample',
    'VBScripting.Includer',
    'VBScripting.StringFormatter',
    'VBScripting.VBSPower',
    'VBScripting.VBSApp',
    'VBScripting.KeyDeleter'

) | ForEach-Object {

    "`n`n---------------------- $_ members ----------------------------------------"
    New-Object -ComObject $_ | Get-Member
}
