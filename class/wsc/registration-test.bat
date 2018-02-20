@echo Expected outcome: Listing of COM object methods and properties
@powershell -noexit "New-Object -ComObject VBScripting.StringFormatter | Get-Member"