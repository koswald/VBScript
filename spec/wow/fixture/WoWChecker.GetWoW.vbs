
'fixture for ../WoWChecker.spec.vbs

'outputs True if hosted by 32-bit cscript.exe
'outputs False if hosted by 64-bit cscript.exe

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "WoWChecker" )
End With

WScript.StdOut.WriteLine New WoWChecker.isWoW
