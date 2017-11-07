
'test VBSProgressBar.dll

Option Explicit

Initialize
Main
Cleanup

Sub Main
    Const pause = 1000
    pb.Visible = True
    Dim i : For i = 1 To pb.Maximum
        WScript.Sleep pause
        pb.PerformStep
    Next
End Sub

Dim pb

Sub Initialize
    Set pb = CreateObject("VBScript.ProgressBar")
    'pb.SuspendLayout
    pb.Debug = True
    pb.Caption = "testing - VBScript.ProgressBar"
    pb.SetIconByIcoFile "%drop%\h\+\BlueStarSVG.ico"
    pb.SetIconByDllFile "%SystemRoot%\System32\shell32.dll", 42
    pb.SetIconByDllFile "%SystemRoot%\System32\msdt.exe", 0
    pb.FormBorderStyle pb.BorderStyle.FixedDialog 'Fixed3D, FixedDialog, FixedSingle, FixedToolWindow, None, Sizable, SizableToolWindow 
    pb.FormLocationByPercentage 100, 100
    pb.FormSize 500, 100
    pb.PBarSize 400, 30
    pb.PBarLocation 50, 40
    pb.Minimum = 1
    pb.Maximum = 20
    pb.Step = 1
    pb.Value = 1
    'pb.ResumeLayout False
End Sub

Sub Cleanup
    Set pb = Nothing
End Sub
