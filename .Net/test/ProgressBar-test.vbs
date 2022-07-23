'Manual integration test for ProgressBar.dll

Option Explicit
Dim pb 'VBScripting.ProgressBar object
Dim i 'integer

Set pb = CreateObject( "VBScripting.ProgressBar" )
pb.SetIconByDllFile "%SystemRoot%\System32\msdt.exe", 0
pb.FormLocationByPercentage 100, 100
pb.FormSize 500, 100
pb.PBarSize 400, 30
pb.PBarLocation 50, 40
pb.Visible = True

pb.Caption = "Continuous"
pb.Style = 1 'style 1 is continuous
pb.Minimum = 1
pb.Maximum = 300
For i = pb.Minimum To pb.Maximum
    pb.Value = i
    WScript.Sleep 1 
Next

pb.Caption = "Marquee"
pb.Style = 2 'style 2 is marquee
For i = 1 To 300
    WScript.Sleep 1
Next

Set pb = Nothing
