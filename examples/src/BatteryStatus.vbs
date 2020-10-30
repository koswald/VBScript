' script for BatteryStatus.hta

Option Explicit

Dim batteryMonitor

Sub Window_OnLoad
    Dim width, height 'window size in percent of screen
    Dim xPos, yPos 'window position in percent of screen
    Dim pxWidth, pxHeight
    Dim application 'hta object
    Dim app 'VBSApp object
    width = 22 ' default values
    height = 14
    xPos = 100
    yPos = 50
    With document.parentWindow.screen
        pxWidth = .availWidth * width * .01
        pxHeight = .availHeight * height * .01
        self.ResizeTo pxWidth, pxHeight
        self.MoveTo _
            (.availWidth - pxWidth) * xPos * .01, _
            (.availHeight - pxHeight) * yPos * .0102
    End With
    Set app = New VBSApp
    Set batteryMonitor = New VBSBatteryMonitor
    batteryMonitor.Monitor
    window.setTimeout "self.Close()", 7777, "VBScript"
End Sub

Sub Document_OnKeyUp
    If Esc = window.event.keyCode Then
        self.Close
    End If
    Const Esc = 27
End Sub

Class VBSBatteryMonitor

    Sub Monitor
        On Error Resume Next
            extractor.SetImageFormatPng
            If Err Then sh.PopUp "Error setting image format to PNG.", 3, app.GetFileName, vbInformation + vbSystemModal
        On Error Goto 0
        extractor.Save resFile, IconIndex, imageFile, largeIcon
        Set image = document.createElement("img")
        image.src = imageFile
        image.alt = "Battery image"
        imageContainer.insertBefore image

        Set messageDiv = document.createElement("div")
        textContainer.insertBefore messageDiv
        messageDiv.innerHTML = format(Array( _
            "Battery charge %s% <br /> %s", _
            Charge, Status _
        ))
    End Sub

    Property Get IconIndex
        Dim percent : percent = Charge
        If percent < 11 Then
            IconIndex = 9
        ElseIf percent < 31 Then
            IconIndex = 10
        ElseIf percent < 51 Then
            IconIndex = 11
        ElseIf percent < 71 Then
            IconIndex = 12
        Else IconIndex = 13
        End If
    End Property

    Property Get Charge
        Charge = Battery.EstimatedChargeRemaining
    End Property

    Property Get Status
        If Battery.BatteryStatus = 2 Then
            Status = "Plugged in"
        Else Status = "Not plugged in"
        End If
    End Property

    Property Get Battery : Set Battery = wmi.Battery : End Property

    Property Get Expand(str)
        Expand = sh.ExpandEnvironmentStrings(str)
    End Property

    Private largeIcon
    Private resFile 
    Private includer, extractor, wmi, format, app 'project objects
    Private image, messageDiv 'html objects
    Private imageFile 'string
    Private fso, sh

    Sub Class_Initialize
        largeIcon = True
        resFile = "%SystemRoot%\System32\wpdshext.dll"
        Set includer = CreateObject("VBScripting.Includer")
        Set extractor = CreateObject("VBScripting.IconExtractor")
        Execute includer.Read("WMIUtility")
        Set wmi = New WMIUtility
        Set sh = CreateObject("WScript.Shell")
        Set app = CreateObject("VBScripting.VBSApp")
        If IsEmpty(document) Then
            app.Init WScript
        Else app.Init document
        End If
        Set format = CreateObject("VBScripting.StringFormatter")
        imageFile = format(Array( Expand("%AppData%\VBScripting\%s.png"), app.GetBaseName ))
        Set fso = CreateObject("Scripting.FileSystemObject")
    End Sub

    Sub Class_Terminate
        If fso.FileExists(imageFile) Then
            fso.DeleteFile imageFile
        Else sh.PopUp "Image file not found.", 5, app.GetFileName, vbInformation
        End If
        Set fso = Nothing
        Set sh = Nothing
        Set format = Nothing
        Set app = Nothing
        Set includer = Nothing
        Set extractor = Nothing
    End Sub 

End Class
