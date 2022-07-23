'Script for PresentationSettings.hta

Const synchronous = True
Const ForWriting = 2, CreateNew = True
Const timeoutKey = "SYSTEM\CurrentControlSet\Control\Power\PowerSettings\7516b95f-f776-4464-8c53-06167f40cc99\8EC4B3A5-6868-48c2-BE75-4F3044BE88A7"
Dim watcher, sh, regUtil, app, fso
Dim optionEnabled, optionDisabled
Dim errMsg

Sub Window_OnLoad
    self.resizeTo 450, 630
    self.moveTo 830, 0
    Set watcher = CreateObject( "VBScripting.Watcher" )
    Set app = CreateObject( "VBScripting.VBSApp" )
    app.Init document
    Dim includer : Set includer = CreateObject( "VBScripting.Includer" )
    Execute includer.Read( "RegistryUtility" )
    Set includer = Nothing
    Set regUtil = New RegistryUtility
    Set sh = CreateObject( "WScript.Shell" )
    sh.CurrentDirectory = app.GetParentFolderName
    Set fso = CreateObject( "Scripting.FileSystemObject" )

    Set enabler.OnChange = GetRef( "OptionChanged" )
    enabler.style.marginTop = "10px"
    Set errMsg = document.getElementsByTagName( "div" )(0)
    If Not watcher.Privileged Then
        enabler.disabled = True
        acSeconds.disabled = True
        dcSeconds.disabled = True
        saveAC.disabled = True
        saveDC.disabled = True
        Dim msg : msg = "Elevated privileges are required to change this setting."
        enabler.Title = msg
        acSeconds.Title = msg
        dcSeconds.Title = msg
    End If
    Set optionDisabled = document.createElement( "option" )
    Set optionEnabled = document.createElement( "option" )
    optionDisabled.innerHTML = "Disabled"
    optionEnabled.innerHTML = "Enabled"
    optionDisabled.value = disableTimeout
    optionEnabled.value = enableTimeout
    enabler.insertBefore optionDisabled
    enabler.insertBefore optionEnabled
    If disableTimeout = GetTimeoutAttributes Then
        enabler.SelectedIndex = 0
    ElseIf enableTimeout = GetTimeoutAttributes Then
        enabler.SelectedIndex = 1
    Else
        MsgBox "Sub Window_OnLoad: Unexpected timeout attributes value: " & GetTimeoutAttributes, vbExclamation, document.Title
    End If
    NormalMode
End Sub
Sub Window_OnUnload
    watcher.Dispose
    Set watcher = Nothing
    Set sh = Nothing
    Set regUtil = Nothing
    Set fso = Nothing
End Sub

Sub PresentationMode
    watcher.Watch = True
    document.Title = "Presentation mode is on"
    Enable offButton
    Disable onButton
    PublishStatus "Presentation"
End Sub

Sub NormalMode
    watcher.Watch = False
    document.Title = "Presentation mode is off"
    Enable onButton
    Disable offButton
    PublishStatus "Normal"
End Sub

Sub ChargerMode
    sh.Run "rundll32 user32.dll,LockWorkStation",, synchronous
    app.Sleep 1000
    watcher.MonitorOff
    app.Sleep 1000
    PresentationMode
    PublishStatus "Presentation"
End Sub

Sub Disable(htmlElement)
    htmlElement.disabled = True
End Sub
Sub Enable(htmlElement)
    htmlElement.disabled = False
End Sub

Sub PublishStatus(newStatus)
    Dim folder : folder = sh.ExpandEnvironmentStrings("%AppData%\VBScripting")
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    Dim statusFile : statusFile = folder & "\" & app.GetBaseName & ".status"
    Dim stream : Set stream = fso.OpenTextFile(statusFile, ForWriting, CreateNew)
    stream.WriteLine(newStatus)
    stream.Close
    Set stream = Nothing
End Sub

'Set the timeoutKey 'Attributes' DWord to 2 to enable 'Console lock display off timeout' in Advanced power options settings, Display section.
'Set to 1 to disable (default). https://www.windowscentral.com/how-extend-lock-screen-timeout-display-turn-windows-10
Const enableTimeout = 2
Const disableTimeout = 1
Sub SetTimeoutAttributes(newAttribute)
    regUtil.SetDWordValue regUtil.HKLM, timeoutKey, "Attributes", newAttribute
    If disableTimeout = newAttribute Then
        Disable acSeconds
        Disable saveAC
        Disable dcSeconds
        Disable saveDC
    ElseIf disableTimeout = newAttribute Then
        Enable acSeconds
        Enable saveAC
        Enable dcSeconds
        Enable saveDC
    End If
End Sub
Function GetTimeoutAttributes
    GetTimeoutAttributes = regUtil.GetDWordValue( regUtil.HKLM, timeoutKey, "Attributes" )
End Function
Sub OptionChanged
        SetTimeoutAttributes enabler.options(enabler.SelectedIndex).value
End Sub
Sub ChangeACTimeout
    sh.Run "cmd /c echo Changing AC VideoCOnLock to " & acSeconds.value & " seconds... & powercfg.exe /SETACVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOCONLOCK " & acSeconds.value & " & echo Done. & echo This setting may require a computer restart. & echo Press any key to exit & pause > nul",, synchronous
    sh.Run "cmd /c echo Updating the current power scheme... & powercfg.exe /S SCHEME_CURRENT & echo Done. Press any key to exit & pause > nul"
    acSeconds.value = ""
End Sub
Sub ChangeDCTimeout
    sh.Run "cmd /c echo Changing DC VideoCOnLock to " & dcSeconds.value & " seconds... & powercfg.exe /SETDCVALUEINDEX SCHEME_CURRENT SUB_VIDEO VIDEOCONLOCK " & dcSeconds.value & " & echo Done. & echo This setting may require a computer restart. & echo Press any key to exit & pause > nul",, synchronous
    sh.Run "cmd /c echo Updating the current power scheme... & powercfg.exe /S SCHEME_CURRENT & echo Done. Press any key to exit & pause > nul"
    dcSeconds.value = ""
End Sub
