
'Get a folder or file chosen by the user

'Usage example
'
'' With CreateObject("VBScripting.Includer")
''     Execute .read("Chooser")
'' End With
'' 
'' Dim choose : Set choose = New Chooser
'' MsgBox choose.folder
'' MsgBox choose.file
'
'Browse for file <a href="http://stackoverflow.com/questions/21559775/vbscript-to-open-a-dialog-to-select-a-filepath"> reference</a>.
'Browse for folder <a href="http://ss64.com/vb/browseforfolder.html"> reference</a>.
'
Class Chooser

    'Function File
    'Returns a file path
    'Remark: Opens a Choose File dialog and returns the path of a file chosen by the user. Returns an empty string if no folder was selected. Note: The title bar text will say Choose File to Upload.
    Function File
        Dim oExec : Set oExec = sh.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
        SetMaxExecLifetime oExec, "mshta.exe", BFFileTimeout * 1000
        If SendKeysIsUnwiselyEnabled Then GoToRootPath
        File = oExec.StdOut.ReadLine
        Set oExec = Nothing
        If SendKeysIsUnwiselyEnabled Then MsgBox "" 'doesn't work otherwise
    End Function

    'Function Folder
    'Returns a folder path
    'Remark: Opens a Browse For Folder dialog and returns the path of a folder chosen by the user. Returns an empty string if no folder was selected.
    Function Folder
        Dim fo : Set fo = FolderObject
        If TypeName(fo) = "Nothing" Then Folder = "" : Exit Function
        Folder = fo.Self.Path
    End Function

    'Function FolderTitle
    'Returns the folder title
    'Remark: Opens a Browse For Folder dialog and returns the title of a folder chosen by the user. The title for a normal folder is just the folder name. For a special folder like %UserProfile%, it may be something entirely different. Returns an empty string if no folder was selected.
    Function FolderTitle
        Dim fo : Set fo = FolderObject
        If TypeName(fo) = "Nothing" Then FolderTitle = "" : Exit Function
        FolderTitle = fo.Title
    End Function

    'Function FolderObject
    'Returns an object
    'Remark: Opens a Browse For Folder dialog and returns a Shell.Application BrowseForFolder object for a folder chosen by the user. This object has methods Title and Self.Path, corresponding to this class's FolderTitle and FolderPath, respectively. This method is recommended for when you need both the FolderTitle and FolderPath but only want the user to have to choose once. If no folder was selected, then TypeName(folderObj) = "Nothing" is True.
    Function FolderObject
        Const WINDOW_HANDLE = 0
        Set FolderObject = sa.BrowseForFolder(WINDOW_HANDLE, WindowTitle, WindowOptions, sh.ExpandEnvironmentStrings(RootPath))
    End Function

    'Method SetWindowTitle
    'Parameter: a string
    'Remark: Sets the title of the Browse For Folder window: i.e. the text below the titlebar.
    Sub SetWindowTitle(newWindowTitle)
        WindowTitle = newWindowTitle
    End Sub

    'Method SetWindowOptions
    'Parameter: a hex value
    'Remark: Sets the behavior or behaviors for the Browse For Folder window. The parameter is one or more of the BIF_ constants:  e.g. obj.BIF_EDITBOX + obj.BIF_NONEWFOLDER.
    Sub SetWindowOptions(options)
        WindowOptions = options
    End Sub

    'Method AddWindowOptions
    'Parameter: a hex value
    'Remark: Adds a behavior or behaviors to the Browse For Folder window. The parameter is one or more of the BIF_ constants:  e.g. obj.BIF_EDITBOX + obj.BIF_NONEWFOLDER.
    Sub AddWindowOptions(newOptions)
        WindowOptions = WindowOptions + newOptions
    End Sub

    'Some procedures below are intentionally excluded
    'from the documentation by inserting two single quotes
    'instead of one before the procedure type

    'Property BIF_RETURNONLYFSDIRS
    'Returns &H0001
    Property Get BIF_RETURNONLYFSDIRS : BIF_RETURNONLYFSDIRS = &H0001 : End Property 'the default
    'Property BIF_DONTGOBELOWDOMAIN
    'Returns &H0002
    Property Get BIF_DONTGOBELOWDOMAIN : BIF_DONTGOBELOWDOMAIN = &H0002 : End Property
    'Property BIF_STATUSTEXT
    'Returns &H0004
    Property Get BIF_STATUSTEXT : BIF_STATUSTEXT = &H0004 : End Property
    'Property BIF_RETURNFSANCESTORS
    'Returns &H0008
    Property Get BIF_RETURNFSANCESTORS : BIF_RETURNFSANCESTORS = &H0008 : End Property
    ''Property BIF_EDITBOX
    'Returns &H0010
    Property Get BIF_EDITBOX : BIF_EDITBOX = &H0010 : End Property
    ''Property BIF_VALIDATE
    'Returns &H0020
    'Remark: Not recommended. Confusing behavior in Windows 10: Edit box appears but doesn't seem to change the functionality.
    Property Get BIF_VALIDATE : BIF_VALIDATE = &H0020 : End Property
    'Property BIF_NONEWFOLDER
    'Returns &H0200
    Property Get BIF_NONEWFOLDER : BIF_NONEWFOLDER = &H0200 : End Property
    'Property BIF_BROWSEFORCOMPUTER
    'Returns &H1000
    Property Get BIF_BROWSEFORCOMPUTER : BIF_BROWSEFORCOMPUTER = &H1000 : End Property
    'Property BIF_BROWSEFORPRINTER
    'Returns &H2000
    Property Get BIF_BROWSEFORPRINTER : BIF_BROWSEFORPRINTER = &H2000 : End Property
    ''Property BIF_BROWSEINCLUDEFILES
    'Returns &H4000
    'Remark: Not recommended. Confusing behavior with Windows 10: Files appear in the dialog, but if selected an error is raised.
    Property Get BIF_BROWSEINCLUDEFILES : BIF_BROWSEINCLUDEFILES = &H4000 : End Property

    'Method SetRootPath
    'Parameter: a folder path
    'Remark: Sets the root folder that the Browse For Folder window will allow browsing. Environment variables are allowed. See also the UnwiselyEnableSendKeys method.
    Sub SetRootPath(newRootPath)
        RootPath = newRootPath
    End Sub

    ''Property DESKTOP
    'Returns 0
    Property Get DESKTOP : DESKTOP = 0 : End Property
    ''Property PROGRAMS
    'Returns 2
    Property Get PROGRAMS : PROGRAMS = 2 : End Property
    ''Property DRIVES
    'Returns 17
    Property Get DRIVES : DRIVES = 17 : End Property
    ''Property NETWORK
    'Returns 18
    Property Get NETWORK : NETWORK = 18 : End Property
    ''Property NETHOOD
    'Returns 19
    Property Get NETHOOD : NETHOOD = 19 : End Property
    ''Property PROGRAMFILES
    'Returns 38
    Property Get PROGRAMFILES : PROGRAMFILES = 38 : End Property
    ''Property PROGRAMFILESx86
    'Returns 48
    Property Get PROGRAMFILESx86 : PROGRAMFILESx86 = 48 : End Property
    ''Property WINDOWS
    'Returns 36
    Property Get WINDOWS : WINDOWS = 36 : End Property

    'Method UnwiselyEnableSendKeys
    'Remark: Optional. Not recommended. Enables sending keystrokes to the Choose File to Upload dialog in order to open at the RootFolder. There is a risk whenever using the WScript.Shell SendKeys method that keystrokes will be sent to the wrong window.
    Sub UnwiselyEnableSendKeys : SendKeysIsUnwiselyEnabled = True : End Sub

    'Method WiselyDisableSendKeys
    'Remark: Default setting. Disables SendKeys. The Choose File to Upload dialog will open to the last place a file was selected, regardless of the RootFolder setting.
    Sub WiselyDisableSendKeys : SendKeysIsUnwiselyEnabled = False : End Sub

    'Method SetPatience
    'Parameter: time in seconds
    'Remark: Sets the maximum time in seconds that the File method waits for the Choose File to Upload dialog to appear before abandoning attempts to open the dialog at the folder specified by RootFolder. Applies only when SendKeys is enabled. Default is 5 (seconds).
    Sub SetPatience(newPatience) : patience = newPatience : End Sub

    'Function DialogHasOpened
    'Parameter: a string or an object
    'Returns a boolean
    'Remark: Waits for the specified dialog to appear, then returns False if the specified doesn't appear within the time specified by SetPatience, by default 5 (seconds). Parameter is either a string to match with the title bar text, as when browsing for a file, or else a WshScriptExec object, as when browsing for a folder. Used internally and by the unit test.
    Function DialogHasOpened(ActivateBy)
        Dim pActivateBy 'parsed ActivateBy
        If "String" = TypeName(ActivateBy) Then
            pActivateBy = ActivateBy
        Else
            pActivateBy = ActivateBy.ProcessId
        End If
        Const loopPause = 10 'milliseconds
        DialogHasOpened = True
        Dim StartTime : StartTime = Now
        While patience > DateDiff("s", StartTime, Now)
            If sh.AppActivate(pActivateBy) Then Exit Function
            app.Sleep loopPause
        Wend
        DialogHasOpened = False
    End Function

    'Navigate the browse for file dialog to the specified RootPath
    Private Sub GoToRootPath
        If Not SendKeysIsUnwiselyEnabled Then Exit Sub
        If DialogHasOpened(BFFileTitle) Then
            sh.AppActivate BFFileTitle
            app.Sleep pause
            sh.SendKeys "%n" 'Alt N => focus on file name field
            app.Sleep pause
            sh.SendKeys sh.ExpandEnvironmentStrings(RootPath)
            app.Sleep pause
            sh.SendKeys "{ENTER}"
        End If
    End Sub

    Sub SetPause(newPause) : pause = newPause : End Sub
    Property Get BFFolderTitle : BFFolderTitle = "Browse For Folder" : End Property
    Property Get BFFileTitle : BFFileTitle = "Choose File to Upload" : End Property

    'Method SetBFFileTimeout
    'Parameter: an integer
    'Remark: Sets the time in seconds after which the Browse For File (Choose File to Upload) dialog will be terminated if a file has not been chosen. A timeout of 0 will allow the dialog to remain open indefinitely. Intended to allow improved testing reliability. Default is 0.
    Sub SetBFFileTimeout(newBFFileTimeout)
        BFFileTimeout = newBFFileTimeout
    End Sub

    Private sh, sa, app
    Private WindowTitle, WindowOptions, RootPath
    Private SendKeysIsUnwiselyEnabled, patience, pause
    Private BFFileTimeout

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        Set sa = CreateObject("Shell.Application")
        With CreateObject("VBScripting.Includer")
            Execute .read("VBSApp")
        End With
        Set app = New VBSApp
        SetWindowTitle "Please select the folder"
        SetWindowOptions 0
        SetRootPath ""
        SetPatience 5
        WiselyDisableSendKeys
        SetPause 10
        SetBFFileTimeout 0
    End Sub

    'Method SetMaxExecLifetime
    'Parameters: WShellExec object, exe, milliseconds
    'Remark: Terminates a WShellExec process (the Browse for File window for example) after the specified time in milliseconds. Timeout of 0 prevents termination. An example of the exe: "mshta.exe".
    Sub SetMaxExecLifetime(oExec, exe, timeout)
        If 0 = timeout Then Exit Sub
        With CreateObject("VBScripting.Includer")
            Execute .read("WMIUtility")
        End With
        Dim wmi : Set wmi = New WMIUtility
        wmi.TerminateProcessByIdAndNameDelayed oExec.ProcessId, exe, timeout
    End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set sa = Nothing
    End Sub
End Class
