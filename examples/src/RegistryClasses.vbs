'Script for RegistryClasses.hta

Const HKCU = &H80000001
Const HKLM = &H80000002
Const uneditablePids = "Folder"
Const undeletableVerbs = "open opennew print explore find openas properties printto runas runasuser" 'cannonical verbs
Const Esc = 27
Const F1 = 112
Const F6 = 117
Const F7 = 118
Const F8 = 119
Const Expanded = 0 'REG_EXPAND_SZ
Const NotExpanded = 1 'REG_SZ
Const NotFound = 2
Const noVerbsFound = -1 'UBound return value for zero-length array.

Dim fso 'Windows-native Scripting.FileSystemObject
Dim sh 'Windows-native WScriptShell object
Dim sa 'Windows-native Shell.Application object
Dim includer 'VBScripting.Includer object
Dim format 'VBScripting.StringFormatter object
Dim admin 'VBScripting.Admin object
Dim browser 'VBScripting.FileChooser object
Dim initialVerb 'string: an empty string or the currently selected verb
Dim focusOnThis 'an html element preselected/intended to get the focus
Dim reg 'RegistryUtility object
Dim regProv 'StdRegProv object
Dim deleter 'KeyDeleter object
Dim invalidChrs 'array of invalid filename chars
Dim application 'HTML Application element/object
Dim timeoutID 'integer: setTimeout return value. Can be used with the clearTimeout method to cancel/reset the timer.
Dim descr 'RegString object
Dim icon 'RegString object
Dim iconKey 'string: registry path
Dim template 'RegString object
Dim templateKey 'string: registry path
Dim commandType 'integer: a registry string type, either Expanded or NotExpanded.
Dim commandKey 'string: a registry path for reading/retrieving a verb command.
Dim verbs 'array of strings
Dim root 'long: registry hive specified by the user via the select element hiveSelector: either HKCU (&H80000001) or HKLM (&H80000002)
Dim pid 'string: pid may be may be * for all file types, or else it equals progId.Value or is read from the registry. Examples: txtFile, VBSFile, Word.Document.6. Note: progId is the html input element, a text box.
Dim typeKey 'string: a registry path. Example: Software\Classes\.txt
Dim pidKey 'string: a registry path. Example: Software\Classes\txtfile
Dim verbKey 'string: a registry path. Example: Software\Classes\txtfile\shell
Dim pidIsFromRegistry 'boolean
Dim baseWidth, baseHeight 'size in pixels of the main window
Dim NewVerbItemsHeight, ConfigureClassItemsHeight 'size of expanded portions of the window

Sub Window_OnLoad
    Dim optionHKCU, optionHKLM 'option elements

    baseWidth = 444
    baseHeight = 400
    NewVerbItemsHeight = 100
    ConfigureClassItemsHeight = 330
    Self.ResizeTo baseWidth, baseHeight
    Self.MoveTo 800, 000

    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set sa = CreateObject( "Shell.Application" )
    Set includer = CreateObject( "VBScripting.Includer" )
    Set format = CreateObject( "VBScripting.StringFormatter" )
    Set admin = CreateObject( "VBScripting.Admin" )
    Set browser = CreateObject( "VBscripting.FileChooser" )
    
    Execute includer.Read( "RegistryUtility" )
    Set reg = New RegistryUtility
    Set regProv = reg.Reg
    ExecuteGlobal includer.Read( "ValidFileName" )
    invalidChrs = InvalidWindowsFilenameChars
    Execute includer.Read( "KeyDeleter" )
    Set deleter = New KeyDeleter

    Set application = document.getElementsByTagName( "application" )(0)
    document.Title = application.applicationName
    Set optionHKCU = document.createElement( "option" )
    optionHKCU.value = HKCU
    optionHKCU.innerHTML = "HKEY_CURRENT_USER"
    hiveSelector.insertBefore optionHKCU
    Set optionHKLM = document.createElement( "option" )
    optionHKLM.value = HKLM
    optionHKLM.innerHTML = "HKEY_LOCAL_MACHINE"
    hiveSelector.insertBefore optionHKLM

    Feedback "WARNING: Backup the registry before modifying it!"

    fileType.value = ""
    progId.value = ""
    hiveSelector.selectedIndex = 0
    initialVerb = ""
    Set focusOnThis = fileType

    RefreshFields
End Sub

Sub Window_OnUnload
    Set fso = Nothing
    Set sh = Nothing
    Set sa = Nothing
    Set includer = Nothing
    Set format = Nothing
    Set admin = Nothing
    Set browser = Nothing
End Sub

Class RegString
    Public Value 'string
    Public StringType 'integer
    Public Exists 'boolean
    Function Init(newValue, newType, newExists)
        Value = newValue
        StringType = newType
        Exists = newExists
        Set Init = me
    End Function
End Class

Sub RefreshFields
    DeriveKeys
    DeriveVerbs
    RefreshHtmlElements
End Sub

'Generate the registry keys typeKey, pidKey, and verbKey.
'Software\Classes\.txt
'Software\Classes\txtfile
'Software\Classes\txtfile\shell
Sub DeriveKeys
    Dim pidString 'RegString object. See Class RegString in this file.

    root = CLng(hiveSelector.options(hiveSelector.selectedIndex).value)
    fileType.value = Trim(fileType.value)
    progId.value = Trim(progId.value)
    If "" = fileType.value Then
        typeKey = ""
    Else typeKey = format(Array("Software\Classes\.%s", fileType.value))
    End If
    Set pidString = GetRegString( root, typeKey, "" )

    'Don't allow * for a progId.Value
    If "*" = progId.Value Then
        fileType.Value = "*"
        progId.Value = ""
        Set focusOnThis = fileType

    '* was manually entered in the fileType input box.
    '* specifies all files.
    'Autofill the progId input box with an empty string.
    ElseIf "*" = fileType.value Then
        pid = "*"
        typeKey = ""
        progId.value = ""
        pidIsFromRegistry = False

    'The "Use registry value [for ProgId] if avaliable" checkbox is checked, so autofill the progId.Value with the value from the registry.
    'A typical scenario: With HKLM selected, the string "txt" was entered in the filetype input box, so the string "txtfile", read from Software\Classes\.txt is used to autofill the progId input box.
    ElseIf pidString.Exists _
    And useRegPid.checked _
    And Not "" = typeKey Then
        progId.value = pidString.value
        pid = pidString.value
        pidIsFromRegistry = True

    'waiting for a valid entry in one of the input fields
    ElseIf Not pidString.Exists _
    And useRegPid.checked Then
        progId.value = ""
        pid = ""
        Set focusOnThis = fileType
        pidIsFromRegistry = False

    'otherwise, use progId.Value for the pid
    Else
        pid = progId.value
        pidIsFromRegistry = False
    End If

    pidKey = format(Array("Software\Classes\%s", pid))
    verbKey = format(Array("%s\shell", pidKey))
    newVerbLegend.innerHTML = format(Array("New verb at %s\%s", GetRootString(root), verbKey))
End Sub

'Read the verbs for the registry keys just generated, and populate the verbSelector with options.
Sub DeriveVerbs
    regProv.EnumKey root, verbKey, verbs
    If Not IsArray(verbs) Then verbs = Array()
    verbSelector.innerHTML = ""
    If noVerbsFound = UBound(verbs) Then
        Dim nullOption
        Set nullOption = document.createElement( "option" )
        nullOption.innerHTML = "No verbs available"
        verbSelector.insertBefore nullOption
        Exit Sub
    End If
    Dim i, verb
    For i = 0 To UBound(verbs)
        Set verb = document.createElement( "option" )
        verb.innerHTML = verbs(i)
        verb.value = verbs(i)
        verbSelector.insertBefore verb
    Next
    Set verb = Nothing
End Sub

Sub RefreshHtmlElements
    Dim i 'integer
    Dim shiftRequired 'RegString object. See Class RegString in this file.
    Dim commandString 'RegString object. See Class RegString in this file.
    Hide newVerbSubfields
    Hide classSubfields
    If noVerbsFound = UBound(verbs) Then
        Hide commandFields
        EnableDisable
        FocusOnPreselectedElement
        Exit Sub
    End If
    For i = 0 To verbSelector.options.length - 1
        If LCase(initialVerb) = LCase(verbSelector.options(i).innerHTML) Then
            verbSelector.selectedIndex = i
        End If
    Next
    initialVerb = ""
    Set shiftRequired = GetRegString(root, format(Array( _
        "%s\%s", _
        verbKey, verbs(verbSelector.selectedIndex) _
    )), "Extended")
    If NotFound = shiftRequired.StringType Then
        extended.checked = False
    Else extended.checked = True
    End If
    commandKey = format(Array( _
        "%s\%s\Command", _
        verbKey, verbs(verbSelector.selectedIndex) _
    ))
    Set commandString = GetRegString( root, commandKey, "" )
    command.value = commandString.Value
    commandType = commandString.StringType
    Unhide commandFields
    EnableDisable
    FocusOnPreselectedElement
End Sub

'Attempts to read a registry string value. Returns an object with three properties. Value, StringType, and Exists. StringType is an integer, either Expanded, NotExpanded, or NotFound.
Function GetRegString(root, key, valueName)
    Dim value 'string
    Dim strType 'integer
    Dim exists 'boolean
    value = reg.GetStringValue(root, key, valueName)
    strType = NotExpanded
    exists = True
    If "Null" = TypeName(value) Then
        value = reg.GetExpandedStringValue(root, key, valueName)
        strType = Expanded
        exists = True
    End If
    If "Null" = TypeName(value) Then
        value = ""
        strType = NotFound
        exists = False
    End If
    Set GetRegString = New RegString.Init(value, strType, exists)
End Function

Sub TryDeletingKey(root, key)
    Const success = 0
    Dim response
    response = MsgBox(format(Array( _
        "Do you really want to delete the registry key %s\%s?", _
        GetRootString(root), key _
    )), vbExclamation + vbOKCancel + vbDefaultButton2, document.Title)
    If vbCancel = response Then Exit Sub
    If success = regProv.DeleteKey(root, key) Then
        RefreshFields
        Exit Sub
    End If
    response = MsgBox(format(Array( _
        "On the first attempt, which was expected to fail when there are subkeys, key ""%s\%s"" failed to delete. Do you want to use a more powerful key deleter?%s%s", _
        GetRootString(root), key, vbLf & vbLf, _
        "WARNING: All subkeys will be deleted!" _
    )), vbExclamation + vbOKCancel + vbDefaultButton2, document.Title)
    If vbCancel = response Then
        Exit Sub
    End If
    deleter.DeleteKey root, key
    RefreshFields
End Sub

Function GetRootString(root)
    If HKLM = root Then
        GetRootString = "HKLM"
    ElseIf HKCU = root Then
        GetRootString = "HKCU"
    End If
End Function

'Update the Disabled property of selected html elements
Sub EnableDisable
    Enable deleteFileType
    Enable deleteProgId
    Enable deleteVerb
    Enable verbSelector
    Enable progId
    Enable useRegPid
    Enable command
    Enable commandEnabler
    Enable browseForCommandButton
    Enable extended
    Enable newVerbButton
    Enable configureClassButton
    Enable elevateButton

    Enable nullFileCheckbox
    Enable templateFile
    Enable iconFile
    Enable classDescription
    Enable saveClassButton
    Enable cancelClassButton
    Enable browse4IconButton
    Enable browse4TemplateButton

    If Not VerbIsDeletable Then
        Disable deleteVerb
        Disable extended
    End If
    If Not FileTypeIsDeletable Then
        Disable deleteFileType
    End If
    If Not ProgIdIsDeletable Then
        Disable deleteProgId
    End If
    If HKLM = root And Not admin.PrivilegesAreElevated Then
        Disable deleteVerb
        Disable deleteFileType
        Disable deleteProgId
        Disable newVerbButton
        Disable command
        Disable commandEnabler
        Disable browseForCommandButton
        Disable extended

        Disable nullFileCheckbox
        Disable templateFile
        Disable iconFile
        Disable classDescription
        Disable saveClassButton
        Disable browse4IconButton
        Disable browse4TemplateButton
    ElseIf admin.PrivilegesAreElevated Then
        elevateButton.title = ""
        Disable elevateButton
    End If
    If InStr( verbKey, "\\" ) _
    Or "\" = Left(verbKey, 1) Then
        Disable newVerbButton
        Disable configureClassButton
    End If
    If Not commandEnabler.checked Then
        Disable command
        Disable browseForCommandButton
    End If
    If "" = fileType.value Then
        Disable deleteFileType
        Disable configureClassButton
    End If
    If "" = progId.value Then
        Disable deleteProgId
    End If
    If "" = pid Then
        Disable configureClassButton
    End If
    Dim p : For Each p In Split(uneditablePids)
        If LCase(p) = LCase(progId.value) Then
            Disable configureClassButton
        End If
    Next
    If "*" = pid Then
        Disable configureClassButton
        Disable useRegPid
        Disable progId
    End If
    If noVerbsFound = UBound(verbs) Then
        Disable verbSelector
        Disable deleteVerb
        Disable extended
    End If
    If ProgIdHasUndeletableVerb Then
        Disable deleteProgId
    End If
    If deleteVerb.disabled And Not noVerbsFound = UBound(verbs) Then
        Disable deleteProgId
    End If
    If deleteVerb.disabled Then
        Disable commandEnabler
    End If
    If commandEnabler.disabled Then
        commandEnabler.checked = False
        Disable command
        Disable browseForCommandButton
    End If
End Sub

'Event handlers

Sub deleteFileType_OnClick
    If Not FileTypeIsDeletable Then Exit Sub
    TryDeletingKey root, typeKey
    RefreshFields
End Sub

Sub deleteProgId_OnClick
    If Not ProgIdIsDeletable Then Exit Sub
    TryDeletingKey root, format(Array( _
        "Software\Classes\%s", pid))
    RefreshFields
End Sub

Sub deleteVerb_OnClick
    If Not VerbIsDeletable Then Exit Sub
    Dim verb : verb = verbs(verbSelector.selectedIndex)
    Dim prompt : prompt = format(Array( _
        "Do you really want to delete the verb ""%s"" at %s?", _
        verb, verbKey _
    ))
    Dim settings : settings = vbOKCancel + vbExclamation + vbDefaultButton2
    If vbCancel = MsgBox(prompt, settings, document.Title) Then
        Exit Sub
    End If
    regProv.DeleteKey root, format(Array( _
        "%s\%s\Command", verbKey, verb _
    ))
    regProv.DeleteKey root, format(Array( _
        "%s\%s", verbKey, verb _
    ))
    RefreshFields
End Sub

Sub fileType_OnKeyUp
    Set focusOnThis = fileType
    RefreshFields
End Sub

Sub command_OnKeyUp
    regProv.CreateKey root, commandKey
    Dim cmd : Set cmd = GetRegString( root, commandKey, "" )
    If Expanded = cmd.StringType Then
        regProv.SetExpandedStringValue root, commandKey, "", command.value
    Else regProv.SetStringValue root, commandKey, "", command.value
    End If
End Sub

Sub browseForCommandButton_OnClick
    browser.Title = format(Array( _
        "Choose a file to use with the verb ""%s""", _
        verbs(verbSelector.selectedIndex) _
    ))
    browser.InitialDirectory = "%UserProfile%"
    browser.Filter = "Script files (*.vbs; *.wsf)|*.vbs;*.wsf|HTA files (*.hta)|*.hta|EXE files (*.exe)|*.exe|All files (*.*)|*.*"
    browser.FilterIndex = 4
    Dim file : file = browser.FileName
    If Not "" = file Then
        Dim starter : starter = ""
        If "vbs" = LCase(fso.GetExtensionName(file)) Then starter = "wscript "
        If "wsf" = LCase(fso.GetExtensionName(file)) Then starter = "wscript "
        If "hta" = LCase(fso.GetExtensionName(file)) Then starter = "mshta "
        command.value = format(Array( _
            "%s""%s"" ""%1""", starter, file _
        ))
        command_OnKeyUp
    End If
End Sub

Sub commandEnabler_OnClick
    If commandEnabler.checked Then
        Enable command
        Enable browseForCommandButton
    Else Disable command
        Disable browseForCommandButton
    End If
End Sub

Sub progId_OnKeyUp
    If "" = progId.value Then
        useRegPid.checked = True
    Else useRegPid.checked = False
        Set focusOnThis = progId
        RefreshFields
    End If
End Sub

Sub verbSelector_OnChange
    initialVerb = verbs(verbSelector.selectedIndex)
    Set focusOnThis = verbSelector
    RefreshFields
End Sub

Sub extended_OnClick
    Dim key : key = format(Array( _
        "%s\%s", _
        verbKey, _
        verbs(verbSelector.selectedIndex) _
    ))
    If extended.checked Then
        regProv.SetStringValue root, key, "Extended", ""
    Else regProv.DeleteValue root, key, "Extended"
    End If
End Sub

Sub newVerbButton_OnClick
    If newVerbSubfields.style.display = "none" Then
        Hide classSubfields
        Unhide newVerbSubfields
    Else Hide newVerbSubfields
    End If
    newVerb.focus
End Sub

Sub newVerb_OnKeyUp
    If NewVerbIsValid Then
        Enable saveNewVerbButton
    Else Disable saveNewVerbButton
    End If
End Sub

Sub saveNewVerbButton_OnClick
    If Not NewVerbIsValid Then
        MsgBox "Verb is invalid.", vbInformation, document.Title
        Exit Sub
    End If
    newVerb.value = Trim(newVerb.value)
    Dim key : key = format(Array( _
        "%s\%s\command", _
        verbKey, newVerb.value _
    ))
    regProv.CreateKey root, key
    regProv.SetExpandedStringValue root, key, "", ""
    newVerbSubfields.style.display = "none"
    initialVerb = newVerb.value
    newVerb.value = ""
    If Not pidIsFromRegistry _
    And Not "*" = pid _
    And Not "" = typeKey Then
        regProv.CreateKey root, typeKey
        reg.SetStringValue root, typeKey, "", pid
    End If
    RefreshFields
    initialVerb = ""
    Enable command
    Enable browseForCommandButton
    commandEnabler.checked = True
    command.focus
End Sub

Sub cancelNewVerbButton_OnClick
    newVerb.value = ""
    newVerbSubfields.style.display = "none"
End Sub

Sub elevateButton_OnClick
    sa.ShellExecute "mshta", application.commandLine,, "runas"
    document.ParentWindow.close
End Sub

Sub nullFileCheckbox_OnClick
    If nullFileCheckbox.checked Then
        Disable templateFile
        Disable browse4TemplateButton
    Else Enable templateFile
        Enable browse4TemplateButton
    End If
End Sub

Sub browse4IconButton_OnClick
    browser.Filter = "Icon resources (*.exe; *.dll; *.ico)|*.exe; *.dll; *.ico|All files (*.*)|*.*"
    browser.InitialDirectory = "%ProgramFiles%"
    browser.Title = "Browse for the default icon file for progId: " & progId.value
    Dim file : file = browser.FileName
    If Not "" = file Then
        iconFile.value = file
    End If
End Sub

Sub browse4TemplateButton_OnClick
    browser.Filter = format(Array( _
        "%s files (*.%s)|*.%s|All files (*.*)|*.*" , _
        UCase(fileType.value), fileType.value, fileType.value _
    ))
    browser.InitialDirectory = "%ProgramFiles%"
    browser.Title = "Browse for the template file for ." & fileType.value & " files."
    Dim file : file = browser.FileName
    If Not "" = file Then
        templateFile.value = file
    End If
End Sub

Sub configureClassButton_OnClick
    If "" = typeKey Then Exit Sub
    If classSubfields.style.display = "none" Then
        Hide newVerbSubfields
        Unhide classSubfields
        shellNewLegend.innerHTML = format(Array("%s\%s\%s\ShellNew", GetRootString(root), typeKey, pid))
        progIdLegend.innerHTML = format(Array("%s\%s", getRootString(root), pidKey))
        Set descr = GetRegString( root, pidKey, "" )
        classDescription.value = descr.value
        iconKey = format(Array("%s\DefaultIcon", pidKey))
        Set icon = GetRegString( root, iconKey, "" )
        iconFile.value = icon.value
        templateKey = format(Array("%s\%s\ShellNew", typeKey, pid))
        Set template = GetRegString( root, templateKey, "FileName" )
        templateFile.value = template.value
        Dim nfString : Set nfString = GetRegString( root, templateKey, "NullFile" )
        If nfString.Exists Then
            templateFile.disabled = True
            nullFileCheckbox.checked = True
        Else nullFileCheckbox.checked = False
        End If
    Else Hide classSubfields
    End If
End Sub
Sub saveClassButton_OnClick
    regProv.CreateKey root, iconKey
    regProv.CreateKey root, templateKey
    reg.SetStringValue root, typeKey, "", pid
    Set descr = GetRegString( root, pidKey, "" )
    If NotExpanded = descr.StringType Then
        reg.SetStringValue root, pidKey, "", classDescription.value
    Else reg.SetExpandedStringValue root, pidKey, "", classDescription.value
    End If
    Set icon = GetRegString( root, iconKey, "" )
    If NotExpanded = icon.StringType Then
        reg.SetStringValue root, iconKey, "", iconFile.value
    Else reg.SetExpandedStringValue root, iconKey, "", iconFile.value
    End If
    Set template = GetRegString( root, templateKey, "FileName" )
    If NotExpanded = template.StringType Then
        reg.SetStringValue root, templateKey, "FileName", templateFile.value
    Else reg.SetExpandedStringValue root, templateKey, "FileName", templateFile.value
    End If
    If nullFileCheckbox.checked Then
        regProv.SetStringValue root, templateKey, "NullFile", ""
    Else regProv.DeleteValue root, templateKey, "NullFile"
    End If
    Hide classSubfields
End Sub

Sub cancelClassButton_OnClick
    Hide classSubfields
End Sub

Sub Document_OnKeyUp
    If Esc = window.event.keyCode Then
        self.Close
    ElseIf F1 = window.event.keyCode Then
        sh.Run "RegistryClasses.md"
    ElseIf F6 = window.event.keyCode Then
        sh.Run "notepad " & application.commandLine
    ElseIf F7 = window.event.keyCode Then
        sh.Run """%ProgramFiles%\Windows NT\Accessories\wordpad.exe"" " & application.commandLine
    ElseIf F8 = window.event.keyCode Then
        sh.Run """%ProgramFiles(x86)%\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe"" /edit " & application.commandLine
    End If
End Sub

Function NewVerbIsValid
    If Len(Trim(newVerb.value)) > 0 _
    And FileNameIsValid(newVerb.value) Then
        NewVerbIsValid = True
    Else NewVerbIsValid = False
    End If
End Function

Sub FocusOnPreselectedElement
    focusOnThis.focus
    If "input" = focusOnThis.tagName _
    And "text" = focusOnThis.type Then
        MoveCursorToTheEnd
    End If
End Sub
Sub MoveCursorToTheEnd
    Dim range : Set range = focusOnThis.CreateTextRange
    range.Collapse False
    range.Select
End Sub

Sub Hide(element)
    element.style.display = "none"
    Self.ResizeTo baseWidth, baseHeight
End Sub
Sub Unhide(element)
    element.style.display = "block"
    If classSubfields Is element Then
        Self.ResizeTo baseWidth, (baseHeight + ConfigureClassItemsHeight)
    ElseIf newVerbSubfields Is element Then
        Self.ResizeTo baseWidth,  (baseHeight + NewVerbItemsHeight)
    End If
End Sub
Sub Disable(element) : element.disabled = True : End Sub
Sub Enable(element) : element.disabled = False : End Sub

Function FileNameIsValid(name)
    Dim i : For i = 0 To UBound(invalidChrs)
        If InStr(name, invalidChrs(i)) Then FileNameIsValid = False : Exit Function
    Next
    FileNameIsValid = True
End Function

Function FileTypeIsDeletable
    Const dontDelete = "* exe com"
    Dim i, donts : donts = Split(LCase(dontDelete))
    For i = 0 To UBound(donts)
        If LCase(donts(i)) = LCase(fileType.value) Then FileTypeIsDeletable = False : Exit Function
    Next
    FileTypeIsDeletable = True
End Function

Function ProgIdIsDeletable
    Const dontDelete = "* Folder exefile"
    Dim i, donts : donts = Split(LCase(dontDelete))
    For i = 0 To UBound(donts)
        If LCase(donts(i)) = LCase(progId.value) Then ProgIdIsDeletable = False : Exit Function
    Next
    ProgIdIsDeletable = True
End Function

Function VerbIsDeletable
    Dim i, donts, verb : donts = Split(LCase(undeletableVerbs))
    verb = verbSelector.options(verbSelector.selectedIndex).value
    For i = 0 To UBound(donts)
        If LCase(donts(i)) = LCase(verb) Then VerbIsDeletable = False : Exit Function
    Next
    VerbIsDeletable = True
End Function

Function ProgIdHasUndeletableVerb
    Dim i, j, donts : donts = Split(LCase(undeletableVerbs))
    For j = 0 To UBound(verbs)
        For i = 0 To UBound(donts)
            If LCase(donts(i)) = LCase(verbs(j)) Then ProgIdHasUndeletableVerb = True : Exit Function
        Next
    Next
    ProgIdHasUndeletableVerb = False
End Function

Sub Feedback(newFeedback)
    feedbackDiv.innerHTML = newFeedback & "<br />" & feedbackDiv.innerHTML
    window.clearTimeout(timeoutID)
    timeoutID = window.setTimeout("ClearFeedback", 15000, "VBScript")
End Sub

Sub ClearFeedback
    feedbackDiv.innerHTML = ""
End Sub
