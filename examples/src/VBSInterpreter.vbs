'Script for VBSInterpreter.hta

Dim resetButton 'button
Dim rewriteHistory 'button
Dim table 'table for showing VBScript and results
Dim typedStatement 'text input for VBScript statements
Dim help 'div element for F1 help
Dim feedback 'div element for misc. messaging
Dim autoExecHistory 'checkbox
Dim autoExecHistDiv 'div element
Dim fso 'Scripting.FileSystemObject object
Dim sh 'WScript.Shell object
Dim format 'VBScripting.StringFormatter object
Dim historyStream 'text stream object
Dim savedStatements 'delimited string of typed statements
Dim expressionResult 'string: result of Show method
Dim statementIndex 'integer
Dim historyStreamStatus 'integer: either statusReading or statusWriting
Dim rowCount 'integers
Dim showingHelp 'boolean
Dim historyFile 'filespec: where to save typed statements
Dim configFile 'filespec of the configuration file
Const EnterKey = 13 'window.event.keyCode
Const UpArrow = 38
Const DownArrow = 40
Const Esc = 27
Const F1 = 112
Const F9 = 120
Const ForReading = 1, ForWriting = 2, ForAppending = 8, CreateNew = True 'for OpenTextFile method
Const statusReading = 0, statusWriting = 1
Const delimiter = "_-_"

Sub Window_OnLoad
    Dim appData 'expanded %AppData% folder
    Dim baseName 'base name of this file
    Self.ResizeTo 660, 400
    Self.MoveTo 600, 150
    rowCount = 0
    savedStatements = ""
    statementIndex = -1
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set sh = CreateObject( "WScript.Shell" )
    Set format = CreateObject( "VBScripting.StringFormatter" )
    appData = sh.ExpandEnvironmentStrings("%AppData%")
    baseName = fso.GetBaseName(document.location.href)
    historyFile = format(Array( _
        "%s\VBScripting\%s.history", _
        appData, baseName _
    ))
    configFile = format(Array( _
        "%s\VBScripting\%s.config", _
        appData, baseName _
    ))
    showingHelp = False
    historyStreamStatus = statusReading
    SetupHtmlElements
    If autoExecHistory.checked Then
        ExecuteHistory
    End If
    historyStreamStatus = statusWriting
    ShowError ""
    On Error Resume Next
        Set historyStream = fso.OpenTextFile(historyFile, ForAppending, CreateNew)
        If Err Then ShowError Err.Description
        typedStatement.focus
    On Error Goto 0
End Sub

Sub Window_OnUnload
    On Error Resume Next
        historyStream.Close
        Set historyStream = Nothing
        Set fso = Nothing
        Set sh = Nothing
    On Error Goto 0
End Sub
    
Sub SetupHtmlElements
    Dim ckBoxLabel 'text node: checkbox label
    document.Title = document.getElementsByTagName( "application" )(0).applicationName
    document.body.style.fontFamily = "sans-serif"
    ' table
    Set table = document.createElement( "table" )
    document.body.insertBefore table
    StyleTable
    AddTableRow "Statement", "Result", False 'header row
    ' statement input
    Set typedStatement = document.createElement( "input" )
    typedStatement.type = "text"
    typedStatement.style.width = "100%"
    document.body.insertBefore typedStatement
    ' clear history button
    Set resetButton = document.createElement( "input" )
    resetButton.type = "button"
    resetButton.value = "Clear history"
    Set resetButton.OnClick = GetRef( "ClearHistory" )
    document.body.insertBefore resetButton
    ' edit history button
    Set rewriteHistory = document.createElement( "input" )
    rewriteHistory.type = "button"
    rewriteHistory.value = "Edit history"
    rewriteHistory.style.marginLeft = "30px"
    Set rewriteHistory.OnClick = GetRef( "EditHistory" )
    document.body.insertBefore rewriteHistory
    ' check box
    Set autoExecHistDiv = document.createElement( "div" )
    autoExecHistDiv.style.fontSize = "smaller"
    Set autoExecHistDiv.OnMouseOver = GetRef( "CheckBoxDivOnHover" )
    document.body.insertBefore autoExecHistDiv
    Set autoExecHistory = document.createElement( "input" )
    autoExecHistory.type = "checkbox"
    Set autoExecHistory.OnClick = GetRef( "UpdateConfigFileFromGui" )
    autoExecHistory.style.marginRight = "10px"
    autoExecHistDiv.insertBefore autoExecHistory
    Set ckBoxLabel = document.createTextNode("Execute history on load")
    autoExecHistDiv.insertBefore ckBoxLabel
    ' error message div
    Set feedback = document.createElement( "div" )
    feedback.style.color = "#700"
    document.body.insertBefore feedback
    ' F1 help message div
    Set help = document.createElement( "div" )
    document.body.insertBefore help
    ' update html element(s) from .config file
    UpdateGuiFromConfigFile
End Sub

Sub Document_OnKeyUp
    If Esc = window.event.keyCode Then
        typedStatement.value = ""
        typedStatement.focus
    ElseIf F1 = window.event.keyCode Then
        ToggleHelp
        typedStatement.focus
    ElseIf F9 = window.event.keyCode Then
        sh.Run "notepad """ & configFile & """"
    ElseIf EnterKey = window.event.keyCode Then
        xecute typedStatement.value
    ElseIf UpArrow = window.event.keyCode Then
        DecrementIndex
        ShowSavedStatement
    ElseIf DownArrow = window.event.keyCode Then
        IncrementIndex
        ShowSavedStatement
    End If
End Sub

Sub DecrementIndex
    If statementIndex > 0 Then
        statementIndex = statementIndex - 1
    End If
End Sub
Sub IncrementIndex
    If statementIndex < UBound(Split(savedStatements, delimiter)) Then
        statementIndex = statementIndex + 1
    End If
End Sub
Sub ShowSavedStatement
    typedStatement.value = Split(savedStatements, delimiter)(statementIndex)
End Sub

Sub CheckBoxDivOnHover
    autoExecHistDiv.style.cursor = "default"
End Sub

Sub xecute(statement)
    Dim result 'what to put in table column #2
    Dim erred 'boolean
    expressionResult = ""
    On Error Resume Next
        ExecuteGlobal statement
        If Err Then
            result = Err.Description
            erred = True
        ElseIf Len(expressionResult) Then
            result = expressionResult
            erred = False
        Else
            result = "ok"
            erred = False
        End If
    On Error Goto 0
    AddTableRow statement & " ", result, erred
    If historyStreamStatus = statusWriting Then
        historyStream.WriteLine statement
    End If
    savedStatements = savedStatements & delimiter & statement
    If Left(savedStatements, Len(delimiter)) = delimiter Then
        'remove leading delimiter from the very first statement
        savedStatements = Right(savedStatements, Len(savedStatements) - Len(delimiter))
    End If
    statementIndex = UBound(Split(savedStatements, delimiter)) + 1
    typedStatement.value = ""
    typedStatement.focus
    window.scrollBy 0, 30
End Sub

Sub AddTableRow(statement, result, erred)
    Dim row 'a new row to be added to the table
    Dim statementCell 'a new cell in the new row
    Dim resultsCell 'another new cell
    Set row = table.InsertRow(-1)
    Set statementCell = row.InsertCell(-1)
    Set resultsCell = row.InsertCell(-1)
    statementCell.InnerHTML = statement
    resultsCell.InnerHTML = result
    StyleRow row, erred, statementCell, resultsCell
    rowCount = rowCount + 1
End Sub

'evaluate an expression and show the result
'by typing "show CBool(-1)" for example
Sub Show(expression)
    expressionResult = expression
End Sub

Sub ClearHistory
    Dim m, i, s 'MsgBox settings for opt out
    m = "Do you really want to clear the history?"
    i = vbOKCancel + vbInformation
    i = i + vbDefaultButton2
    s = document.Title
    If vbCancel = MsgBox( m, i, s ) Then
        Exit Sub
    End If
    historyStream.Close
    If fso.FileExists(historyFile) Then
        fso.DeleteFile historyFile
    End If
    document.parentWindow.location.reload
End Sub

Sub ExecuteHistory
    On Error Resume Next
        Set historyStream = fso.OpenTextFile(historyFile)
        If Err Then Exit Sub
    On Error Goto 0
    While Not historyStream.AtEndOfStream
        Xecute historyStream.ReadLine
    Wend
    historyStream.Close
End Sub

Sub UpdateGuiFromConfigFile
    Dim executeHistoryOnLoad 'boolean
    'set default(s)
    executeHistoryOnLoad = True
    'get values from the .config file
    On Error Resume Next
        Execute fso.OpenTextFile(configFile).ReadAll
    On Error Goto 0
    'update the gui
    autoExecHistory.checked = executeHistoryOnLoad
End Sub

Sub UpdateConfigFileFromGui
    Dim stream 'text stream for writing
    Set stream = fso.OpenTextFile(configFile, ForWriting, CreateNew)
    stream.WriteLine "executeHistoryOnLoad = " & autoExecHistory.checked
    stream.Close
End Sub

Sub ToggleHelp
    If showingHelp Then
        help.innerHTML = ""
        showingHelp = False
    Else showingHelp = True
        help.innerHTML = _
            "<ul>" & _
                "<li> To execute a globally-scoped VBScript statement, " & _
                    "type it in the input box and press Enter. </li>" & _
                "<li> To show the return value of a function or expression, " & _
                    "enter <pre>show &lt;expression&gt;</pre> </li>" & _
                "<li> To cycle through the statement history, " & _
                    "press the up and down arrows. </li>" & _
                "<li> To clear the input field, press Esc. </li>" & _
                "<li> To toggle this message, press F1. </li>" & _
                "<li> To open the .config file, press F9. </li>" & _
            "</ul>"
            window.scrollBy 0, 300
        End If
End Sub

Sub EditHistory
    sh.Run "notepad """ & historyFile & """"
    document.parentWindow.Close
End Sub

Sub ShowError(errorDescription)
    feedback.innerHTML = errorDescription
End Sub

Sub StyleTable : With table.style
    .borderCollapse = "collapse"
    .fontSize = "13"
End With : End Sub

Sub StyleRow(row, erred, statementCell, resultsCell)
    With row.style
        If 0 = rowCount Then .fontWeight = "bold"
        If rowCount mod 2 Then .backgroundColor = "#eeefee"
    End With
    statementCell.style.paddingLeft = "10px"
    With resultsCell.style
        .paddingLeft = "10px"
        .paddingRight = "10px"
        If erred Then .color = "#700"
    End With
End Sub
