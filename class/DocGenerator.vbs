
'Generate html and markdown documentation for VBScript code based on well-formed comments.

'Usage Example
'<pre> With CreateObject("VBScripting.Includer") <br />     Execute .read("DocGenerator") <br /> End With <br /> With New DocGenerator <br />     .SetTitle "VBScript Utility Classes Documentation" <br />     .SetDocName "TheDocs.html" <br />     .SetFilesToDocument "*.vbs | *.wsf | *.wsc" <br />     .SetScriptFolder = "..\..\class" <br />     .SetDocFolder = "..\.." <br />     .Generate <br />     .View <br /> End With </pre>
'
'<h5> Example of well-formed comments before a Sub statement </h5>
' Note: A remark is required for Methods (Subs).

'<pre>'Method: SubName<br />'Parameters: varName, varType<br />'Remark: Details about the parameters.</pre>

'<h5> Example of well-formed comments before a Property or Function statement </h5>
' Note: A Returns (or Return or Returns: or Return:) is required with a Property or Function.

'<pre>'Property: PropertyName<br />'Returns: a string<br />'Remark: A remark is not required for a Property or Function.</pre>

'<h5> Notes for the general comment syntax at the beginning of a script </h5>

'Use a single quote (') for general comments <br />
'-- lines without html will be wrapped with p tags <br />
'-- lines with html will not be wrapped with p tags <br />
'-- use a single quote by itself for an empty line <br />
'-- for an empty line within a &ltpre&gt block, use two single quotes followed by a space. If you are using Visual Studio, you may need to change an option: Tools | Options | Environment | Trailing Whitespace | Remove Whitespace on Save: False <br />
'Use two single quotes for code: the text will be wrapped with pre tags. But for multi-line code snippets, enclose all lines, separated by &lt;br /&gt;, in single set of pre tags.<br />
'Use three single quotes for remarks that should not appear in the documentation <br />

'<h5> Notes for when the script does not contain a Class statement </h5>

'If the script doesn't contain a class statement, then the general comments at the beginning of the file must be separated from the rest of the file with line that begins with '''' (four single quotes)
'

Class DocGenerator

    Private outputStreamer, inputStreamer 'streamer objects
    Private script, doc 'input and output text streams
    Private md 'output text stream
    Private File 'fso File object
    Private fs 'VBSFileSystem object
    Private rf 'RegExFunctions object
    Private re 'RegExp object
    Private sh, fso

    Private indentUnit
    Private indent_ 'current indentation
    Private docFile
    Private methodPattern, propertyPattern, parametersPattern, returnsPattern, remarksPattern
    Private classPattern, altClassPattern, generalPattern, prePattern, ignorePattern
    Private routinePattern, routineType
    Private routineName, routineContent '*should* be the same; routineName is read from the code, routineContent is read from help comment
    Private method, property_, parameters, returns, remarks, general, pre, ignore 'enums
    Private methodContent, propertyContent, parametersContent, returnsContent, remarksContent, preContent 'strings
    Private generalContent 'delimited string
    Private oMatch, oMatches, subs 'objects which will need memory cleanup
    Private id
    Private TableHeaderWritten, ScriptHeaderWritten
    Private status, preClassStatement, postClassStatement
    Private filesToDocument, defaultFilesToDocument
    Private scriptFolder, docFolder, docName, docTitle 'required to be set by the calling script before calling the Generate method

    Sub Class_Initialize
        With CreateObject("VBScripting.Includer")
            Execute .read("TextStreamer")
            Execute .read("VBSFileSystem")
            Execute .read("RegExFunctions")
            ExecuteGlobal .Read("EscapeMd")
        End With

        'prepare output streamer
        Set outputStreamer = New TextStreamer
        outputStreamer.SetForWriting

        'prepare input streamer
        Set inputStreamer = New TextStreamer
        inputStreamer.SetForReading

        'more initialization
        Set fs = New VBSFileSystem
        Set rf = New RegExFunctions
        Set re = New RegExp
        re.IgnoreCase = True
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        InitializeLiterals
        indentUnit = "   "
        indent_ = ""
        ResetHelpContent
        SetFilesToDocument(defaultFilesToDocument)
        status = preClassStatement
        id = 0
        SetDocName ""
        scriptFolder = "" 'don't use the setter yet, or else an empty string will be resolved to an existing folder before being validated
        SetTitle ""
    End Sub

    Private Sub InitializeDocFiles
        docFile = docFolder & "\" & docName
        outputStreamer.SetFile docFile
        Set doc = outputStreamer.Open
        Set md = fso.OpenTextFile(docFolder & "\" & fso.GetBaseName(docName) & ".md", 2, True)
    End Sub

    Private Sub InitializeLiterals

        'regex patterns that identify lines commented out in the code
        methodPattern     = "^\s*'\s*Method\s*:?\s*(.*)\s*$"
        propertyPattern   = "^\s*'\s*(?:Property|Function)\s*:?\s*(.*)\s*$"
        parametersPattern = "^\s*'\s*Parameters?\s*:?\s*(.*)\s*$"
        returnsPattern    = "^\s*'\s*Returns?\s*:?\s*(.*)\s*$"
        remarksPattern    = "^\s*'\s*Remarks?\s*:?\s*(.*)\s*$"
        generalPattern    = "^\s*'(.*)$"
        prePattern        = "^\s*''(.*)$"
        ignorePattern     = "^\s*'''(.*)$"
        'identify a line that begins a routine
        ' no comment (') or End; may specify public or Private; may be Sub, Function, or Property; only one (?:match) is (captured): the routine name; allow comments afterwards
        routinePattern = "^[^'(?:End)]*(?:Public\s+|Public\s+Default\s+|Private\s+){0,1}(?:Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(\w+).*$"
        'identify a line that begins a class
        classPattern   = "^[^'(?:End)]*(?:Class\s+)(\w+).*$"
        altClassPattern = "^\s*''''(.*)$"

        'wildcard pattern(s)
        defaultFilesToDocument = "*.vbs" 'file types to document, by name

        method = "Method"
        property_ = "Property"
        parameters = "parameters"
        returns = "returns"
        remarks = "remarks"
        general = "general"
        pre = "pre"
        ignore = "ignore"
        preClassStatement = "preClassStatement"
        postClassStatement = "postClassStatement"
    End Sub

    Private Sub ResetHelpContent
        methodContent = ""
        propertyContent = ""
        parametersContent = ""
        returnsContent = ""
        remarksContent = ""
        generalContent = ""
        preContent = ""
    End Sub

    'Method SetScriptFolder
    'Parameter: a folder
    'Remark: Required. Must be set before calling the Generate method. Sets the folder containing the scripts to include in the generated documentation. Environment variables OK. Relative paths OK.
    Sub SetScriptFolder(newScriptFolder) : scriptFolder = fs.Resolve(newScriptFolder) : End Sub

    'Method SetDocFolder
    'Parameter: a folder
    'Remark: Required. Must be set before calling the Generate method. Sets the folder of the documentation file. Environment variables OK. Relative paths OK.
    Sub SetDocFolder(newDocFolder) : docFolder = fs.Resolve(newDocFolder) : End Sub

    'Method SetDocName
    'Parameter: a filename
    'Remark: Required. Must be set before calling the Generate method. Specifies the name of the documentation file, including the filename extension (.html suggested).
    Sub SetDocName(newDocName) : docName = newDocName : End Sub

    'Method SetTitle
    'Parameter: a string
    'Remark: Required. Must be set before calling the Generate method. Sets the title for the documentation.
    Sub SetTitle(newDocTitle) : docTitle = newDocTitle : End Sub

    'Method SetFilesToDocument
    'Parameter: A regular expression
    'Remark: Optional. Specifies which files to document: default is <strong> *.vbs </strong>. Separate multiple wildcards with " | ".
    Sub SetFilesToDocument(newFilesToDocument) : filesToDocument = rf.Pattern(newFilesToDocument) : End Sub

    Private Sub ValidateConfiguration
        Dim msg
        msg = "A title for the document must be set using SetTitle."
        If "" = docTitle Then Err.Raise 41, fs.SName, msg

        msg = "An existing folder containing the scripts to document must be specified with SetScriptFolder."
        If Not fso.FolderExists(scriptFolder) Then Err.Raise 42, fs.SName, msg

        msg = "An existing folder to contain the document must be specified with SetDocFolder."
        If Not fso.FolderExists(docFolder) Then Err.Raise 43, fs.SName, msg

        msg = "The name of the doc file must be specified with SetDocName."
        If "" = docName Then Err.Raise 44, fs.SName, msg
        InitializeDocFiles
    End Sub

    'Method Generate
    'Remark: Generate comment-based documentation for the scripts in the specified folder.
    Sub Generate
        ValidateConfiguration
        WriteTopSection
        'for each class file, look through the file for comments to add to the documentation
        For Each File In fso.GetFolder(scriptFolder).Files
            re.Pattern = filesToDocument
            If re.Test(File.Name) Then WriteScriptSection(File)
        Next
        WriteBottomSection
    End Sub

    'Method View
    'Remark: Open the documentation file for viewing
    Sub View
        sh.Run """" & docFile & """"
    End Sub

    Private Sub WriteScriptSection(File)
        If InStr(fso.GetBaseName(File.Name), ".") Then Exit Sub
        inputStreamer.SetFile(File.Path)
        Set script = inputStreamer.Open
        TableHeaderWritten = False
        ScriptHeaderWritten = False
        ResetHelpContent
        status = preClassStatement

        While Not script.AtEndOfStream
            ProcessLine(script.ReadLine)
        Wend

        If TableHeaderWritten Then
            CloseTagsForTheTable
        End If

        If ScriptHeaderWritten Then
            CloseTagsForTheScript
            id = id + 1
        End If

        script.Close
    End Sub

    'Look for "help content"
    'That is, look for comments intended to be included in the
    'documnetation: Method or Property, Parameters, Returns, Remarks;
    'also look for routines: Sub, Property, Function
    Private Sub ProcessLine(line)
        If LineStartsAClass(line) Then
            status = postClassStatement
            ResetHelpContent
        ElseIf LineStartsARoutine(line) Then
            ValidateHelpContent
            WriteHelpContentToDoc
            ResetHelpContent
        Else
            GetAnyHelpContent(line)
        End If
    End Sub

    'Write the initial html for the current script, not including the general comments nor the table header
    Private Sub WriteScriptHeader
        doc.WriteLine ""
        WriteLine "<div>"
        IndentIncrease
        WriteLine "<h4 class=""heading"" id=" & id & ">" & fso.GetBaseName(File.Name) & "</h4>"
        WriteLine "<div class=""detail"">"
        md.WriteLine ""
        md.WriteLine "## "& fso.GetBaseName(File.Name)
        ScriptHeaderWritten = True
    End Sub

    Private Sub CloseTagsForTheScript
        WriteLine "</div>"
        IndentDecrease
        WriteLine "</div>"
    End Sub

    'Write the table header, which immediately follows the script header and general comments, if any
    Private Sub WriteTableHeader
        If Not ScriptHeaderWritten Then
            WriteScriptHeader
        End If
        IndentIncrease
        WriteLine "<table>"
        IndentIncrease
        WriteLine "<tr>"
        IndentIncrease
        WriteLine "<th>Procedure type</th>"
        WriteLine "<th>Name</th>"
        WriteLine "<th>Parameter(s)</th>"
        WriteLine "<th>Return value</th>"
        WriteLine "<th>Comment</th>"
        IndentDecrease
        WriteLine "</tr>"
        md.WriteLine "| Procedure | Name | Parameter | Return | Comment |"
        md.WriteLine "| :-------- | :--- | :-------- | :----- | :------ |"
        TableHeaderWritten = True
    End Sub

    Private Sub CloseTagsForTheTable
        IndentDecrease
        WriteLine "</table>"
        IndentDecrease
    End Sub


    'Write the general help content to file; don't include <p> tags if line already contains html
    Private Sub WriteGeneralContentToDoc
        If postClassStatement = status Then Exit Sub
        If Not ScriptHeaderWritten Then
            WriteScriptHeader
        End If
        IndentIncrease
        If Instr(generalContent, "<") Then
            WriteLine generalContent
        Else
            WriteLine "<p>" & generalContent & "</p>"
        End If
        IndentDecrease
        
        If InStr(generalContent, "<pre>") Then
            Dim lines : lines = Replace(generalContent, "<pre>", "```vb" & vbCrLf)
            lines = Replace(lines, "<br />", vbCrLf)
            lines = Replace(lines, "</pre>", vbCrLf & "```")
            lines = Replace(lines, "&lt;", "<")
            lines = Replace(lines, "&LT;", "<")
            md.WriteLine lines & "  "
        Else
            md.WriteLine generalContent & "  "
        End If
    End Sub

    Private Sub WritePreContentToDoc
        If postClassStatement = status Then Exit Sub
        If Not ScriptHeaderWritten Then
            WriteScriptHeader
        End If
        IndentIncrease
        WriteLine "<pre>" & preContent & "</pre>"
        IndentDecrease
        md.WriteLine "<pre>" & preContent & "</pre>"
    End Sub

    'Write the help content for the current routine in the current script
    Private Sub WriteHelpContentToDoc
        If NoComments Then Exit Sub 'don't require comments
        If Not TableHeaderWritten Then
            WriteTableHeader
        End If
        WriteLine "<tr>"
        IndentIncrease
        WriteLine "<td>" & routineType & "</td>"
        WriteLine "<td>" & routineName & "</td>"
        WriteLine "<td>" & parametersContent & "</td>"
        WriteLine "<td>" & returnsContent & "</td>"
        WriteLine "<td>" & remarksContent & "</td>"
        IndentDecrease
        WriteLine "</tr>"
        md.WriteLine "|" & routineType & "|" & routineName & "|" & parametersContent & "|" & returnsContent & "|" & EscapeMd(remarksContent) & "|"
    End Sub

    Private Property Get NoComments
        If Len(methodContent) Or Len(propertyContent) Then NoComments = False Else NoComments = True
    End Property

    Private Property Get LineStartsAClass(line)
        re.Pattern = classPattern
        If re.Test(line) Then LineStartsAClass = True : Exit Property
        re.Pattern = altClassPattern
        If re.Test(line) Then LineStartsAClass = True : Exit Property
        LineStartsAClass = False
    End Property

    Private Property Get LineStartsARoutine(line)
        LineStartsARoutine = False
        re.Pattern = routinePattern
        If re.Test(line) Then
            LineStartsARoutine = True
            Set subs = GetSubMatches(line)
            routineName = subs(0)
        End If
    End Property

    Private Sub ValidateHelpContent
        Dim msg
        If NoComments Then Exit Sub 'don't require comments

        msg = "Content can't have both Method and Property content."
        If Len(methodContent) And Len(propertyContent) Then RaiseContentError msg

        msg = "The help content descriptor of the method or property or function should equal the method or property or function name."
        If Len(methodContent) Then
            routineType = method
            routineContent = methodContent
        Else
            routineType = property_
            routineContent = propertyContent
        End If
        If Not routineName = routineContent Then RaiseContentError msg
        If method = routineType Then

            msg = "Methods may have parameters; may not have Returns; must have Remarks."
            If "" = remarksContent Then RaiseContentError msg
            If Len(returnsContent) Then RaiseContentError msg Else returnsContent = "N/A"
            If "" = parametersContent Then parametersContent = "None"
        ElseIf property_ =  routineType Then

            msg = "Properties may have parameters; may have Remarks; must have Returns. NOTE: Property Set and Property Let do not require a return value but still must have a 'Return or 'Returns comment."
            'TODO: allow Property Let and Property Set to have no return comment
            If "" = returnsContent Then RaiseContentError msg
            If "" = parametersContent Then parametersContent = "None"
            If "" = remarksContent Then remarksContent = "None"
        End If

    End Sub

    Private Sub RaiseContentError(msg) : Err.Raise 1,, File.Name & "::" & routineName & ": " & msg : End Sub

    'Get help content from a line
    Private Sub GetAnyHelpContent(line)
        If HasHelpContent(method, line) And postClassStatement = status Then
            methodContent = subs(0)
        ElseIf HasHelpContent(property_, line)  And postClassStatement = status Then
            propertyContent = subs(0)
        ElseIf HasHelpContent(parameters, line) And postClassStatement = status Then
            parametersContent = subs(0)
        ElseIf HasHelpContent(returns, line) And postClassStatement = status Then
            returnsContent = subs(0)
        ElseIf HasHelpContent(remarks, line) And postClassStatement = status Then
            remarksContent = subs(0)
        ElseIf HasHelpContent(ignore, line) Then
            Exit Sub
        ElseIf HasHelpContent(pre, line) Then
            preContent = subs(0)
            WritePreContentToDoc
        ElseIf HasHelpContent(general, line) Then
            generalContent = subs(0)
            WriteGeneralContentToDoc
        End If
    End Sub

    'Return True if a line has help content
    'Set the submatches object, which contains the help content
    Private Property Get HasHelpContent(helpType, line)
        If method = helpType Then
            re.Pattern = methodPattern
        ElseIf property_ = helpType Then
            re.Pattern = propertyPattern
        ElseIf parameters = helpType Then
            re.Pattern = parametersPattern
        ElseIf returns = helpType Then
            re.Pattern = returnsPattern
        ElseIf remarks = helpType Then
            re.Pattern = remarksPattern
        ElseIf ignore = helpType Then
            re.Pattern = ignorePattern
        ElseIf general = helpType Then
            re.Pattern = generalPattern
        ElseIf pre = helpType Then
            re.Pattern = prePattern
        End If
        If re.Test(line) Then
            HasHelpContent = True
            Set subs = GetSubMatches(line) 'get the help content
        Else
            HasHelpContent = False
        End If
    End Property

    'Return the desired content from a line, exclusive of white space and help-topic method
    Private Property Get GetSubMatches(line)
        're pattern has been set already in Property HasHelpContent
        Set oMatches = re.Execute(line)
        Set oMatch = oMatches(0)
        Set GetSubMatches = oMatch.SubMatches
    End Property

    Private Sub WriteLine(line) : doc.WriteLine indent_ & line : End Sub
    Private Sub Write_(str) : doc.Write str : End Sub
    Private Sub Indent : Write_ indent_ : End Sub
    Private Sub IndentIncrease : indent_ = indent_ & indentUnit : End Sub
    Private Sub IndentDecrease : indent_ = Replace(indent_, indentUnit, "", 1, 1) : End Sub

    Private Sub WriteTopSection
        WriteLine "<!DOCTYPE html>"
        WriteLine "<html>"
        IndentIncrease
        WriteLine "<!-- This file is automatically generated, so any changes that you make may be overwritten -->"
        WriteLine "<head>"
        IndentIncrease
        WriteLine "<title> " & docTitle & " </title>"
        WriteLine "<link type=""text/css"" rel=""stylesheet"" href=""lib/docStyle.css"" />"
        WriteLine "<link type=""text/css"" rel=""stylesheet"" href=""lib/docStyleTable.css"" />"
        IndentDecrease
        WriteLine "</head>"
        WriteLine "<body onclick=""docScript.toggleDetail(event)"">"
        WriteLine "<h3>" & docTitle & "</h3>"
        doc.WriteLine ""
        IndentIncrease

        md.WriteLine "# VBScript Classes"
        md.WriteLine ""
        md.WriteLine "### Contents"
        md.WriteLine ""
        Dim baseName
        For Each File In fso.GetFolder(scriptFolder).Files
            re.Pattern = filesToDocument
            baseName = fso.GetBaseName(File.Name)
            If re.Test(File.Name) And Not CBool(InStr(baseName, ".")) Then
                md.WriteLine "[" & baseName & "](#" & LCase(baseName) & ")  "
            End If
        Next
        md.WriteLine ""
    End Sub

    Private Sub WriteBottomSection
        WriteLine "<p> <em> See also the <a href=""https://msdn.microsoft.com/en-us/library/t0aew7h6(v=vs.84).aspx""> VBScript docs </a> </em> </p>"
        WriteLine "<span class=""debugOutput""></span>"
        WriteLine "<script type=""text/javascript"" src=""lib/docScript.js""></script>"
        IndentDecrease
        WriteLine "</body>"
        IndentDecrease
        WriteLine "</html>"
    End Sub

    Sub Class_Terminate
        doc.Close
        Set subs = Nothing
        Set oMatch = Nothing
        Set oMatches = Nothing
    End Sub
End Class
