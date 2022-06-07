'Generate html and markdown documentation for VBScript code based on well-formed code comments.

'Usage Example
'<pre> With CreateObject( "VBScripting.Includer" )<br />     Execute .Read( "DocGenerator" )<br /> End With<br /> With New DocGenerator<br />     .SetTitle "VBScript Utility Classes Documentation"<br />     .SetDocName "VBScriptClasses"<br />     .SetFilesToDocument "*.vbs | *.wsf | *.wsc"<br />     .SetScriptFolder "..\class"<br />     .SetDocFolder "..\docs"<br />     .Generate<br />     .ViewMarkdown<br /> End With</pre>
'
'Example of well-formed comments before a Sub statement
' Note: A remark is required for Methods (Subs).
'
'<pre>'Method: SubName<br />'Parameters: varName, varType<br />'Remark: Details about the parameters.</pre>

'Example of well-formed comments before a Property or Function statement.
'Note: A Returns (or Return or Returns: or Return:) is required with a Property or Function.
'
'<pre>'Property: PropertyName<br />'Returns: a string<br />'Remark: A remark is not required for a Property or Function.</pre>

'Notes for the comment syntax at the beginning of a script

'Use a single quote ( ' ) for general comments <br />
'- use a single quote by itself for an empty line <br />
'- Wrap VBScript code with <code>pre</code> tags, separating multiple lines with &lt;br /&gt;. <br />
'- Wrap other code with <code> code</code> tags, separating multiple lines with &lt;br /&gt;. <br />
'
'Use three single quotes for remarks that should not appear in the documentation <br />
'
'Use four single quotes ( '''' ), if the script doesn't contain a class statement, to separate the general comments at the beginning of the file from the rest of the file.
'
'For some characters to render correctly, they may need to be replaced by escape codes, even when used within &#60;code&#62; or &#60;pre&#62; tags:
' for &#124; use &#38;#124; (vertical bar)
' for &#60; use &#38;#60; (less than)
' for &#62; use &#38;#62; (greater than)
' for &#92; use &#38;#92; (backslash)
' for &#38; use &#38;#38; (ampersand)
'For other characters,  <code>examples\HTML Escape Codes.hta</code> can be used to generate an escape code that works with both of the generated files: Markdown and HTML. The numerical portion of the escape code is returned by the VBScript function Asc.
'
'Visual Studio and VS Code extensions may render Markdown files differently than Git-Flavored Markdown.
'
'Issues:
'- Introductory comments at the beginning of a class file should be followed by a line containing a single quote character, or else the markdown table may not render correctly.
'

Class DocGenerator

    Private fs 'VBSFileSystem project object
    Private rf 'RegExFunctions project object
    Private outputStreamer, inputStreamer 'TextStreamer project objects
    Private sh 'WScript.Shell COM object
    Private fso 'Scripting.FileSystemObject COM object
    Private script, doc 'input and output text streams
    Private md 'output text stream
    Private File 'fso File object
    Private re 'RegExp object

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
        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "TextStreamer" )
            Execute .Read( "VBSFileSystem" )
            Execute .Read( "RegExFunctions" )
            ExecuteGlobal .Read( "EscapeMd" )
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
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
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
        Colorize = True
    End Sub

    Private Sub InitializeDocFiles
        docFile = docFolder & "\" & docName & ".html"
        outputStreamer.SetFile docFile
        Set doc = outputStreamer.Open
        Set md = fso.OpenTextFile(docFolder & "\" & docName & ".md", 2, True)
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
    'Remark: Required. Must be set before calling the Generate method. Specifies the name of the documentation file. Do not include the extension name.
    Sub SetDocName(newDocName) : docName = newDocName : End Sub

    'Method SetTitle
    'Parameter: a string
    'Remark: Required. Must be set before calling the Generate method. Sets the title for the documentation.
    Sub SetTitle(newDocTitle) : docTitle = newDocTitle : End Sub

    'Method SetFilesToDocument
    'Parameter: wildcard(s)
    'Remark: Specifies which files to document. Optional. Default is <strong> *.vbs </strong>. Separate multiple wildcards with &#124;
    Sub SetFilesToDocument(newFilesToDocument) : filesToDocument = rf.Pattern(newFilesToDocument) : End Sub

    Private Sub ValidateConfiguration
        Dim msg
        msg = "A title for the document must be set using SetTitle."
        If "" = docTitle Then Err.Raise 449, fs.SName, msg

        msg = "An existing folder containing the scripts to document must be specified with SetScriptFolder."
        If Not fso.FolderExists(scriptFolder) Then Err.Raise 449, fs.SName, msg

        msg = "An existing folder to contain the document must be specified with SetDocFolder."
        If Not fso.FolderExists(docFolder) Then Err.Raise 449, fs.SName, msg

        msg = "The name of the doc file must be specified with SetDocName."
        If "" = docName Then Err.Raise 449, fs.SName, msg
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
    'Remark: Open the html document in the default viewer. Same as ViewHtml.
    Sub View
        sh.Run """" & docFile & """"
    End Sub

    'Method ViewHtml
    'Remark: Open the html document in the default viewer. Same as View method.
    Sub ViewHtml : View : End Sub

    'Method ViewMarkdown
    'Remark: Open the markdown document in the default viewer.
    Sub ViewMarkdown
        sh.Run """" & docFolder & "\" & docName & ".md"""
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
    'That is, look for comments intended to be included in the documentation: Method or Property, Parameters, Returns, Remarks; also look for routines: Sub, Property, Function
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
        Dim baseName
        baseName = fso.GetBaseName(File.Name) 
        doc.WriteLine ""
        WriteLine "<div>"
        IndentIncrease
        WriteLine "<a id=""" & LCase(baseName) & """></a>"
        WriteLine "<h2 class=""heading"" id=" & id & ">" & baseName & "</h2>"
        WriteLine "<div class=""detail"">"
        md.WriteLine ""
        md.WriteLine "## "& baseName
        md.WriteLine
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
        WriteLine "<th>Member type</th>"
        WriteLine "<th>Name</th>"
        WriteLine "<th>Parameter(s)</th>"
        WriteLine "<th>Return value</th>"
        WriteLine "<th>Comment</th>"
        IndentDecrease
        WriteLine "</tr>"
        md.WriteLine "| Member type | Name | Parameter | Returns | Comment |"
        md.WriteLine "| :---------- | :--- | :-------- | :------ | :------ |"
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
        Else WriteLine "<p>" & generalContent & "</p>"
        End If
        IndentDecrease
        md.WriteLine GetColorizedOrGetNowrap(generalContent)
    End Sub

    Function GetColorizedOrGetNowrap(markup)
        If Not CBool(InStr(markup, "<pre>")) Then
            GetColorizedOrGetNowrap = markup & "  "
        ElseIf colorize_ Then
            GetColorizedOrGetNowrap = GetColorized(markup)
        Else GetColorizedOrGetNowrap = GetNowrap(markup)
        End If
    End Function

    Function GetNowrap(markup)
        Dim lines : lines = markup
        lines = Replace(lines, "<br />", "<br/>")
        lines = Replace(lines, " ", "�") 'Alt+0160 = non-breaking space
        lines = Replace(lines, "<pre>", "<pre><code style='white-space: nowrap;'>")
        lines = Replace(lines, "</pre>", "</code></pre>")
        GetNowrap = lines
    End Function

    Function GetColorized(markup)
        Dim lines : lines = markup
        lines = Replace(lines, "<pre>", "```vb" & vbCrLf)
        lines = Replace(lines, "<br />", vbCrLf)
        lines = Replace(lines, "</pre>", vbCrLf & "```")
        lines = Replace(lines, "&lt;", "<")
        lines = Replace(lines, "&gt;", ">")
        GetColorized = lines
    End Function

    'Property Colorize
    'Parameters: boolean
    'Returns: boolean
    'Remarks: Gets or sets whether &lt;pre&gt; code blocks (assumed to be VBScript) in the markdown document are colorized. If False (experimental, with Git Flavored Markdown), the code lines should not wrap. Default is True.
    Property Get Colorize : Colorize = colorize_ : End Property
    Property Let Colorize(value) : colorize_ = value : End Property
    Private colorize_

    'pre content has been deprecated, that is, preceeding code with two single quotes ('') has been deprecated in favor of wrapping VBScript code with "pre" tags and other code with "code" tags. See the class introductory comments.
    Private Sub WritePreContentToDoc
        If postClassStatement = status Then Exit Sub
        If Not ScriptHeaderWritten Then
            WriteScriptHeader
        End If
        IndentIncrease
        WriteLine "<pre>" & preContent & "</pre>"
        IndentDecrease
        'md.WriteLine "<pre>" & preContent & "</pre>"
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
        md.WriteLine "| " & routineType & " | " & routineName & " | " & parametersContent & " | " & returnsContent & " | " & remarksContent & " |"
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
        If CBool(Len(methodContent)) And CBool(Len(propertyContent)) Then RaiseContentError msg

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
            If "" = remarksContent Then
                RaiseContentError msg & " (Remarks content is empty.)"
            End If
            If Len(returnsContent) Then
                RaiseContentError msg & " (Returns content is not empty.)"
            Else returnsContent = "N/A"
            End If
            If "" = parametersContent Then
                parametersContent = "None"
            End If

        ElseIf property_ =  routineType Then

            msg = "Properties may have parameters; may have Remarks; must have Returns. NOTE: Property Set and Property Let do not require a return value but still must have a 'Return or 'Returns comment."
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
        WriteLine "<h1>" & docTitle & "</h1>"
        doc.WriteLine ""
        IndentIncrease

        md.WriteLine "# " & docTitle
        md.WriteLine ""
        md.WriteLine "## Contents"
        md.WriteLine ""
        Dim baseName
        For Each File In fso.GetFolder(scriptFolder).Files
            re.Pattern = filesToDocument
            baseName = fso.GetBaseName(File.Name)
            If re.Test(File.Name) And Not CBool(InStr(baseName, ".")) Then
                md.WriteLine "[" & baseName & "](#" & LCase(baseName) & ")  "
            End If
        Next
    End Sub

    Private Sub WriteBottomSection
        WriteLine "<p> <em> See also the <a target=""_blank"" href=""https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/t0aew7h6(v=vs.84)""> VBScript docs </a> </em> </p>"
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
