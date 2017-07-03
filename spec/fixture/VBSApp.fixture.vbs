'test fixture for ..\VBSApp.spec.vbs

'initialize
Option Explicit
Const debugging = False
On Error Resume Next
If debugging Then On Error Goto 0
With CreateObject("includer")
    Execute(.read("VBSApp"))
    Execute(.read("VBSTimer"))
    Dim upperLimit, lowerLimit, testSleep
    Execute(.read("..\spec\VBSApp.spec.config"))
End With
Dim app : Set app = New VBSApp
Dim tmr : Set tmr = New VBSTimer
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim base : base = "fixture\VBSApp.fixture."
Const ForWriting = 2
Const CreateNew = True
Dim stream : Set stream = fso.OpenTextFile(base & "VbsOut.txt", ForWriting, CreateNew)

'output selected command-line argument
'Dim args : args = app.GetArgs
stream.WriteLine app.GetArg(1) 'args(1)

'output the command-line string
stream.WriteLine app.GetArgsString

'output the argument count
stream.WriteLine app.GetArgsCount

'output the filespec
stream.WriteLine app.GetFullName

'output the file name
stream.WriteLine app.GetFileName

'output the base file name
stream.WriteLine app.GetBaseName

'output the file extension name
stream.WriteLine app.GetExtensionName

'output the host .exe
stream.WriteLine app.GetExe

'attempt to invoke the Sleep method
On Error Resume Next
    app.Sleep 1
    If Err Then
        stream.WriteLine Err.Description
    Else
        stream.WriteLine Err 'Err.Number
    End If
On Error Goto 0

'output the actual sleep duration
tmr.Reset
app.Sleep testSleep
stream.WriteLine tmr.Split

'cleanup
stream.Close
Set stream = Nothing
Set fso = Nothing
