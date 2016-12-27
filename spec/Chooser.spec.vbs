
'test Chooser.vbs

With CreateObject("includer")
    Execute(.read("Chooser"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    .describe "Chooser class"
        Dim ch : Set ch = New Chooser

    'setup
        Dim sh : Set sh = CreateObject("WScript.Shell")
        Dim warning : Set warning = sh.Exec("wscript fixture/Chooser.warn.vbs")
        Dim pipe, pause

        'rig for busy CPU, as at startup
        pause = 60
        ch.SetPatience 30

    .it "should open a browse for file window"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.file.vbs")
        .AssertEqual ch.DialogHasOpened(ch.BFFileTitle), True

    .it "should return the path of a user-selected file"
        'simulate user selecting this script file
        sh.AppActivate ch.BFFileTitle
        WScript.Sleep pause
        sh.SendKeys "%n" 'Alt N to focus on file name field
        WScript.Sleep pause
        sh.SendKeys WScript.ScriptFullName
        WScript.Sleep pause
        sh.SendKeys "{ENTER}"
        .AssertEqual pipe.StdOut.ReadLine, WScript.ScriptFullName

    .it "should return an empty string if no file was selected"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.file.vbs")
        If ch.DialogHasOpened(ch.BFFileTitle) Then
            'simulate user escaping out of the dialog
            sh.AppActivate ch.BFFileTitle
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If
        .AssertEqual pipe.StdOut.ReadLine, ""

    'rig for not-so-busy CPU
        ch.SetPatience 5
        pause = 0

    .it "should open a browse for folder window"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.folder.vbs")
        .AssertEqual ch.DialogHasOpened(pipe), True

    .it "should return the path of a user-selected folder"
        'simulate user selecting the folder that the fixture opens to
        sh.AppActivate pipe.ProcessID
        WScript.Sleep pause
        sh.SendKeys "{ENTER}"
        .AssertEqual pipe.StdOut.ReadLine, sh.ExpandEnvironmentStrings("%tmp%")

    .it "should return an empty string if no folder was selected: .Folder"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.folder.vbs")
        'simulate the user canceling out of the dialog
        If ch.DialogHasOpened(pipe) Then
            sh.AppActivate pipe.ProcessID
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If
        .AssertEqual pipe.StdOut.ReadLine, ""

    .it "should return the title of a folder"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.title.vbs")
        'simulate user selecting the folder that the fixture opens to
        If ch.DialogHasOpened(pipe) Then
            sh.AppActivate pipe.ProcessID
            WScript.Sleep pause
            sh.SendKeys "{ENTER}"
        End If
        .AssertEqual pipe.StdOut.ReadLine, "Temp"

    .it "should return an empty string if no folder was selected: .FolderTitle"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.title.vbs")
        'simulate user cancelling
        If ch.DialogHasOpened(pipe) Then
            sh.AppActivate pipe.ProcessID
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If
        .AssertEqual pipe.StdOut.ReadLine, ""

    .it "should return a folder object of a user-selected folder"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.obj.vbs")
        'simulate user selecting the folder that the fixture opens to
        If ch.DialogHasOpened(pipe) Then
            sh.AppActivate pipe.ProcessID
            WScript.Sleep pause
            sh.SendKeys "{ENTER}"
        End If
        .AssertEqual pipe.StdOut.ReadLine, "Temp"

    .it "should return ""Nothing"" if no folder was selected"
        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.obj2.vbs")
        If ch.DialogHasOpened(pipe) Then
            'simulate user escaping out of the dialog
            sh.AppActivate pipe.ProcessID
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If
        .AssertEqual pipe.StdOut.ReadLine, "Nothing"

End With

'breakdown

Set pipe = Nothing
Set sh = Nothing
warning.Terminate
Set warning = Nothing
