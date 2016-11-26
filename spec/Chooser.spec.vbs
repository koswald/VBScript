
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
        Dim pipe
        Dim pause : pause = 10

    .it "should open a browse for file window"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.file.vbs")
        ch.SetPatience 30 'in case the pc is very busy with other tasks, as at startup

        .AssertEqual ch.DialogHasOpened(ch.BFFileTitle), True

        ch.SetPatience 5 'restore default

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

        pause = 0

    .it "should open a browse for folder window"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.folder.vbs")

        .AssertEqual ch.DialogHasOpened(ch.BFFolderTitle), True

    .it "should return the path of a user-selected folder"

        'simulate user selecting the folder that the fixture opens to

        sh.AppActivate ch.BFFolderTitle
        WScript.Sleep pause
        sh.SendKeys "{ENTER}"

        .AssertEqual pipe.StdOut.ReadLine, sh.ExpandEnvironmentStrings("%tmp%")

    .it "should return an empty string if no folder was selected on Folder call"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.folder.vbs")

        'simulate the user canceling out of the dialog

        If ch.DialogHasOpened(ch.BFFolderTitle) Then
            sh.AppActivate ch.BFFolderTitle
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If

        .AssertEqual pipe.StdOut.ReadLine, ""

    .it "should return the title of a folder"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.title.vbs")

        'simulate user selecting the folder that the fixture opens to

        If ch.DialogHasOpened(ch.BFFolderTitle) Then
            sh.AppActivate ch.BFFolderTitle
            WScript.Sleep pause
            sh.SendKeys "{ENTER}"
        End If

        .AssertEqual pipe.StdOut.ReadLine, "Temp"

    .it "should return an empty string if no folder was selected on FolderTitle call"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.title.vbs")

        'simulate user cancelling

        If ch.DialogHasOpened(ch.BFFolderTitle) Then
            sh.AppActivate ch.BFFolderTitle
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If

        .AssertEqual pipe.StdOut.ReadLine, ""

    .it "should return a folder object of a user-selected folder"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.obj.vbs")

        'simulate user selecting the folder that the fixture opens to

        If ch.DialogHasOpened(ch.BFFolderTitle) Then
            sh.AppActivate ch.BFFolderTitle
            WScript.Sleep pause
            sh.SendKeys "{ENTER}"
        End If

        .AssertEqual pipe.StdOut.ReadLine, "Temp"

    .it "should return ""Nothing"" if no folder was selected"

        Set pipe = sh.Exec("cscript //nologo fixture/Chooser.obj2.vbs")

        If ch.DialogHasOpened(ch.BFFolderTitle) Then

            'simulate user escaping out of the dialog

            sh.AppActivate ch.BFFolderTitle
            WScript.Sleep pause
            sh.SendKeys "{ESC}"
        End If

        .AssertEqual pipe.StdOut.ReadLine, "Nothing"

End With

'breakdown

Set pipe = Nothing
Set sh = Nothing
