'Script for StartItems.hta

Option Explicit
Dim si 'StartupItems object
Dim enableWowForHKCU 'boolean. See StartItems.configure for explanation
Dim privilegesAreElevated 'boolean
Dim hiveSelector, keySelector 'select elements
Dim selectDiv, tableDiv 'div elements (containers)
Dim nameInput, valueInput 'input elements for creating a new startup item
Dim feedback 'p element for message display
Dim application 'application element
Dim OsIs64Bit 'boolean
Dim elButton 'Elevate Privileges button
Dim tmButton 'Open Task Manager button
Dim suButton 'Open Settings|Apps|Startup button

Sub Window_OnLoad
    Dim includer 'VBScripting.Includer object
    Dim width, height, xPos, yPos 'window size and position in % of screen
    Dim pxWidth, pxHeight 'calculated window dimensions in pixels

    Set includer = CreateObject( "VBScripting.Includer" )
    Execute includer.Read( "StartupItems" )
    Set si = New StartupItems

    Execute includer.Read( "Configurer" )
    With New Configurer
        If .Exists( "enableHKCUWow" ) Then
            enableWowForHKCU = .Item( "enableHKCUWow" )
        Else enableWowForHKCU = False
        End If
        If .Exists( "width" ) Then
            width = .Item( "width" )
        Else width = 65
        End If
        If .Exists( "height" ) Then
            height = .Item( "height" )
        Else height = 52
        End If
        If .Exists( "xPos" ) Then
            xPos = .Item( "xPos" )
        Else xPos = 50
        End If
        If .Exists( "yPos" ) Then
            yPos = .Item( "yPos" )
        Else yPos = 50
        End If
    End With

    With self.Screen
        pxWidth = .AvailWidth * width * .01
        pxHeight = .AvailHeight * height * .01
        self.ResizeTo pxWidth, pxHeight
        self.MoveTo _
            (.AvailWidth - pxWidth) * xPos * .01005, _
            (.AvailHeight - pxHeight) * yPos * .0102
    End With

    With CreateObject( "VBScripting.Admin" )
        privilegesAreElevated = .PrivilegesAreElevated
    End With

    Set application = document.getElementsByTagName( "application" )(0)
    document.Title = application.applicationName
    If privilegesAreElevated Then
        document.Title = document.Title & " - Administrator"
    End If

    Execute includer.Read( "WoWChecker" )
    With New WoWChecker
        OsIs64Bit = .OsIs64Bit
    End With

    DrawSelectDiv
    Set tableDiv = document.createElement( "div" )
    document.body.insertBefore tableDiv
    DrawTableDiv
    hiveSelector.selectedIndex = 0
    Set includer = Nothing
End Sub

Sub Window_OnUnload
    Set si = Nothing
End Sub

Sub DrawSelectDiv
    Dim hkcuOption, hklmOption 'select element options

    Set selectDiv = document.createElement( "div" )
    document.body.insertBefore selectDiv
    selectDiv.style.marginBottom = "15px"
    Set hiveSelector = document.createElement( "select" )
    Set keySelector = document.createElement( "select" )
    selectDiv.insertBefore hiveSelector
    selectDiv.insertBefore document.createTextNode(" \ ")
    selectDiv.insertBefore keySelector
    Set hkcuOption = document.createElement( "option" )
    Set hklmOption = document.createElement( "option" )
    hkcuOption.value = si.HKCU
    hklmOption.value = si.HKLM
    hkcuOption.innerHTML = "HKEY_CURRENT_USER"
    hklmOption.innerHTML = "HKEY_LOCAL_MACHINE"
    hiveSelector.insertBefore hkcuOption
    hiveSelector.insertBefore hklmOption
    hiveSelector.OnChange = GetRef( "RootChanged" )
    DrawKeySelector
End Sub

Sub DrawKeySelector
    Dim wowOption, nonWowOption 'select element options

    keySelector.innerHTML = ""
    Set nonWowOption = document.createElement( "option" )
    Set wowOption = document.createElement( "option" )
    nonWowOption.value = si.StandardBranch
    wowOption.value = si.WoWBranch
    nonWowOption.innerHTML = si.StandardBranch
    wowOption.innerHTML = si.WoWBranch
    keySelector.insertBefore nonWowOption
    If OsIs64Bit _
    And (enableWowForHKCU Or si.HKLM = si.Root) Then
        keySelector.insertBefore wowOption
    End If
    keySelector.OnChange = GetRef( "KeyChanged" )
End Sub

Sub RootChanged
    si.Root = CLng(hiveSelector.options(hiveSelector.selectedIndex).value)
    keySelector.selectedIndex = 0
    si.Key = si.StandardBranch
    DrawKeySelector
    DrawTableDiv
End Sub

Sub KeyChanged
    si.Key = keySelector.options(keySelector.selectedIndex).value
    DrawTableDiv
End Sub

Sub DrawTableDiv
    Dim items 'collection of NameValue objects: each object contains the name and value of a startup item from the registry for the selected hive and key.
    Dim tbl 'html element: table
    Dim rowIndex 'integer: loop iterator
    Dim row 'html element: table row
    Dim nameCell 'html element: row cell
    Dim valueCell 'html element: row cell
    Dim buttonCell 'html element: row cell
    Dim button 'html element: button

    tableDiv.innerHTML = ""
    items = si.Items
    Set tbl = document.createElement( "table" )
    tableDiv.insertBefore tbl
    For rowIndex = -1 To UBound(items) + 1
        Set row = tbl.insertRow(-1)
        Set nameCell = row.insertCell(-1)
        Set valueCell = row.insertCell(-1)
        Set buttonCell = row.insertCell(-1)
        Set button = document.createElement( "input" )
        button.type = "button"

        'create the header row
        If -1 = rowIndex Then
            nameCell.innerHTML = "Name"
            valueCell.innerHTML = "Value"
            row.style.fontWeight = "bold"

        'show registry data for the selected hive (root) and key
        ElseIf Not UBound(items) + 1 = rowIndex Then
            nameCell.innerHTML = items(rowIndex).Name
            valueCell.innerHTML = items(rowIndex).Value
            buttonCell.insertBefore button
            button.value = "   Remove   "
            Set button.OnClick = GetRef( "RemoveItem" )

            'even-numbered row: darken background--two cells only
            If Not CBool(rowIndex mod 2) Then
                nameCell.style.backgroundColor = "#eee"
                valueCell.style.backgroundColor = "#eee"
            End If

        'for the last row, create the input fields for creating a new entry
        Else Set nameInput = document.createElement( "input" )
            Set valueInput = document.createElement( "input" )
            nameInput.style.width = "100%"
            valueInput.style.width = "99%"
            nameCell.insertBefore nameInput
            valueCell.insertBefore valueInput
            buttonCell.insertBefore button
            button.value = "  Add  "
            Set button.OnClick = GetRef( "SaveItem" )
        End If

        valueCell.style.paddingLeft = "20px"
        buttonCell.style.paddingLeft = "20px"
        button.style.width = "100%"
    Next

    With document.body.style
        .cursor = "default"
        .fontFamily = "sans-serif"
        .fontSize = "13"
    End With
    With tbl.style
        .borderCollapse = "collapse"
        .marginTop = "15px"
        .marginRight = "15px"
    End With
    Set feedback = document.createElement( "p" )
    tableDiv.insertBefore feedback
    Set elButton = document.createElement( "input" )
    elButton.type = "button"
    elButton.value = "Elevate privileges"
    Set elButton.onclick = GetRef( "Elevate" )
    tableDiv.insertBefore elButton
    tableDiv.insertBefore document.createTextNode("  ")
    Set tmButton = document.createElement( "input" )
    tmButton.type = "button"
    tmButton.value = "Open Task Manager"
    Set tmButton.onclick = GetRef( "OpenTaskMgr" )
    tableDiv.insertBefore tmButton
    tableDiv.insertBefore document.createTextNode("  ")
    Set suButton = document.createElement( "input" )
    suButton.type = "button"
    suButton.value = "Open Settings | Apps | Startup"
    Set suButton.onclick = GetRef( "OpenStartupAppsSettings")
    tableDiv.insertBefore suButton
    CheckPrivileges
End Sub

Sub OpenTaskMgr
    si.OpenTaskMgr
End Sub

Sub OpenStartupAppsSettings
    With CreateObject( "WScript.Shell" )
        .Run "ms-settings:startupapps"
    End With
End Sub

Sub Elevate
    With CreateObject( "Shell.Application" )
        .ShellExecute "mshta", application.CommandLine,, "runas"
    End With
    self.close
End Sub

Sub SaveItem
    si.CreateItem nameInput.value, valueInput.value
    DrawTableDiv
End Sub

Sub RemoveItem
    Dim items 'collection of NameValue objects
    Dim inputs 'collection all of the input elements in the document
    Dim itemIndex 'index of the item to remove in the items collection
    Dim inputIndex 'index of the item to remove in the inputs collection
    Dim msg, settings, caption 'MsgBox arguments

    items = si.Items

    'get the index of the items array
    Set inputs = document.getElementsByTagName( "input" )
    For inputIndex = 0 To inputs.length - 1
        If window.event.srcElement Is inputs(inputIndex) Then
            itemIndex = inputIndex
            Exit For
        End If
    Next

    'optout
    msg = "Do you want to remove this item?" & vbLf & vbLf & _
        items(itemIndex).Name & vbLf & items(itemIndex).Value
    settings = vbOKCancel + vbExclamation + vbDefaultButton2
    caption = application.applicationName
    If vbCancel = MsgBox(msg, settings, caption) Then
        Exit Sub
    End If

    'remove the item
    si.DeleteItem items(itemIndex).Name
    DrawTableDiv
End Sub

Sub CheckPrivileges
    Dim currentHive
    currentHive = CLng(hiveSelector.options(hiveSelector.selectedIndex).value)
    If Not privilegesAreElevated And si.HKLM = currentHive Then
        InputsEnabled False
        feedback.innerHTML = "Elevated privileges are required to edit HKEY_LOCAL_MACHINE."
    Else InputsEnabled True
        feedback.innerHTML = ""
    End If
    If privilegesAreElevated Then
        elButton.disabled = True
    Else elButton.disabled = False
    End If
    tmButton.disabled = False
End Sub

Sub InputsEnabled(enabling)
    Dim inputs : Set inputs = document.getElementsByTagName( "input" )
    Dim inputIndex
    For inputIndex = 0 To inputs.length - 1
        inputs(inputIndex).disabled = Not enabling
    Next
End Sub
