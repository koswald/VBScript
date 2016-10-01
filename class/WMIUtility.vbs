
Class WMIUtility 'usage examples of the Windows Management Instrumentation object
    
    Private computer, select_, all, from, where 'select is a keyword, so the variable was named select_

    Sub Class_Initialize
         SetPC(localPC)
         select_ = "Select " : all = " * " : from = " from " : where = " where "
    End Sub

    Sub SetPC(newPC) : computer = newPC : End Sub
    Private Property Get localPC : localPC = "." : End Property

    'Terminate a process by id

    Function TerminateProcessById(id)
        Dim scrubId : scrubId = Scrub(id)
        Dim process
        For Each process in GetProcessesWhere("ProcessID = '" & scrubId & "'")
            process.Terminate()
            TerminateProcessById = True
            Exit Function
        Next
        TerminateProcessById = False
    End Function

    'Terminate a process by id and name

    Function TerminateProcessByIdAndName(id, name)
        Dim scrubId : scrubId = Scrub(id)
        Dim scrubName : scrubName = Scrub(name)
        Dim process
        For Each process in GetProcessesWhere("ProcessID = '" & scrubId & "' and Name = '" & scrubName & "'")
            process.Terminate()
            TerminateProcessByIdAndName = True
            Exit Function
        Next
        TerminateProcessByIdAndName = False
    End Function

    'Return an array of ProcessId values for processes with specified name, e.g. notepad.exe

    Function GetProcessIDsByName(pName)
        Dim s : s = ""
        Dim scrubName : scrubName = Scrub(pName)
        Dim process
        For Each process in GetProcessesWhere("Name = '" & scrubName & "'")
            s = s & " " & process.ProcessID
        Next
        GetProcessIDsByName = split(Trim(s))
    End Function

    'Return process-names array with names like _____ e.g. %svc%

    'parameter "string_" typically has one or more % wildcard characters
    'to filter out duplicates, use filteredArray = VBSArrays.Uniques(unfilteredArray)

    Function GetProcessesWithNamesLike(string_) 
        Dim s : s = ""
        Dim scrubString : scrubString = Scrub(string_)
        Dim process
        For Each process in GetProcessesWhere("Name like '" & scrubString & "'")
            s = s & " " & process.Name
        Next
        GetProcessesWithNamesLike = split(Trim(s))
    End Function

    'Scrub parameters before query

    Property Get Scrub(param)
        Dim s : s = param
        Dim removes : removes = Array("=", " ", ";", "'", "\", "/", ":", "*", "?", """", "<", ">", "|", "%20")

        For i = 0 To UBound(removes)
            s = Replace(s, removes(i), "")
        Next
        Scrub = s
    End Property

    Private Function GetProcessesWhere(condition)
        Set GetProcessesWhere = GetResults(select_ & all & from & Win32_Process & where & condition)
    End Function

    Property Get Win32_Process : Win32_Process = "Win32_Process" : End Property

    Private Function GetResults(query)
        Set GetResults = GetObject(GetWmiToken).ExecQuery(query)
    End Function

    Private Property Get wmiToken1 : wmiToken1 = "winmgmts:\\" : End Property
    Private Property Get wmiToken2 : wmiToken2 = "winmgmts:{impersonationLevel=impersonate}!\\" : End Property
    Property Get GetWmiToken : GetWmiToken = wmiToken1 & computer & "\root\cimv2" : End Property

    'Return a collection of partition objects, each with the following methods: 

    'Caption, Name, DiskIndex, Index, PrimaryPartition, Bootable, BootPartition, Description, Type, Size, 
    'StartingOffset, BlockSize, DeviceID, Access, Availability, ErrorMethodology, HiddenSectors, Purpose, Status

    Function partitions 
        Set partitions = GetResults(select_ & all & from & Win32_DiskPartition)
    End Function

    Property Get Win32_DiskPartition : Win32_DiskPartition = "Win32_DiskPartition" : End Property

    'Return a collection of disk objects, each with these methods: .FileSystem, .DeviceID

    Function disks
        Set disks = GetResults(select_ & all & from & Win32_LogicalDisk)
    End Function

    Property Get Win32_LogicalDisk : Win32_LogicalDisk = "Win32_LogicalDisk" : End Property

    'Return an object with these methods: .Architecture, .Description

    Function cpu 
        Dim process
        For Each process in GetResults(select_ & all & from & Win32_Processor)
            Set cpu = process : Exit For
        Next
    End Function

    Property Get Win32_Processor : Win32_Processor = "Win32_Processor" : End Property

    'Return an OS object with these methods:

    '.Name, .Version, .Manufacturer, .WindowsDirectory, .Locale, .FreePhysicalMemory, 
    '.TotalVirtualMemorySize, .FreeVirtualMemory, .SizeStoredInPagingFiles

    Function os
        Dim process
        For Each process in GetResults(select_ & all & from & Win32_OperatingSystem)
            Set os = process : Exit For
        Next
    End Function

    Property Get Win32_OperatingSystem : Win32_OperatingSystem = "Win32_OperatingSystem" : End Property

    'Return a PC object with these methods: 
    
    '.Name, .Manufacturer, .Model, .CurrentTimeZone, .TotalPhysicalMemory

    Function pc 
        Dim process
        For Each process in GetResults(select_ & all & from & Win32_ComputerSystem)
            Set pc = process : Exit For
        Next
    End Function

    Property Get Win32_ComputerSystem : Win32_ComputerSystem = "Win32_ComputerSystem" : End Property

    'Return a BIOS object with method(s): .Version

    Function Bios 
        Dim process
        For Each process in GetResults(select_ & all & from & Win32_Bios)
            Set Bios = process : Exit For
        Next
    End Function

    Property Get Win32_Bios : Win32_Bios = "Win32_Bios" : End Property

End Class
