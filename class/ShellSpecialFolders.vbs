' ShellSpecialFolders class
'
' Adapted from <a href="https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants"> ShellSpecialFolderConstants enumeration (shldisp.h)</a>: Specifies unique, system-independent values that identify special folders. These folders are frequently used by applications but which may not have the same name or location on any given system. For example, the system folder can be "C:\Windows" on one system and "C:\Winnt" on another.
'
Class ShellSpecialFolders

    Private sa 'Shell.Application object

    Sub Class_Initialize
        Set sa = CreateObject("Shell.Application")
    End Sub

    'Property Path
    'Returns a path
    'Parameter: an integer
    'Remark: Returns the path to a special folder. The parameter is one of the ssf constants. This path is suitable for navigating in Windows Explorer. For ssfCONTROLS, ssfPRINTERS, ssfBITBUCKET, ssfDRIVES, and ssfNETWORK, the return value looks different than a typical path: for ssfDrives it is ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}.
    Public Property Get Path(ssfConstant)
        Path = sa.NameSpace(ssfConstant).Self.Path
    End Property

    'Property AllConstants
    'Returns an array
    'Remarks: Returns an array with all of the ssf constants (integers).
    Property Get AllConstants
        AllConstants = Array( _
            me.ssfDESKTOP, _
            me.ssfPROGRAMS, _
            me.ssfCONTROLS, _
            me.ssfPRINTERS, _
            me.ssfPERSONAL, _
            me.ssfFAVORITES, _
            me.ssfSTARTUP, _
            me.ssfRECENT, _
            me.ssfSENDTO, _
            me.ssfBITBUCKET, _
            me.ssfSTARTMENU, _
            me.ssfDESKTOPDIRECTORY, _
            me.ssfDRIVES, _
            me.ssfNETWORK, _
            me.ssfNETHOOD, _
            me.ssfFONTS, _
            me.ssfTEMPLATES, _
            me.ssfCOMMONSTARTMENU, _
            me.ssfCOMMONPROGRAMS, _
            me.ssfCOMMONSTARTUP, _
            me.ssfCOMMONDESKTOPDIR, _
            me.ssfAPPDATA, _
            me.ssfPRINTHOOD, _
            me.ssfLOCALAPPDATA, _
            me.ssfALTSTARTUP, _
            me.ssfCOMMONALTSTARTUP, _
            me.ssfCOMMONFAVORITES, _
            me.ssfINTERNETCACHE, _
            me.ssfCOOKIES, _
            me.ssfHISTORY, _
            me.ssfCOMMONAPPDATA, _
            me.ssfWINDOWS, _
            me.ssfSYSTEM, _
            me.ssfPROGRAMFILES, _
            me.ssfMYPICTURES, _
            me.ssfPROFILE, _
            me.ssfSYSTEMx86, _
            me.ssfPROGRAMFILESx86 )
    End Property

    'Property AllPaths
    'Returns an array
    'Remarks: Returns an array with all of the ssf paths.
    Property Get AllPaths
        AllPaths = Array( _
            me.Path( me.ssfDESKTOP ), _
            me.Path( me.ssfPROGRAMS ), _
            me.Path( me.ssfCONTROLS ), _
            me.Path( me.ssfPRINTERS ), _
            me.Path( me.ssfPERSONAL ), _
            me.Path( me.ssfFAVORITES ), _
            me.Path( me.ssfSTARTUP ), _
            me.Path( me.ssfRECENT ), _
            me.Path( me.ssfSENDTO ), _
            me.Path( me.ssfBITBUCKET ), _
            me.Path( me.ssfSTARTMENU ), _
            me.Path( me.ssfDESKTOPDIRECTORY ), _
            me.Path( me.ssfDRIVES ), _
            me.Path( me.ssfNETWORK ), _
            me.Path( me.ssfNETHOOD ), _
            me.Path( me.ssfFONTS ), _
            me.Path( me.ssfTEMPLATES ), _
            me.Path( me.ssfCOMMONSTARTMENU ), _
            me.Path( me.ssfCOMMONPROGRAMS ), _
            me.Path( me.ssfCOMMONSTARTUP ), _
            me.Path( me.ssfCOMMONDESKTOPDIR ), _
            me.Path( me.ssfAPPDATA ), _
            me.Path( me.ssfPRINTHOOD ), _
            me.Path( me.ssfLOCALAPPDATA ), _
            me.Path( me.ssfALTSTARTUP ), _
            me.Path( me.ssfCOMMONALTSTARTUP ), _
            me.Path( me.ssfCOMMONFAVORITES ), _
            me.Path( me.ssfINTERNETCACHE ), _
            me.Path( me.ssfCOOKIES ), _
            me.Path( me.ssfHISTORY ), _
            me.Path( me.ssfCOMMONAPPDATA ), _
            me.Path( me.ssfWINDOWS ), _
            me.Path( me.ssfSYSTEM ), _
            me.Path( me.ssfPROGRAMFILES ), _
            me.Path( me.ssfMYPICTURES ), _
            me.Path( me.ssfPROFILE ), _
            me.Path( me.ssfSYSTEMx86 ), _
            me.Path( me.ssfPROGRAMFILESx86 ))
    End Property

    'Property ssfDESKTOP
    'Returns 0
    Public Property Get ssfDESKTOP : ssfDESKTOP = 0 : End Property

    'Property ssfPROGRAMS
    'Returns &h2
    Public Property Get ssfPROGRAMS : ssfPROGRAMS = &h2 : End Property

    'Property ssfCONTROLS
    'Returns &h3
    'Remarks: Virtual folder that contains icons for the Control Panel applications.
    Public Property Get ssfCONTROLS : ssfCONTROLS = &h3 : End Property

    'Property ssfPRINTERS
    'Returns &h4
    'Remarks: Virtual folder that contains installed printers.
    Public Property Get ssfPRINTERS : ssfPRINTERS = &h4 : End Property

    'Property ssfPERSONAL
    'Returns &h5
    'Remarks: File system directory that serves as a common repository for a user's documents. A typical path is C:\Users&#92;<em>username</em>\Documents.
    Public Property Get ssfPERSONAL : ssfPERSONAL = &h5 : End Property

    'Property ssfFAVORITES
    'Returns &h6
    Public Property Get ssfFAVORITES : ssfFAVORITES = &h6 : End Property

    'Property ssfSTARTUP
    'Returns &h7
    Public Property Get ssfSTARTUP : ssfSTARTUP = &h7 : End Property

    'Property ssfRECENT
    'Returns &h8
    Public Property Get ssfRECENT : ssfRECENT = &h8 : End Property

    'Property ssfSENDTO
    'Returns &h9
    Public Property Get ssfSENDTO : ssfSENDTO = &h9 : End Property

    'Property ssfBITBUCKET
    'Returns &ha
    'Remarks: According to the <a href="https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants"> docs</a>: "Virtual folder that contains the objects in the user's Recycle Bin."
    Public Property Get ssfBITBUCKET : ssfBITBUCKET = &ha : End Property

    'Property ssfSTARTMENU
    'Returns &hb
    Public Property Get ssfSTARTMENU : ssfSTARTMENU = &hb : End Property

    'Property ssfDESKTOPDIRECTORY
    'Returns &h10
    'Remarks: According to the <a href="https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants"> docs</a>: "File system directory used to physically store the file objects that are displayed on the desktop. It is not to be confused with the desktop folder itself, which is a virtual folder." A typical path is C:\Users&#92;<em>username</em>\Desktop.
    Public Property Get ssfDESKTOPDIRECTORY : ssfDESKTOPDIRECTORY = &h10 : End Property

    'Property ssfDRIVES
    'Returns &h11
    'Remarks: My Computer—the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel. This folder can also contain mapped network drives.
    Public Property Get ssfDRIVES : ssfDRIVES = &h11 : End Property

    'Property ssfNETWORK
    'Returns &h12
    'Remarks: Network Neighborhood—the virtual folder that represents the root of the network namespace hierarchy.
    Public Property Get ssfNETWORK : ssfNETWORK = &h12 : End Property

    'Property ssfNETHOOD
    'Returns &h13
    'Remarks: A file system folder that contains any link objects in the My Network Places virtual folder. It is not the same as ssfNETWORK, which represents the network namespace root. A typical path is C:\Users&#92;<em>username</em>\AppData\Roaming\Microsoft\Windows\Network Shortcuts.
    Public Property Get ssfNETHOOD : ssfNETHOOD = &h13 : End Property

    'Property ssfFONTS
    'Returns &h14
    Public Property Get ssfFONTS : ssfFONTS = &h14 : End Property

    'Property ssfTEMPLATES
    'Returns &h15
    Public Property Get ssfTEMPLATES : ssfTEMPLATES = &h15 : End Property

    'Property ssfCOMMONSTARTMENU
    'Returns &h16
    Public Property Get ssfCOMMONSTARTMENU : ssfCOMMONSTARTMENU = &h16 : End Property

    'Property ssfCOMMONPROGRAMS
    'Returns &h17
    Public Property Get ssfCOMMONPROGRAMS : ssfCOMMONPROGRAMS = &h17 : End Property

    'Property ssfCOMMONSTARTUP
    'Returns &h18
    Public Property Get ssfCOMMONSTARTUP : ssfCOMMONSTARTUP = &h18 : End Property

    'Property ssfCOMMONDESKTOPDIR
    'Returns &h19
    Public Property Get ssfCOMMONDESKTOPDIR : ssfCOMMONDESKTOPDIR = &h19 : End Property

    'Property ssfAPPDATA
    'Returns &h1a
    Public Property Get ssfAPPDATA : ssfAPPDATA = &h1a : End Property

    'Property ssfPRINTHOOD
    'Returns &h1b
    Public Property Get ssfPRINTHOOD : ssfPRINTHOOD = &h1b : End Property

    'Property ssfLOCALAPPDATA
    'Returns &h1c
    Public Property Get ssfLOCALAPPDATA : ssfLOCALAPPDATA = &h1c : End Property

    'Property ssfALTSTARTUP
    'Returns &h1d
    Public Property Get ssfALTSTARTUP : ssfALTSTARTUP = &h1d : End Property

    'Property ssfCOMMONALTSTARTUP
    'Returns &h1e
    Public Property Get ssfCOMMONALTSTARTUP : ssfCOMMONALTSTARTUP = &h1e : End Property

    'Property ssfCOMMONFAVORITES
    'Returns &h1f
    Public Property Get ssfCOMMONFAVORITES : ssfCOMMONFAVORITES = &h1f : End Property

    'Property ssfINTERNETCACHE
    'Returns &h20
    Public Property Get ssfINTERNETCACHE : ssfINTERNETCACHE = &h20 : End Property

    'Property ssfCOOKIES
    'Returns &h21
    Public Property Get ssfCOOKIES : ssfCOOKIES = &h21 : End Property

    'Property ssfHISTORY
    'Returns &h22
    'Remarks: File system directory that serves as a common repository for Internet history items.
    Public Property Get ssfHISTORY : ssfHISTORY = &h22 : End Property

    'Property ssfCOMMONAPPDATA
    'Returns &h23
    Public Property Get ssfCOMMONAPPDATA : ssfCOMMONAPPDATA = &h23 : End Property

    'Property ssfWINDOWS
    'Returns &h24
    Public Property Get ssfWINDOWS : ssfWINDOWS = &h24 : End Property

    'Property ssfSYSTEM
    'Returns &h25
    Public Property Get ssfSYSTEM : ssfSYSTEM = &h25 : End Property

    'Property ssfPROGRAMFILES
    'Returns &h26
    Public Property Get ssfPROGRAMFILES : ssfPROGRAMFILES = &h26 : End Property

    'Property ssfMYPICTURES
    'Returns &h27
    Public Property Get ssfMYPICTURES : ssfMYPICTURES = &h27 : End Property

    'Property ssfPROFILE
    'Returns &h28
    Public Property Get ssfPROFILE : ssfPROFILE = &h28 : End Property

    'Property ssfSYSTEMx86
    'Returns &h29
    Public Property Get ssfSYSTEMx86 : ssfSYSTEMx86 = &h29 : End Property

    'Property ssfPROGRAMFILESx86
    'Returns &h30
    Public Property Get ssfPROGRAMFILESx86 : ssfPROGRAMFILESx86 = &h30 : End Property

End Class