
'Constants for use with WScript.Shell.Run

Class ShellConstants

    'Property RunHidden
    'Returns 0
    'Remark: Window opens hidden. <br /> For use with Run method parameter #2
    Property Get RunHidden : RunHidden = 0 : End Property

    'Property RunNormal
    'Returns 1
    'Remark: Window opens normal. <br /> For use with Run method parameter #2
    Property Get RunNormal : RunNormal = 1 : End Property

    'Property RunMinimized
    'Returns 2
    'Remark: Window opens minimized. <br /> For use with Run method parameter #2
    Property Get RunMinimized : RunMinimized = 2 : End Property

    'Property RunMaximized
    'Returns 3
    'Remark: Window opens maximized. <br /> For use with Run method parameter #2
    Property Get RunMaximized : RunMaximized = 3 : End Property

    'Property Synchronous
    'Returns True
    'Remark: Script execution halts and waits for the called process to exit. <br /> For use with Run method parameter #3
    Property Get Synchronous : Synchronous = True : End Property

    'Property Asynchronous
    'Returns False
    'Remark: Script execution proceeds without waiting for the called process to exit. <br /> For use with Run method parameter #3
    Property Get Asynchronous : Asynchronous = False : End Property

End Class
