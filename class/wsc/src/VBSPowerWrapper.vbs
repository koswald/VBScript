'wrap the VBSPower class for the VBSPower.wsc component
Dim vp : Set vp = New VBSPower
Sub Shutdown : vp.Shutdown : End Sub
Sub Restart : vp.Restart : End Sub
Sub Logoff : vp.Logoff : End Sub
Sub Sleep : vp.Sleep : End Sub
Sub Hibernate : vp.Hibernate : End Sub
Sub EnableHibernation : vp.EnableHibernation : End Sub
Sub DisableHibernation : vp.DisableHibernation : End Sub
Sub SetForce(newValue) : vp.SetForce(newValue) : End Sub
Sub SetDebug(newValue) : vp.SetDebug(newValue) : End Sub
