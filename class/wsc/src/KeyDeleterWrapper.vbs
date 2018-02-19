Dim kd : Set kd = New KeyDeleter

Sub DeleteKey(root, key) : kd.DeleteKey root, key : End Sub
Sub ValidateRoot(root) : kd.ValidateRoot root : End Sub
Sub ValidateKey(key) : kd.ValidateKey key : End Sub
Sub ValidateSubkey(subkey) : kd.ValidateSubkey subkey : End Sub
Sub ValidateBackslashCount(key, subkey) : kd.ValidateBackslashCount key, subkey : End Sub
Function BackslashCount : BackslashCount = kd.BackslashCount : End Function
Function MaxCount : MaxCount = kd.MaxCount : End Function
Function SavedKey : SavedKey = kd.SavedKey : End Function
Function SavedSubkey : SavedSubkey = kd.SavedSubkey : End Function
Function Result : Result = kd.Result : End Function
Function getDelete : getDelete = kd.Delete : End Function
Sub putDelete(newValue) : kd.Delete = newValue : End Sub
Function HKCR : HKCR = kd.HKCR : End Function
Function HKCU : HKCU = kd.HKCU : End Function
Function HKLM : HKLM = kd.HKLM : End Function
Function HKU : HKU = kd.HKU : End Function
Function HKCC : HKCC = kd.HKCC : End Function