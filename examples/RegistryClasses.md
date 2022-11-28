# RegistryClasses.hta

[Overview]  
[Examples]  
[Tips]  
[References]

## Overview

With RegistryClasses.hta you can manage certain 
aspects of Windows classes without directly using 
regedit.exe, such as

- Creating a new class and file type.
- Creating a new verb for an existing class.

## Examples

Create a new file type, .nft

- In the *File type* field, enter `nft`.
- Click the *New verb* button.
- Enter `Open .nft with Notepad` in the *Verb name* field and click the *Save* button.
- In the *Command* field, enter  
    `notepad "%1"`.
- Create a new .nft file, right click it, and you should see a menu item for opening the file in Notepad.
  
Open .hta files with WordPad.

- In the ProgId field, enter `htafile` or enter `hta` in the *File type* field.
- Click the *New verb* button.
- Enter `Open .hta with WordPad` in the *Verb name* field and click the *Save* button.
- In the *Command* field, enter  
    `"C:\Program Files\Windows NT\Accessories\wordpad.exe" "%1"`.
- Right click an .hta file, and you should see a menu item to open the file in WordPad.

Open PowerShell at any folder

- Select HKEY_CURRENT_USER (recommended). Or click Elevate to elevate privileges and then select HKEY_LOCAL_MACHINE.
- In the ProgId field, enter `Directory`.
- Click the *New verb* button.
- Enter `Open PowerShell here` and click the *Ok* button.
- In the *Command* field, enter `powershell -NoExit -NoLogo Set-Location '%1'`. Use `pwsh` instead of `powershell` for PowerShell Core, and see note below. Note the single-quotes around the folder name. Double quotes will not work properly for folders with spaces.  
- Open explorer.exe, right-click a folder, and click *Open PowerShell here*.
- A PowerShell should open at the selected folder.

> *Note:*  There is a known issue with PowerShell for folders with single-quotes in  
> the name. It is recommended not to use single-quotes in folder or file names.  

> *Note:* In PowerShell 6+ (PowerShell Core), a `Set-Location` item in `*profile.ps1` may  
> override this feature. Workaround: Modify the taskbar shortcut's 'Start in' field instead.  

## Tips

- Click F1 in `RegistryClasses.hta` to open this file.

- If you might want to delete a verb in the future, without using regedit.exe, then don't use cannonical verb names such as open, opennew, print, explore, find, openas, properties, printto, runas, and runasuser.

- When creating a new file type, it may take a minute or two before the new file type appears in the shortcut (New item) menu. And it might take a logoff/login in order for the default icon to appear in explorer and in the shortcut menu.

- Using HKEY_CLASSES_USER gives similar functionality without having to elevate privileges, and it doesn't affect other users. HKEY_LOCAL_MACHINE can be viewed without elevated privileges, but not edited.

## References

[Extending Shortcut Menus](https://learn.microsoft.com/en-us/windows/desktop/shell/context "learn.microsoft.com")  
[Use JavaScript to place cursor at end of text in text input element
](https://stackoverflow.com/questions/511088/use-javascript-to-place-cursor-at-end-of-text-in-text-input-element#26900921 "stackoverflow.com")

[Overview]: #overview
[Examples]: #examples
[Tips]: #tips
[References]: #references
