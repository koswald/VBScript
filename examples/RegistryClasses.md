# RegistryClasses.hta

[Overview](#overview)  
[Examples](#examples)  
[Tips](#tips)  
[References](#references)

## Overview

With RegistryClasses.hta you can manage certain 
aspects of Windows classes without directly using 
regedit.exe, such as

- Creating a new file type.
- Adding a new verb to an existing file type.
- Adding a new verb to an existing class.

> WARNING: Backup the registry before making changes!

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

Open command prompt at any folder

- Select HKEY_CURRENT_USER (recommended). Or click Elevate to elevate privileges and then select HKEY_LOCAL_MACHINE.
- In the ProgId field, enter `Folder`.
- Click the *New verb* button.
- Enter `Open command prompt here` and click the *Save* button.
- In the *Command* field, enter `cmd /k cd "%1"`.
- Open explorer.exe, right-click a folder, and click *Open command prompt here*.
- A command prompt should open at the selected folder.

## Tips

- Click F1 to open this file.

- If you might want to delete a verb in the future, without using regedit.exe, then don't use cannonical verb names such as open, opennew, print, explore, find, openas, properties, printto, runas, and runasuser.

- When creating a new file type, it may take a minute or two before the new file type appears in the shortcut (New item) menu. And it might take a logoff/login in order for the default icon to appear in explorer and in the shortcut menu.

- Using HKEY_CLASSES_USER gives similar functionality without having to elevate privileges, and it doesn't affect other users. HKEY_LOCAL_MACHINE can be viewed without elevated privileges, but not edited.

## References

[Extending Shortcut Menus](https://msdn.microsoft.com/en-us/library/windows/desktop/cc144101(v=vs.85).aspx "msdn.microsoft.com")  
[Use JavaScript to place cursor at end of text in text input element
](https://stackoverflow.com/questions/511088/use-javascript-to-place-cursor-at-end-of-text-in-text-input-element#26900921 "stackoverflow.com")