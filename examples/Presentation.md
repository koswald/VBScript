# Presentation.vbs

[Overview]  
[Requirements]  
[References]  

## Overview

[Presentation.vbs] provides a way to keep the computer awake while giving a presentation, for systems that don't have `PresentationSettings.exe`, notably Windows 10 Home systems.  

Click the moon icon in the system tray to see a menu.  
  
When presentation mode is on, the computer and monitor are typically prevented from going into a suspend (sleep) state or hibernation. The computer may still be put to sleep by other applications or by user actions such as closing a laptop lid or pressing a sleep button or power button.  

Phone charger mode is the same as presentation mode except that the monitor is turned off, initially

## Requirements

You must have run [Setup.vbs], which installs the VBScript extensions used by [Presentation.vbs].

[Presentation.vbs] may require changes to Windows settings in order to work properly.

Specifically, you may want to increase the value that controls when the following components go to sleep or turn off for both AC (plugged in) and DC (battery):  

- Display(s)  
- Hard Drive(s)  
- Computer  

[PresentationSettings.hta] is intended to make these settings accessible from a single location.

## References  

[Watcher.cs documentation](../docs/CSharpClasses.md#watcher "github.com\koswald\VBScript")  
[How to change Lock screen timeout before display turn off on Windows 10]  
[Does Your Windows Computer Display Turn Off Every 15 Minutes?]  
[Canonical Names of Control Panel Items]  
[SetThreadExecutionState]  

[Overview]: #overview  
[Requirements]: #requirements  
[References]: #references  

[Setup.vbs]: ../Setup.vbs
[Presentation.vbs]: Presentation.vbs
[PresentationSettings.hta]: PresentationSettings.hta

[How to change Lock screen timeout before display turn off on Windows 10]: https://www.windowscentral.com/how-extend-lock-screen-timeout-display-turn-windows-10 "windowscentral.com"
[Does Your Windows Computer Display Turn Off Every 15 Minutes?]: https://www.online-tech-tips.com/windows-7/does-your-windows-7-computer-display-turn-off-every-15-minutes/ "online-tech-tips.com"
[Canonical Names of Control Panel Items]: https://docs.microsoft.com/en-us/windows/desktop/shell/controlpanel-canonical-names "docs.microsoft.com"
[SetThreadExecutionState]: https://docs.microsoft.com/en-us/windows/desktop/api/winbase/nf-winbase-setthreadexecutionstate "docs.microsoft.com"
