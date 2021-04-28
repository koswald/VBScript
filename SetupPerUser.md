# Setup per user vs per machine

[Overview]  
[Hypothetical advantages of a per-user setup]  
[Disadvantages]  
[References]  

## Overview

It is recommended that [Setup.vbs] be used to setup the project for all users.  

[SetupExperiment.wsf] was used *on an experimental basis* to register objects for the current user only.  

With more work it may be possible in limited circumstances to use a per-user registration of the COM objects in this project, but it is not considered desirable for a number of reasons, and has not been successfully tested.  

## Hypothetical advantages of per-user registration

Per-user registration of COM objects *may appear* to be desirable for at least a couple of use cases.

- You are making changes to the project, but you want the project to be stable for another user, or  

- You want to enable users to install the project without admin privileges.

## What are the disadvantages of a per-user setup

- COM objects registered in HKEY_CURRENT_USER are not available to processes running with elevated privileges.  
This has the potential to greatly complicate registration of the objects. It might be necessary to register some project assemblies or components in HKLM and some in HKCU, depending on which ones would potentially be used by elevated processes. It might be desirable to register some or all objects in both HKLM and HKCU--but this approach was not tested.

- Windows Script Components (.wsc) events do not or may not work properly, ComponentEventExample.wsc for example. To reproduce, run Uninstall.vbs, run SetupExperiment.wsf, then run examples\ComponentEventExample.hta.

## References

[The Dangers of Per-User COM Objects]. It is by design that per-user COM configuration has precedence over machine-wide configuration when privileges are not elevated. And it is by design that elevated processes ignore per-user configuration.  

[Per-User COM Registrations and Elevated Processes with UAC on Windows Vista SP1]. This is written for Vista, but still applies to Windows 10, seemingly. It confirms that the behaviour described above is intentional and gives some rationale for the design decision.  

[search: register dll for current user]  

`` ``

[Overview]: #overview
[Hypothetical advantages of a per-user setup]: #hypothetical-advantages-of-per-user-registration
[Disadvantages]: #what-are-the-disadvantages-of-a-per-user-setup
[References]: #references
[SetupExperiment.wsf]: SetupExperiment.wsf
[Setup.vbs]: Setup.vbs

[search: register dll for current user]: https://www.google.com/search?&q=register+dll+for+current+user  

[Register COM object for Current User]: https://stackoverflow.com/questions/35782404/registering-a-com-without-admin-rights  

[Windows registry information for advanced users]: https://support.microsoft.com/en-us/help/256986 "https://support.microsoft.com"

[The Dangers of Per-User COM Objects]: https://www.virusbulletin.com/uploads/pdf/conference_slides/2011/Larimer-VB2011.pdf  

[Per-User COM Registrations and Elevated Processes with UAC on Windows Vista SP1]: https://techcommunity.microsoft.com/t5/windows-blog-archive/per-user-com-registrations-and-elevated-processes-with-uac-on/ba-p/228531
