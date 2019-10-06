# C# Classes Documentation

### Contents

[Admin](#admin)  
[ComEvent](#comevent)  
[EventLogger](#eventlogger)  
[FileChooser](#filechooser)  
[FolderChooser](#folderchooser)  
[FolderChooser2](#folderchooser2)  
[IconExtractor](#iconextractor)  
[NotifyIcon](#notifyicon)  
[ProgressBar](#progressbar)  
[SpeechSynthesis](#speechsynthesis)  
[Timer](#timer)  
[Watcher](#watcher)  


## Admin

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| Admin | Provide miscellaneous system admin. features.  |  |  | Type | | VBScripting |
| IAdmin | COM interface for VBScripting.Admin  |  |  | Type | | VBScripting |
| IsAdministrator | Gets whether the current user is in the Administrator group (on the current machine). Slow. May take five seconds or longer. Does not necessarily mean that privileges are elevated. Adapted from a <a href="https://stackoverflow.com/questions/44507149/how-to-check-if-current-user-is-in-admin-group-c-sharp#answer-47564106" title="stackoverflow.com" target="_blank"> stackoverflow.com post</a>.  |  |  | Method | Admin | VBScripting |
| Log | Logs the specified message to the event log (source="VBScripting").  |  | message | Method | Admin | VBScripting |
| GetLogs | Get an array of logs entries from the Application log. Returns an array of logs (strings) from the specified event source that contain the specified message string. Searches the Application log only. | an array | source, message | Method | Admin | VBScripting |
| SourceExists | Gets whether the specified EventLog source exists.  | a boolean | source | Method | Admin | VBScripting |
| CreateEventSource | Creates the specified EventLog source.  | an EventLogSourceResult | source | Method | Admin | VBScripting |
| DeleteEventSource | Deletes the specified EventLog source and all of its logs.  | an EventLogSourceResult | source | Method | Admin | VBScripting |
| MonitorOff | Turn off the monitor(s)  |  |  | Method | Admin | VBScripting |
| PrivilegesAreElevated | Gets whether the current process has elevated privileges.  |  |  | Property | Admin | VBScripting |
| EventSource | Gets the name of the event log source for this namespace (VBScripting).  | a string |  | Property | Admin | VBScripting |
| LogName | Gets the name of the log to which events will be logged.  | a string |  | Property | Admin | VBScripting |
| Result | Gets an EventLogResultT object. VBScript example: <pre> Set returnValue = adm.CreateEventSource <br/> If returnValue.Result = adm.Result.SourceCreationException Then <br/>     MsgBox returnValue.Message <br/> End If</pre> | an EventLogResultT |  | Property | Admin | VBScripting |
| EventLogResultT | Provides a set of terse behavior/result descriptions suitable for VBScript comparisons and MsgBox captions. Not directly available to VBScript. See <tt>Admin.Result</tt>. |  |  | Type | | VBScripting |
| SourceAlreadyExists |   | "Source already exists" |  | Property | EventLogResultT | VBScripting |
| SourceCreated |   | "Source created" |  | Property | EventLogResultT | VBScripting |
| SourceCreationException |   | "Source creation error" |  | Property | EventLogResultT | VBScripting |
| SourceDoesNotExist |   | "Source does not exist" |  | Property | EventLogResultT | VBScripting |
| SourceDeleted |   | "Source deleted" |  | Property | EventLogResultT | VBScripting |
| SourceDeletionException |   | "Source deletion error" |  | Property | EventLogResultT | VBScripting |
| EventLogSourceResult | Type returned by CreateEventSource and DeleteEventSource.  |  |  | Type | | VBScripting |
| SourceExists | Returns True if the source exists after the attempted operation has completed.  | a boolean |  | Property | EventLogSourceResult | VBScripting |
| Message | Returns a message descriptive of the outcome of the operation.  | a string |  | Property | EventLogSourceResult | VBScripting |
| Result | Returns a string: one of the EventLogResultT strings.  | a string |  | Property | EventLogSourceResult | VBScripting |

## ComEvent

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| ComEvent | Invokes VBS members from C#. <span class="red"> This class is not accessible from VBScript. </span>  |  |  | Type | | VBScripting |
| InvokeComCallback | Invokes a VBScript method. The parameter <tt>callbackRef</tt> is a reference to a VBScript member returned by the VBScript Function GetRef. |  | callbackRef | Method | ComEvent | VBScripting |

## EventLogger

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| IEventLogger | A COM Interface for VBScripting.EventLogger  |  |  | Type | | VBScripting |
| EventLogger | Provides system logging for VBScript.  |  |  | Type | | VBScripting |
| log | Writes the specified message to the Application event log.  |  | message | Method | EventLogger | VBScripting |

## FileChooser

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| FileChooser | Provides a file chooser dialog for VBScript.  |  |  | Type | | VBScripting |
| IFileChooser | The COM interface for VBScripting.FileChooser  |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  |  | Method | FileChooser | VBScripting |
| FileName | Opens a dialog enabling the user to browse for and choose a file. Returns the filespec of the chosen file. Returns an empty string if the user cancels. |  |  | Property | FileChooser | VBScripting |
| FileNames | Opens a dialog enabling the user to browse for and choose multiple files. Gets a string array of filespecs. Returns an empty array if the user cancels. Requires Multiselect to have been set to True. |  |  | Property | FileChooser | VBScripting |
| FileNamesString | Opens a dialog enabling the user to browse for and choose multiple files. Gets a string of filespecs delimited by a vertical bar (&#124;). Returns an empty string if the user cancels. Requires Multiselect to have been set to True. |  |  | Property | FileChooser | VBScripting |
| InitialDirectory | Gets or sets directory at which the dialog opens.  |  |  | Property | FileChooser | VBScripting |
| ERInitialDirectory | Gets the initial directory with relative path resolved and environment variables expanded. Improves testability. |  |  | Property | FileChooser | VBScripting |
| Filter | Gets or sets the selectable file types. Examples: <pre> fc.Filter = "All files (&#42;.&#42;)&#124;&#42;.&#42;" // the default <br/> fc.Filter = "Text files (&#42;.txt)&#124;&#42;.txt&#124;All files (&#42;.&#42;)&#124;&#42;.&#42;" <br/> fc.Filter = "Image Files(&#42;.BMP;&#42;.JPG;&#42;.GIF)&#124;&#42;.BMP;&#42;.JPG;&#42;.GIF&#124;All files (&#42;.&#42;)&#124;&#42;.&#42;" </pre> |  |  | Property | FileChooser | VBScripting |
| FilterIndex | Gets or sets the index controlling which filter item is initially selected. An integer. The index is 1-based. The default is 1. |  |  | Property | FileChooser | VBScripting |
| Title | Gets or sets the dialog titlebar text. The default text is "Browse for a file." |  |  | Property | FileChooser | VBScripting |
| Multiselect | Gets or sets whether multiple files can be selected. The default is False. |  |  | Property | FileChooser | VBScripting |
| DereferenceLinks | Indicates whether the returned file is the referenced file or the .lnk file itself. Gets or sets, if the selected file is a .lnk file, whether the filespec returned refers to the .lnk file itself (False) or to the file that the .lnk file points to (True). The default is False. |  |  | Property | FileChooser | VBScripting |
| DefaultExt | Gets or sets the file extension name that is automatically supplied when one is not specified. A string. The default is "txt". |  |  | Property | FileChooser | VBScripting |
| ValidateNames | Gets or sets whether to validate the file name(s).  |  |  | Property | FileChooser | VBScripting |
| CheckFileExists | Gets or sets whether to check that the file exists.  |  |  | Property | FileChooser | VBScripting |

## FolderChooser

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| IFolderChooser | COM interface for VBScripting.FolderChooser  |  |  | Type | | VBScripting |
| FolderChooser | Present the Windows Vista-style open file dialog to select a folder. Fall back for older Windows Versions. Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/57611/erike"> EricE</a>. Uses <tt> System.Reflection</tt>. |  |  | Type | | VBScripting |
| InitialDirectory | Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory.  |  |  | Property | FolderChooser | VBScripting |
| Title | Gets or sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder".  |  |  | Property | FolderChooser | VBScripting |
| FolderName | Opens a dialog and returns the folder selected by the user.  | a path |  | Property | FolderChooser | VBScripting |

## FolderChooser2

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| IFolderChooser2 | COM interface for VBScripting.FolderChooser2  |  |  | Type | | VBScripting |
| FolderChooser2 | Present the Windows Vista-style open file dialog to select a folder. Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/15368771/show-detailed-folder-browser-from-a-propertygrid#15386992"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/403671/simon-mourier"> Simon Mourier</a>. Uses <tt> System.Runtime.InteropServices</tt>. |  |  | Type | | VBScripting |
| InitialDirectory | Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory. |  |  | Property | FolderChooser2 | VBScripting |
| Title | Sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder".  |  |  | Property | FolderChooser2 | VBScripting |
| FolderName | Opens a dialog and returns the folder selected by the user.  | a path |  | Property | FolderChooser2 | VBScripting |

## IconExtractor

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| IconExtractor | Extracts an icon from a .dll or .exe file. <span class="red"> Not all members of this class are accessible to VBScript. </span> |  |  | Type | | VBScripting |
| IIconExtractor | A COM Interface for VBScripting.IconExtractor  |  |  | Type | | VBScripting |
| (Constructor) | Constructor.  |  |  | Method | IconExtractor | VBScripting |
| Save | Extracts an icon from a .dll or .exe and saves it to a file. Parameters: resFile is the .dll or .exe file; index selects the icon within the resource file; icoFile is the output file; largeIcon is a boolean: True if a large icon is to be extracted, False for a small icon. Environment variables and relative paths are allowed. |  | resFile, index, icoFile, largeIcon | Method | IconExtractor | VBScripting |
| SetImageFormatBmp | Change the image format to BMP. Default is BMP.  |  |  | Method | IconExtractor | VBScripting |
| SetImageFormatPng | Change the image format to PNG. Default is BMP.  |  |  | Method | IconExtractor | VBScripting |
| Extract | Extracts an icon from the specified .dll or .exe file. <span class="red"> This method is static and so it is not directly available to VBScript. </span> Other parameters: <tt>index</tt> is an integer that specifies the icon's index within the resource. <tt>largeIcon</tt> is a boolean that specifies whether the icon should be a large icon; if False, a small icon is extracted, if available. The icon must be disposed in order to free memory. | an icon | file, index, largeIcon | Method | IconExtractor | VBScripting |
| IconCount | Returns the number of icons in a .dll or .exe file. A relative path or environmental variable is allowed. | an int | filespec (.dll or .exe) | Method | IconExtractor | VBScripting |
| GetPointer | Gets a pointer to an icon. Must be disposed with DisposeIcon(pointer) or Icon.Dispose(), in order to release memory. A relative path or environmental variable is allowed. | integer | file, index, largeIcon | Method | IconExtractor | VBScripting |
| ExtractIcon | Gets an icon. Must be disposed with DisposeIcon(pointer) or Icon.Dispose(). | Icon | integer | Method | IconExtractor | VBScripting |
| DisposeIcon | Dispose an icon by pointer (an int).  | Returns true for success. | pointer | Method | IconExtractor | VBScripting |

## NotifyIcon

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| NotifyIcon | Provides a system tray icon for VBScript, for illustration purposes.  |  |  | Type | | VBScripting |
| INotifyIcon | The COM interface for VBScripting.NotifyIcon  |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  |  | Method | NotifyIcon | VBScripting |
| SetIconByIcoFile | Sets the system tray icon given an .ico file. The parameter <tt>fileName</tt> specifies the filespec of the .ico file. Environment variables and relative paths are allowed. |  | fileName | Method | NotifyIcon | VBScripting |
| SetIconByDllFile | Sets the system tray icon from a .dll or .exe file. Parameters: <tt>fileName</tt> is the path and name of a .dll or .exe file that contains icons. <tt>index</tt> is an integer that specifies which icon to use. <tt>largeIcon</tt> is a boolean that specifies whether to use a large or small icon. |  | fileName, index, largeIcon | Method | NotifyIcon | VBScripting |
| SetBalloonTipIcon | Sets the icon of the "balloon tip" or notification. The parameter <tt>type</tt> is an integer that specifies which icon to use: Return values of ToolTipIcon properties can be used: Error = 1, Info = 2, None = 3, Warning = 4. |  | type | Method | NotifyIcon | VBScripting |
| Dispose | Disposes of the icon resources when it is no longer needed. If this method is not called, the icon may persist in the system tray until the mouse hovers over it, even after the object instance has lost scope. |  |  | Method | NotifyIcon | VBScripting |
| ShowBalloonTip | Show the balloon tip.  |  |  | Method | NotifyIcon | VBScripting |
| AddMenuItem | Add a menu item to the system tray icon's context menu. This method can be called only from VBScript. The parameter <tt>menuText</tt> is a string that specifies the text that appears in the menu. The parameter <tt>callbackRef</tt> is a VBScript object reference returned by the VBScript GetRef Function. |  | menuText, callbackRef | Method | NotifyIcon | VBScripting |
| InvokeCallbackByIndex | Provide callback testability from VBScript.  |  |  | Method | NotifyIcon | VBScripting |
| DisableMenuItem | Disable a menu item  |  |  | Method | NotifyIcon | VBScripting |
| EnableMenuItem | Enable a menu item  |  |  | Method | NotifyIcon | VBScripting |
| ShowContextMenu | Show the context menu. Public in order to provide testability from VBScript. |  |  | Method | NotifyIcon | VBScripting |
| SetBalloonTipCallback | Sets the VBScript callback Sub or Function reference invoked when the notification "balloon" is clicked. VBScript example: <pre>    obj.SetBalloonTipCallback GetRef("ProcedureName") </pre> |  |  | Method | NotifyIcon | VBScripting |
| Text | Gets or sets the text shown when the mouse hovers over the system tray icon.  |  |  | Property | NotifyIcon | VBScripting |
| Visible | Gets or sets the icon's visibility. A boolean. Required. Set this property to True after initializing other settings. |  |  | Property | NotifyIcon | VBScripting |
| BalloonTipTitle | Gets or sets the title of the "balloon tip" or notification.  |  |  | Property | NotifyIcon | VBScripting |
| BalloonTipText | Gets or sets the text of the "balloon tip" or notification.  |  |  | Property | NotifyIcon | VBScripting |
| BalloonTipLifetime | Gets or sets the lifetime of the "balloon tip" or notification. An integer (milliseconds). Deprecated as of Windows Vista, the value is overridden by accessibility settings.  |  |  | Property | NotifyIcon | VBScripting |
| ToolTipIcon | Gets an object useful in VBScript for selecting a ToolTipIcon type. The properties Error, Info, None, and Warning may be used with SetBalloonTipIcon. VBScript example: <pre>    obj.SetBallonTipIcon obj.ToolTipIcon.Warning </pre> | a ToolTipIconT |  | Property | NotifyIcon | VBScripting |
| ToolTipIconT | Supplies the type required by NotifyIcon.ToolTipIcon This class is not directly accessible from VBScript , however, it is accessible via the <tt>NotifyIcon.ToolTipIcon</tt> property. |  |  | Type | | VBScripting |
| Error |   | 1 |  | Property | ToolTipIconT | VBScripting |
| Info |   | 2 |  | Property | ToolTipIconT | VBScripting |
| None |   | 3 |  | Property | ToolTipIconT | VBScripting |
| Warning |   | 4 |  | Property | ToolTipIconT | VBScripting |
| CallbackEventSettings | Settings for saving VBScript method references. This class is not accessible from VBScript.  |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  |  | Method | CallbackEventSettings | VBScripting |
| AddRef | Adds a CallbackReference instance reference to the list.  |  | callbackRef | Method | CallbackEventSettings | VBScripting |
| Refs | Gets or sets a list of callback references.  |  |  | Property | CallbackEventSettings | VBScripting |
| CallbackReference | An orderly way to save the index and callback reference for a single menu item. This class is not accessible to VBScript. |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  | index, reference | Method | CallbackReference | VBScripting |
| Index | This Index corresponds to the Index of a menuItem in the context menu.  |  |  | Property | CallbackReference | VBScripting |
| Reference | COM object generated by VBScript's GetRef Function.  |  |  | Property | CallbackReference | VBScripting |

## ProgressBar

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| ProgressBar | Supplies a progress bar to VBScript, for illustration purposes.  |  |  | Type | | VBScripting |
| IProgressBar | Exposes the VBScripting.ProgressBar members to COM/VBScript.  |  |  | Type | | VBScripting |
| PerformStep | Advances the progress bar one step.  |  |  | Method | ProgressBar | VBScripting |
| FormSize | Sets the size of the window.  |  | width, height | Method | ProgressBar | VBScripting |
| PBarSize | Sets the size of the progress bar.  |  | width, height | Method | ProgressBar | VBScripting |
| FormLocation | Sets the position of the window.  |  | x, y | Method | ProgressBar | VBScripting |
| FormLocationByPercentage | Sets the position of the window.  |  | x, y | Method | ProgressBar | VBScripting |
| PBarLocation | Sets the location of the progress bar within the window.  |  | x, y | Method | ProgressBar | VBScripting |
| SuspendLayout | Suspends drawing of the window temporarily.  |  |  | Method | ProgressBar | VBScripting |
| ResumeLayout | Resumes drawing the window.  |  |  | Method | ProgressBar | VBScripting |
| SetIconByIcoFile | Sets the icon given the filespec of an .ico file. Environment variables are allowed. |  | fileName | Method | ProgressBar | VBScripting |
| SetIconByDllFile | Sets the icon given the filespec of a .dll or .exe file and an index. The index is an integer that identifies the icon. Environment variables are allowed. |  | fileName, index | Method | ProgressBar | VBScripting |
| Dispose | Disposes of the object's resources.  |  |  | Method | ProgressBar | VBScripting |
| Visible | Gets or sets the progress bar's visibility. A boolean. The default is False. |  |  | Property | ProgressBar | VBScripting |
| Minimum | Gets or sets the value at which there is no apparent progress. An integer. The default is 0. |  |  | Property | ProgressBar | VBScripting |
| Maximum | Gets or sets the value at which the progress appears to be complete. An integer. The default is 100. |  |  | Property | ProgressBar | VBScripting |
| Value | Gets or sets the apparent progress. An integer. Should be at or above the minimum and at or below the maximum. |  |  | Property | ProgressBar | VBScripting |
| Step | Gets or sets the increment between steps.  |  |  | Property | ProgressBar | VBScripting |
| Caption | Gets or sets the window title-bar text.  |  |  | Property | ProgressBar | VBScripting |
| Debug | Gets or sets whether the type is under development. Affects the behavior of two methods, SetIconByIcoFile and SetIconByDllFile, if exceptions are thrown: when debugging, a message box is shown. Default is False. |  |  | Property | ProgressBar | VBScripting |
| BorderStyle | Provides an object useful in VBScript for setting FormBorderStyle.  | a FormBorderStyleT |  | Property | ProgressBar | VBScripting |
| FormBorderStyle | Sets the style of the window border. An integer. One of the BorderStyle property return values can be used: Fixed3D, FixedDialog, FixedSingle, FixedToolWindow, None, Sizable (default), or SizableToolWindow. VBScript example: <pre> pb.FormBorderStyle = pb.BorderStyle.Fixed3D </pre> |  |  | Property | ProgressBar | VBScripting |
| FormBorderStyleT | Enumeration of border styles. This class is available to VBScript via the <tt>ProgressBar.BorderStyle</tt> property. |  |  | Type | | VBScripting |
| Fixed3D |   | 1 |  | Property | FormBorderStyleT | VBScripting |
| FixedDialog |   | 2 |  | Property | FormBorderStyleT | VBScripting |
| FixedSingle |   | 3 |  | Property | FormBorderStyleT | VBScripting |
| FixedToolWindow |   | 4 |  | Property | FormBorderStyleT | VBScripting |
| None |   | 5 |  | Property | FormBorderStyleT | VBScripting |
| Sizable |   | 6 |  | Property | FormBorderStyleT | VBScripting |
| SizableToolWindow |   | 7 |  | Property | FormBorderStyleT | VBScripting |

## SpeechSynthesis

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| SpeechSynthesis | Provide a wrapper for the .Net speech synthesizer for VBScript, for demonstration purposes. Requires an assembly reference to <tt>%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\WPF\System.Speech.dll</tt>. |  |  | Type | | VBScripting |
| ISpeechSynthesis | The COM interface for <tt>VBScripting.SpeechSynthesis</tt>
  |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  |  | Method | SpeechSynthesis | VBScripting |
| Speak | Convert text to speech. This method is synchronous. |  |  | Method | SpeechSynthesis | VBScripting |
| SpeakAsync | Convert text to speech. This method is asynchronous. |  |  | Method | SpeechSynthesis | VBScripting |
| Pause | Pause speech synthesis.  |  |  | Method | SpeechSynthesis | VBScripting |
| Resume | Resume speech synthesis.  |  |  | Method | SpeechSynthesis | VBScripting |
| Voices | Gets an array of the names of the installed, enabled voices. Each element of the array can be used to set <tt>Voice</tt>. |  |  | Method | SpeechSynthesis | VBScripting |
| Dispose | Disposes the SpeechSynthesis object's resources.  |  |  | Method | SpeechSynthesis | VBScripting |
| Voice | Gets or sets the current voice by name. A string. One of the names from the <tt>Voices</tt> array. |  |  | Property | SpeechSynthesis | VBScripting |
| SynthesizerState | Gets the state of the SpeechSynthesizer. Read only. Returns an integer equal to one of the <tt>State</tt> method return values. |  |  | Property | SpeechSynthesis | VBScripting |
| Volume | Gets or sets the volume. An integer from 0 to 100. |  |  | Property | SpeechSynthesis | VBScripting |
| State | Gets an object whose properties (Ready, Paused, and Speaking) provide values useful for comparing to <tt>SynthesizerState</tt>.  | a SynthersizerStateT |  | Property | SpeechSynthesis | VBScripting |
| SynthesizerStateT | Enumerates the synthesizer states. Not intended for use in VBScript. See <tt>SpeechSynthesis.State</tt>. |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  |  | Method | SynthesizerStateT | VBScripting |
| Ready |   | 1 |  | Property | SynthesizerStateT | VBScripting |
| Paused |   | 2 |  | Property | SynthesizerStateT | VBScripting |
| Speaking |   | 3 |  | Property | SynthesizerStateT | VBScripting |
| Unexpected |   | 4 |  | Property | SynthesizerStateT | VBScripting |

## Timer

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| Timer | Wraps the <a href="https://docs.microsoft.com/en-us/dotnet/api/system.timers.timer?view=netframework-4.7.1" title="docs.microsoft.com"> System.Timers.Timer class</a> for VBScript.  |  |  | Type | | VBScripting |
| ITimer | COM interface for VBScripting.Timer  |  |  | Type | | VBScripting |
| (Constructor) | Constructor  |  |  | Method | Timer | VBScripting |
| Start | Starts or restarts the timer.  |  |  | Method | Timer | VBScripting |
| Stop | Stops the timer.  |  |  | Method | Timer | VBScripting |
| Dispose | Disposes of the timer's resources.  |  |  | Method | Timer | VBScripting |
| Interval | Gets or sets the number of milliseconds between when the Start method is called and when the callback is invoked. Default is 100. Max is 2,147,483,647 milliseconds, or 24 days 20 hours 31 minutes 23.647 seconds.  |  |  | Property | Timer | VBScripting |
| Callback | Gets or sets a reference to the VBScript Sub that is called when the interval has elapsed.  |  |  | Property | Timer | VBScripting |
| AutoReset | Gets or sets a boolean determining whether to repeatedly invoke the callback. Default is False. If False, the callback is invoked only once, until the timer is restarted with the Start method.  |  |  | Property | Timer | VBScripting |
| IntervalInHours | Gets or sets the interval in hours.  |  |  | Property | Timer | VBScripting |

## Watcher

| Member name | Remarks | Returns | Parameters | Kind | Member of | Namespace |
| :---------- | :------ | :------ | :--------- | :--- | :-------- | :-------- |
| Watcher | Provides something like presentation mode for Windows systems that don't have presentation.exe: A way to temporarily keep the couputer from going to sleep. Uses <a href="https://docs.microsoft.com/en-us/windows/desktop/api/winbase/nf-winbase-setthreadexecutionstate"> SetThreadExecutionState</a>. Adapted from <a href="https://stackoverflow.com/questions/6302185/how-to-prevent-windows-from-entering-idle-state"> stackoverflow.com</a> and <a href="http://www.pinvoke.net/default.aspx/kernel32.setthreadexecutionstate"> pinvoke.net</a> posts. |  |  | Type | | VBScripting |
| IWatcher | The COM interface for VBScripting.Watcher  |  |  | Type | | VBScripting |
| (Constructor) | Constructor. Starts a private timer that periodically resets the system idle timer with the desired state.  |  |  | Method | Watcher | VBScripting |
| Dispose | Disposes of the object's resources.  |  |  | Method | Watcher | VBScripting |
| MonitorOff | Turn off the monitor(s).  |  |  | Method | Watcher | VBScripting |
| Watch | Gets or sets whether the system and monitor(s) should be kept from going into a suspend (sleep) state. The computer may still be put to sleep by other applications or by user actions such as closing a laptop lid or pressing a sleep button or power button. Default is False.  |  |  | Property | Watcher | VBScripting |
| CurrentState | Gets or sets an integer describing the current thread execution state. Intended for internal use and testing only.  |  |  | Property | Watcher | VBScripting |
| ResetPeriod | Gets or sets the time in milliseconds between idle-timer resets. Optional. Default is 30000. Max 2147483647.  |  |  | Property | Watcher | VBScripting |
| Privileged | Gets a boolean indicating whether privileges are elevated.  |  |  | Property | Watcher | VBScripting |
