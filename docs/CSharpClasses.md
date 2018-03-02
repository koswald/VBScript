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


## Admin

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| Admin| | Provide sys admin features. | | | Type| VBScripting |
| IAdmin| | COM interface for VBScripting.Admin | | | Type| VBScripting |
| Log| Admin| Logs the specified message to the event log. | | message| Method| VBScripting |
| GetLogs| Admin| Get an array of logs entries from the Application log. Returns an array of logs (strings) from the specified event source that contain the specified message string. Searches the Application log only.| an array| source, message| Method| VBScripting |
| SourceExists| Admin| Gets whether the specified EventLog source exists. | a boolean| source| Method| VBScripting |
| CreateEventSource| Admin| Creates the specified EventLog source. | an EventLogSourceResult| source| Method| VBScripting |
| DeleteEventSource| Admin| Deletes the specified EventLog source and all of its logs. | an EventLogSourceResult| source| Method| VBScripting |
| PrivilegesAreElevated| Admin| Gets whether the current process has elevated privileges. | | | Property| VBScripting |
| EventSource| Admin| Gets the name of the event log source for this namespace (VBScripting). | a string| | Property| VBScripting |
| LogName| Admin| Gets the name of the log to which events will be logged. | a string| | Property| VBScripting |
| Result| Admin| Gets an EventLogResultT object. VBScript example: <pre> Set returnValue = adm.CreateEventSource <br/> If returnValue.Result = adm.Result.SourceCreationException Then <br/>     MsgBox returnValue.Message <br/> End If</pre>| an EventLogResultT| | Property| VBScripting |
| EventLogResultT| | Provides a set of terse behavior/result descriptions suitable for VBScript comparisons and MsgBox captions. Not directly available to VBScript. See <tt>Admin.Result</tt>.| | | Type| VBScripting |
| SourceAlreadyExists| EventLogResultT|  | "Source already exists"| | Property| VBScripting |
| SourceCreated| EventLogResultT|  | "Source created"| | Property| VBScripting |
| SourceCreationException| EventLogResultT|  | "Source creation error"| | Property| VBScripting |
| SourceDoesNotExist| EventLogResultT|  | "Source does not exist"| | Property| VBScripting |
| SourceDeleted| EventLogResultT|  | "Source deleted"| | Property| VBScripting |
| SourceDeletionException| EventLogResultT|  | "Source deletion error"| | Property| VBScripting |
| EventLogSourceResult| | Type returned by CreateEventSource and DeleteEventSource. | | | Type| VBScripting |
| SourceExists| EventLogSourceResult| Returns True if the source exists after the attempted operation has completed. | a boolean| | Property| VBScripting |
| Message| EventLogSourceResult| Returns a message descriptive of the outcome of the operation. | a string| | Property| VBScripting |
| Result| EventLogSourceResult| Returns a string: one of the EventLogResultT strings. | a string| | Property| VBScripting |

## ComEvent

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| ComEvent| | Invokes VBS methods from C#. <span class="red"> This class is not callable from VBScript. </span> | | | Type| VBScripting |
| InvokeComCallback| ComEvent| Invokes a VBScript method. The parameter <tt>callbackRef</tt> is an object reference to a VBScript member returned by the VBScript Function GetRef.| | | Method| VBScripting |

## EventLogger

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| IEventLogger| | A COM Interface for VBScripting.EventLogger. | | | Type| VBScripting |
| EventLogger| | Provides system logging for VBScript. | | | Type| VBScripting |
| log| EventLogger| Writes the specified message to the Application event log. | | message| Method| VBScripting |

## FileChooser

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| FileChooser| | Provides a file chooser dialog for VBScript. | | | Type| VBScripting |
| IFileChooser| | The COM interface for FileChooser | | | Type| VBScripting |
| (Constructor)| FileChooser| Constructor | | | Method| VBScripting |
| FileName| FileChooser| Opens a dialog enabling the user to browse for and choose a file. Returns the filespec of the chosen file. Returns an empty string if the user cancels.| | | Property| VBScripting |
| FileNames| FileChooser| Opens a dialog enabling the user to browse for and choose multiple files. Gets a string array of filespecs. Returns an empty array if the user cancels. Requires Multiselect to have been set to True.| | | Property| VBScripting |
| FileNamesString| FileChooser| Opens a dialog enabling the user to browse for and choose multiple files. Gets a string of filespecs delimited by a vertical bar (&#124;). Returns an empty string if the user cancels. Requires Multiselect to have been set to True.| | | Property| VBScripting |
| InitialDirectory| FileChooser| Gets or sets directory at which the dialog opens. | | | Property| VBScripting |
| ERInitialDirectory| FileChooser| Gets the initial directory with relative path resolved and environment variables expanded. Improves testability.| | | Property| VBScripting |
| Filter| FileChooser| Gets or sets the selectable file types. Examples: <pre> fc.Filter = "All files (&#42;.&#42;)&#124;&#42;.&#42;" // the default <br/> fc.Filter = "Text files (&#42;.txt)&#124;&#42;.txt&#124;All files (&#42;.&#42;)&#124;&#42;.&#42;" <br/> fc.Filter = "Image Files(&#42;.BMP;&#42;.JPG;&#42;.GIF)&#124;&#42;.BMP;&#42;.JPG;&#42;.GIF&#124;All files (&#42;.&#42;)&#124;&#42;.&#42;" </pre>| | | Property| VBScripting |
| FilterIndex| FileChooser| Gets or sets the index controlling which filter item is initially selected. An integer. The index is 1-based. The default is 1.| | | Property| VBScripting |
| Title| FileChooser| Gets or sets the dialog titlebar text. The default text is "Browse for a file."| | | Property| VBScripting |
| Multiselect| FileChooser| Gets or sets whether multiple files can be selected. The default is False.| | | Property| VBScripting |
| DereferenceLinks| FileChooser| Indicates whether the returned file is the referenced file or the .lnk file itself. Gets or sets, if the selected file is a .lnk file, whether the filespec returned refers to the .lnk file itself (False) or to the file that the .lnk file points to (True). The default is False.| | | Property| VBScripting |
| DefaultExt| FileChooser| Gets or sets the file extension name that is automatically supplied when one is not specified. A string. The default is "txt".| | | Property| VBScripting |
| ValidateNames| FileChooser| Gets or sets whether to validate the file name(s). | | | Property| VBScripting |
| CheckFileExists| FileChooser| Gets or sets whether to check that the file exists. | | | Property| VBScripting |

## FolderChooser

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| IFolderChooser| | COM interface for FolderChooser. | | | Type| VBScripting |
| FolderChooser| | Present the Windows Vista-style open file dialog to select a folder. Fall back for older Windows Versions. Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/57611/erike"> EricE</a>. Uses <tt> System.Reflection</tt>.| | | Type| VBScripting |
| InitialDirectory| FolderChooser| Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory. | | | Property| VBScripting |
| Title| FolderChooser| Gets or sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder". | | | Property| VBScripting |
| FolderName| FolderChooser| Opens a dialog and returns the folder selected by the user. | a path| | Property| VBScripting |

## FolderChooser2

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| IFolderChooser2| | COM interface for FolderChooser2. | | | Type| VBScripting |
| FolderChooser2| | Present the Windows Vista-style open file dialog to select a folder. Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/15368771/show-detailed-folder-browser-from-a-propertygrid#15386992"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/403671/simon-mourier"> Simon Mourier</a>. Uses <tt> System.Runtime.InteropServices</tt>.| | | Type| VBScripting |
| InitialDirectory| FolderChooser2| Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory.| | | Property| VBScripting |
| Title| FolderChooser2| Sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder". | | | Property| VBScripting |
| FolderName| FolderChooser2| Opens a dialog and returns the folder selected by the user. | a path| | Property| VBScripting |

## IconExtractor

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| IconExtractor| | Extracts an icon from a .dll or .exe file. <span class="red"> This class is not accessible to VBScript. </span>| | | Type| VBScripting |
| Extract| IconExtractor| Extracts an icon from the specified .dll or .exe file. Other parameters: <tt>number</tt> is an integer that specifies the icon's index within the resource. <tt>largeIcon</tt> is a boolean that specifies whether the icon should be a large icon or small icon.| an icon| file, number, largeIcon| Method| VBScripting |

## NotifyIcon

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| NotifyIcon| | Provides a system tray icon for VBScript, for illustration purposes. | | | Type| VBScripting |
| INotifyIcon| | The COM interface for NotifyIcon. | | | Type| VBScripting |
| (Constructor)| NotifyIcon| Constructor | | | Method| VBScripting |
| SetIconByIcoFile| NotifyIcon| Sets the system tray icon given an .ico file. The parameter <tt>fileName</tt> specifies the filespec of the .ico file. Environment variables and relative paths are allowed.| | fileName| Method| VBScripting |
| SetIconByDllFile| NotifyIcon| Sets the system tray icon from a .dll or .exe file. Parameters: <tt>fileName</tt> is the path and name of a .dll or .exe file that contains icons. <tt>index</tt> is an integer that specifies which icon to use. <tt>largeIcon</tt> is a boolean that specifies whether to use a large or small icon.| | fileName, index, largeIcon| Method| VBScripting |
| SetBalloonTipIcon| NotifyIcon| Sets the icon of the "balloon tip" or notification. The parameter <tt>type</tt> is an integer that specifies which icon to use: Return values of ToolTipIcon properties can be used: Error = 1, Info = 2, None = 3, Warning = 4.| | type| Method| VBScripting |
| Dispose| NotifyIcon| Disposes of the icon resources when it is no longer needed. If this method is not called, the icon may persist in the system tray until the mouse hovers over it, even after the object instance has lost scope.| | | Method| VBScripting |
| ShowBalloonTip| NotifyIcon| Show the balloon tip. | | | Method| VBScripting |
| AddMenuItem| NotifyIcon| Add a menu item to the system tray icon's context menu. This method can be called only from VBScript. The parameter <tt>menuText</tt> is a string that specifies the text that appears in the menu. The parameter <tt>callbackRef</tt> is a VBScript object reference returned by the VBScript GetRef Function.| | menuText, callbackRef| Method| VBScripting |
| InvokeCallbackByIndex| NotifyIcon| Provide callback testability from VBScript. | | | Method| VBScripting |
| ShowContextMenu| NotifyIcon| Show the context menu. Public in order to provide testability from VBScript.| | | Method| VBScripting |
| SetBalloonTipCallback| NotifyIcon| Sets the VBScript callback Sub or Function reference. VBScript example: <pre>    obj.SetBalloonTipCallback GetRef("ProcedureName") </pre>| | | Method| VBScripting |
| Text| NotifyIcon| Gets or sets the text shown when the mouse hovers over the system tray icon. | | | Property| VBScripting |
| Visible| NotifyIcon| Gets or sets the icon's visibility. A boolean. Required. Set this property to True after initializing other settings.| | | Property| VBScripting |
| BalloonTipTitle| NotifyIcon| Gets or sets the title of the "balloon tip" or notification. | | | Property| VBScripting |
| BalloonTipText| NotifyIcon| Gets or sets the text of the "balloon tip" or notification. | | | Property| VBScripting |
| BalloonTipLifetime| NotifyIcon| Gets or sets the lifetime of the "balloon tip" or notification. An integer (milliseconds). Deprecated as of Windows Vista, the value is overridden by accessibility settings. | | | Property| VBScripting |
| ToolTipIcon| NotifyIcon| Gets an object useful in VBScript for selecting a ToolTipIcon type. The properties Error, Info, None, and Warning may be used with SetBalloonTipIcon. VBScript example: <pre>    obj.SetBallonTipIcon obj.ToolTipIcon.Warning </pre>| a ToolTipIconT| | Property| VBScripting |
| ToolTipIconT| | Supplies the type required by NotifyIcon.ToolTipIcon This class is not directly accessible from VBScript , however, it is accessible via the <tt>NotifyIcon.ToolTipIcon</tt> property.| | | Type| VBScripting |
| Error| ToolTipIconT|  | 1| | Property| VBScripting |
| Info| ToolTipIconT|  | 2| | Property| VBScripting |
| None| ToolTipIconT|  | 3| | Property| VBScripting |
| Warning| ToolTipIconT|  | 4| | Property| VBScripting |
| CallbackEventSettings| | Settings for saving VBScript method references. This class is not accessible from VBScript. | | | Type| VBScripting |
| (Constructor)| CallbackEventSettings| Constructor | | | Method| VBScripting |
| AddRef| CallbackEventSettings| Adds a CallbackReference instance reference to the list. | | callbackRef| Method| VBScripting |
| Refs| CallbackEventSettings| Gets or sets a list of callback references. | | | Property| VBScripting |
| CallbackReference| | An orderly way to save the index and callback reference for a single menu item. This class is not accessible to VBScript.| | | Type| VBScripting |
| (Constructor)| CallbackReference| Constructor | | index, reference| Method| VBScripting |
| Index| CallbackReference| This Index corresponds to the Index of a menuItem in the context menu. | | | Property| VBScripting |
| Reference| CallbackReference| COM object generated by VBScript's GetRef Function. | | | Property| VBScripting |

## ProgressBar

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| ProgressBar| | Supplies a progress bar to VBScript, for illustration purposes. | | | Type| VBScripting |
| IProgressBar| | Exposes the ProgressBar members to COM/VBScript. | | | Type| VBScripting |
| PerformStep| ProgressBar| Advances the progress bar one step. | | | Method| VBScripting |
| FormSize| ProgressBar| Sets the size of the window. | | width, height| Method| VBScripting |
| PBarSize| ProgressBar| Sets the size of the progress bar. | | width, height| Method| VBScripting |
| FormLocation| ProgressBar| Sets the position of the window. | | x, y| Method| VBScripting |
| FormLocationByPercentage| ProgressBar| Sets the position of the window. | | x, y| Method| VBScripting |
| PBarLocation| ProgressBar| Sets the location of the progress bar within the window. | | x, y| Method| VBScripting |
| SuspendLayout| ProgressBar| Suspends drawing of the window temporarily. | | | Method| VBScripting |
| ResumeLayout| ProgressBar| Resumes drawing the window. | | | Method| VBScripting |
| SetIconByIcoFile| ProgressBar| Sets the icon given the filespec of an .ico file. Environment variables are allowed.| | fileName| Method| VBScripting |
| SetIconByDllFile| ProgressBar| Sets the icon given the filespec of a .dll or .exe file and an index. The index is an integer that identifies the icon. Environment variables are allowed.| | fileName, index| Method| VBScripting |
| Dispose| ProgressBar| Disposes of the object's resources. | | | Method| VBScripting |
| Visible| ProgressBar| Gets or sets the progress bar's visibility. A boolean. The default is False.| | | Property| VBScripting |
| Minimum| ProgressBar| Gets or sets the value at which there is no apparent progress. An integer. The default is 0.| | | Property| VBScripting |
| Maximum| ProgressBar| Gets or sets the value at which the progress appears to be complete. An integer. The default is 100.| | | Property| VBScripting |
| Value| ProgressBar| Gets or sets the apparent progress. An integer. Should be at or above the minimum and at or below the maximum.| | | Property| VBScripting |
| Step| ProgressBar| Gets or sets the increment between steps. | | | Property| VBScripting |
| Caption| ProgressBar| Gets or sets the window title-bar text. | | | Property| VBScripting |
| Debug| ProgressBar| Gets or sets whether the type is under development. Affects the behavior of two methods, SetIconByIcoFile and SetIconByDllFile, if exceptions are thrown: when debugging, a message box is shown. Default is False.| | | Property| VBScripting |
| BorderStyle| ProgressBar| Provides an object useful in VBScript for setting FormBorderStyle. | a FormBorderStyleT| | Property| VBScripting |
| FormBorderStyle| ProgressBar| Sets the style of the window border. An integer. One of the BorderStyle property return values can be used: Fixed3D, FixedDialog, FixedSingle, FixedToolWindow, None, Sizable (default), or SizableToolWindow. VBScript example: <pre> pb.FormBorderStyle = pb.BorderStyle.Fixed3D </pre>| | | Property| VBScripting |
| FormBorderStyleT| | Enumeration of border styles. This class is available to VBScript via the <tt>ProgressBar.BorderStyle</tt> property.| | | Type| VBScripting |
| Fixed3D| FormBorderStyleT|  | 1| | Property| VBScripting |
| FixedDialog| FormBorderStyleT|  | 2| | Property| VBScripting |
| FixedSingle| FormBorderStyleT|  | 3| | Property| VBScripting |
| FixedToolWindow| FormBorderStyleT|  | 4| | Property| VBScripting |
| None| FormBorderStyleT|  | 5| | Property| VBScripting |
| Sizable| FormBorderStyleT|  | 6| | Property| VBScripting |
| SizableToolWindow| FormBorderStyleT|  | 7| | Property| VBScripting |

## SpeechSynthesis

| Member name | Member of | Remarks | Returns | Parameters | Kind | Namespace |
| :---------- | :-------- | :------ | :------ | :--------- | :--- | :-------- |
| SpeechSynthesis| | Provide a wrapper for the .Net speech synthesizer for VBScript, for demonstration purposes. Requires an assembly reference to <tt>%ProgramFiles(x86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll</tt>, which may not be available on older machines.| | | Type| VBScripting |
| ISpeechSynthesis| | The COM interface for <tt>VBScripting.SpeechSynthesis</tt>. | | | Type| VBScripting |
| (Constructor)| SpeechSynthesis| Constructor | | | Method| VBScripting |
| Speak| SpeechSynthesis| Convert text to speech. This method is synchronous.| | | Method| VBScripting |
| SpeakAsync| SpeechSynthesis| Convert text to speech. This method is asynchronous.| | | Method| VBScripting |
| Pause| SpeechSynthesis| Pause speech synthesis. | | | Method| VBScripting |
| Resume| SpeechSynthesis| Resume speech synthesis. | | | Method| VBScripting |
| Voices| SpeechSynthesis| Gets an array of the names of the installed, enabled voices. Each element of the array can be used to set <tt>Voice</tt>.| | | Method| VBScripting |
| Dispose| SpeechSynthesis| Disposes the SpeechSynthesis object's resources. | | | Method| VBScripting |
| Voice| SpeechSynthesis| Gets or sets the current voice by name. A string. One of the names from the <tt>Voices</tt> array.| | | Property| VBScripting |
| SynthesizerState| SpeechSynthesis| Gets the state of the SpeechSynthesizer. Read only. Returns an integer equal to one of the <tt>State</tt> method return values.| | | Property| VBScripting |
| Volume| SpeechSynthesis| Gets or sets the volume. An integer from 0 to 100.| | | Property| VBScripting |
| State| SpeechSynthesis| Gets an object whose properties (Ready, Paused, and Speaking) provide values useful for comparing to <tt>SynthesizerState</tt>. | a SynthersizerStateT| | Property| VBScripting |
| SynthesizerStateT| | Enumerates the synthesizer states. Not intended for use in VBScript. See <tt>SpeechSynthesis.State</tt>.| | | Type| VBScripting |
| (Constructor)| SynthesizerStateT| Constructor | | | Method| VBScripting |
| Ready| SynthesizerStateT|  | 1| | Property| VBScripting |
| Paused| SynthesizerStateT|  | 2| | Property| VBScripting |
| Speaking| SynthesizerStateT|  | 3| | Property| VBScripting |
| Unexpected| SynthesizerStateT|  | 4| | Property| VBScripting |
