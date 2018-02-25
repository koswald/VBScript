# C# Classes Documentation


## Admin

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| Admin| Type| | | | Provide sys admin features. | VBScripting |
| IAdmin| Type| | | | COM interface for VBScripting.Admin | VBScripting |
| Log| Method| Admin| | | Logs an event to the event log. | VBScripting |
| GetLogs| Method| Admin| source, message| an array| Get an array of logs entries from the Application log. Returns an array of logs (strings) from the specified event source that contain the specified message string. Searches the Application log only.| VBScripting |
| SourceExists| Method| Admin| source| a boolean| Gets whether the specified EventLog source exists. | VBScripting |
| CreateEventSource| Method| Admin| source| an EventLogSourceResult| Creates the specified EventLog source. | VBScripting |
| DeleteEventSource| Method| Admin| source| an EventLogSourceResult| Deletes the specified EventLog source and all of its logs. | VBScripting |
| PrivilegesAreElevated| Property| Admin| | | Gets whether the current process has elevated privileges. | VBScripting |
| EventSource| Property| Admin| | a string| Gets the name of the event log source for this namespace (VBScripting). | VBScripting |
| LogName| Property| Admin| | a string| Gets the name of the log to which events will be logged. | VBScripting |
| Result| Property| Admin| | an EventLogResultT| Gets an EventLogResultT object. VBScript example: <pre> Set returnValue = adm.CreateEventSource <br/> If returnValue.Result = adm.Result.SourceCreationException Then <br/>     MsgBox returnValue.Message <br/> End If</pre>| VBScripting |
| EventLogResultT| Type| | | | Provides a set of terse behavior/result descriptions suitable for VBScript comparisons and MsgBox captions. Not directly available to VBScript. See <tt>Admin.Result</tt>.| VBScripting |
| SourceAlreadyExists| Property| EventLogResultT| | "Source already exists"|  | VBScripting |
| SourceCreated| Property| EventLogResultT| | "Source created"|  | VBScripting |
| SourceCreationException| Property| EventLogResultT| | "Source creation error"|  | VBScripting |
| SourceDoesNotExist| Property| EventLogResultT| | "Source does not exist"|  | VBScripting |
| SourceDeleted| Property| EventLogResultT| | "Source deleted"|  | VBScripting |
| SourceDeletionException| Property| EventLogResultT| | "Source deletion error"|  | VBScripting |
| EventLogSourceResult| Type| | | | Type returned by CreateEventSource and DeleteEventSource. | VBScripting |
| SourceExists| Property| EventLogSourceResult| | a boolean| Returns True if the source exists after the attempted operation has completed. | VBScripting |
| Message| Property| EventLogSourceResult| | a string| Returns a message descriptive of the outcome of the operation. | VBScripting |
| Result| Property| EventLogSourceResult| | a string| Returns a string: one of the EventLogResultT strings. | VBScripting |

## ComEvent

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| ComEvent| Type| | | | Invokes VBS methods from C#. <span class="red"> This class is not callable from VBScript. </span> | VBScripting |
| InvokeComCallback| Method| ComEvent| | | Invokes a VBScript method. The parameter <tt>callbackRef</tt> is an object reference to a VBScript member returned by the VBScript Function GetRef.| VBScripting |

## EventLogger

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| IEventLogger| Type| | | | A COM Interface for VBScripting.EventLogger. | VBScripting |
| EventLogger| Type| | | | Provides system logging for VBScript. | VBScripting |
| log| Method| EventLogger| message| | Writes the specified message to the Application event log. | VBScripting |

## FileChooser

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| FileChooser| Type| | | | Provides a file chooser dialog for VBScript. | VBScripting |
| IFileChooser| Type| | | | The COM interface for FileChooser | VBScripting |
| (Constructor)| Method| FileChooser| | | Constructor | VBScripting |
| FileName| Property| FileChooser| | | Opens a dialog enabling the user to browse for and choose a file. Returns the filespec of the chosen file. Returns an empty string if the user cancels.| VBScripting |
| FileNames| Property| FileChooser| | | Opens a dialog enabling the user to browse for and choose multiple files. Gets a string array of filespecs. Returns an empty array if the user cancels. Requires Multiselect to have been set to True.| VBScripting |
| FileNamesString| Property| FileChooser| | | Opens a dialog enabling the user to browse for and choose multiple files. Gets a string of filespecs delimited by a vertical bar (\|). Returns an empty string if the user cancels. Requires Multiselect to have been set to True.| VBScripting |
| InitialDirectory| Property| FileChooser| | | Gets or sets directory at which the dialog opens. | VBScripting |
| ExpandedResolvedInitialDirectory| Property| FileChooser| | | Gets the initial directory with relative path resolved and environment variables expanded. Improves testability.| VBScripting |
| Filter| Property| FileChooser| | | Gets or sets the selectable file types. Examples: <pre> fc.Filter = "All files (*.*)\|*.*" // the default <br/> fc.Filter = "Text files (*.txt)\|*.txt\|All files (*.*)\|*.*" <br/> fc.Filter = "Image Files(*.BMP;*.JPG;*.GIF)\|*.BMP;*.JPG;*.GIF\|All files (*.*)\|*.*" </pre>
| VBScripting |
| FilterIndex| Property| FileChooser| | | Gets or sets the index controlling which filter item is initially selected. An integer. The index is 1-based. The default is 1.| VBScripting |
| Title| Property| FileChooser| | | Gets or sets the dialog titlebar text. The default text is "Browse for a file."| VBScripting |
| Multiselect| Property| FileChooser| | | Gets or sets whether multiple files can be selected. The default is False.| VBScripting |
| DereferenceLinks| Property| FileChooser| | | Indicates whether the returned file is the referenced file or the .lnk file itself. Gets or sets, if the selected file is a .lnk file, whether the filespec returned refers to the .lnk file itself (False) or to the file that the .lnk file points to (True). The default is False.| VBScripting |
| DefaultExt| Property| FileChooser| | | Gets or sets the file extension name that is automatically supplied when one is not specified. A string. The default is "txt".| VBScripting |
| ValidateNames| Property| FileChooser| | | Gets or sets whether to validate the file name(s). | VBScripting |
| CheckFileExists| Property| FileChooser| | | Gets or sets whether to check that the file exists. | VBScripting |

## FolderChooser

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| IFolderChooser| Type| | | | COM interface for FolderChooser. | VBScripting |
| FolderChooser| Type| | | | Present the Windows Vista-style open file dialog to select a folder. Fall back for older Windows Versions. Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/57611/erike"> EricE</a>. Uses <tt> System.Reflection</tt>.| VBScripting |
| InitialDirectory| Property| FolderChooser| | | Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory. | VBScripting |
| Title| Property| FolderChooser| | | Gets or sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder". | VBScripting |
| FolderName| Property| FolderChooser| | a path| Opens a dialog and returns the folder selected by the user. | VBScripting |

## FolderChooser2

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| IFolderChooser2| Type| | | | COM interface for FolderChooser2. | VBScripting |
| FolderChooser2| Type| | | | Present the Windows Vista-style open file dialog to select a folder. Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/15368771/show-detailed-folder-browser-from-a-propertygrid#15386992"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/403671/simon-mourier"> Simon Mourier</a>. Uses <tt> System.Runtime.InteropServices</tt>.| VBScripting |
| InitialDirectory| Property| FolderChooser2| | | Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory.| VBScripting |
| Title| Property| FolderChooser2| | | Sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder". | VBScripting |
| FolderName| Property| FolderChooser2| | a path| Opens a dialog and returns the folder selected by the user. | VBScripting |

## IconExtractor

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| IconExtractor| Type| | | | Extracts an icon from a .dll or .exe file. <span class="red"> This class is not accessible to VBScript. </span>| VBScripting |
| Extract| Method| IconExtractor| file, number, largeIcon| an icon| Extracts an icon from the specified .dll or .exe file. Other parameters: <tt>number</tt> is an integer that specifies the icon's index within the resource. <tt>largeIcon</tt> is a boolean that specifies whether the icon should be a large icon or small icon.| VBScripting |

## NotifyIcon

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| NotifyIcon| Type| | | | Provides a system tray icon for VBScript, for illustration purposes. | VBScripting |
| INotifyIcon| Type| | | | The COM interface for NotifyIcon. | VBScripting |
| (Constructor)| Method| NotifyIcon| | | Constructor | VBScripting |
| SetIconByIcoFile| Method| NotifyIcon| fileName| | Sets the system tray icon given an .ico file. The parameter <tt>fileName</tt> specifies the filespec of the .ico file. Environment variables and relative paths are allowed.| VBScripting |
| SetIconByDllFile| Method| NotifyIcon| fileName, index, largeIcon| | Sets the system tray icon from a .dll or .exe file. Parameters: <tt>fileName</tt> is the path and name of a .dll or .exe file that contains icons. <tt>index</tt> is an integer that specifies which icon to use. <tt>largeIcon</tt> is a boolean that specifies whether to use a large or small icon.| VBScripting |
| SetBalloonTipIcon| Method| NotifyIcon| type| | Sets the icon of the "balloon tip" or notification. The parameter <tt>type</tt> is an integer that specifies which icon to use: Return values of ToolTipIcon properties can be used: Error = 1, Info = 2, None = 3, Warning = 4.| VBScripting |
| Dispose| Method| NotifyIcon| | | Disposes of the icon resources when it is no longer needed. If this method is not called, the icon may persist in the system tray until the mouse hovers over it, even after the object instance has lost scope.| VBScripting |
| ShowBalloonTip| Method| NotifyIcon| | | Show the balloon tip. | VBScripting |
| AddMenuItem| Method| NotifyIcon| menuText, callbackRef| | Add a menu item to the system tray icon's context menu. This method can be called only from VBScript. The parameter <tt>&gt;menuText</tt> is a string that specifies the text that appears in the menu. The parameter <tt>callbackRef</tt> is a VBScript object reference returned by the VBScript GetRef Function.| VBScripting |
| InvokeCallbackByIndex| Method| NotifyIcon| | | Provide callback testability from VBScript. | VBScripting |
| ShowContextMenu| Method| NotifyIcon| | | Show the context menu. Public in order to provide testability from VBScript.| VBScripting |
| SetBalloonTipCallback| Method| NotifyIcon| | | Sets the VBScript callback Sub or Function reference. VBScript example: <pre>    obj.SetBalloonTipCallback GetRef("ProcedureName") </pre>
| VBScripting |
| Text| Property| NotifyIcon| | | Gets or sets the text shown when the mouse hovers over the system tray icon. | VBScripting |
| Visible| Property| NotifyIcon| | | Gets or sets the icon's visibility. A boolean. Required. Set this property to True after initializing other settings.| VBScripting |
| BalloonTipTitle| Property| NotifyIcon| | | Gets or sets the title of the "balloon tip" or notification. | VBScripting |
| BalloonTipText| Property| NotifyIcon| | | Gets or sets the text of the "balloon tip" or notification. | VBScripting |
| BalloonTipLifetime| Property| NotifyIcon| | | Gets or sets the lifetime of the "balloon tip" or notification. An integer (milliseconds). Deprecated as of Windows Vista, the value is overridden by accessibility settings. | VBScripting |
| ToolTipIcon| Property| NotifyIcon| | a ToolTipIconT| Gets an object useful in VBScript for selecting a ToolTipIcon type. The properties Error, Info, None, and Warning may be used with SetBalloonTipIcon. VBScript example: <pre>    obj.SetBallonTipIcon obj.ToolTipIcon.Warning </pre>
| VBScripting |
| ToolTipIconT| Type| | | | Supplies the type required by NotifyIcon.ToolTipIcon This class is not directly accessible from VBScript , however, it is accessible via the <tt>NotifyIcon.ToolTipIcon</tt> property.| VBScripting |
| Error| Property| ToolTipIconT| | 1|  | VBScripting |
| Info| Property| ToolTipIconT| | 2|  | VBScripting |
| None| Property| ToolTipIconT| | 3|  | VBScripting |
| Warning| Property| ToolTipIconT| | 4|  | VBScripting |
| CallbackEventSettings| Type| | | | Settings for saving VBScript method references. This class is not accessible from VBScript. | VBScripting |
| (Constructor)| Method| CallbackEventSettings| | | Constructor | VBScripting |
| AddRef| Method| CallbackEventSettings| callbackRef| | Adds a CallbackReference instance reference to the list. | VBScripting |
| Refs| Property| CallbackEventSettings| | | Gets or sets a list of callback references. | VBScripting |
| CallbackReference| Type| | | | An orderly way to save the index and callback reference for a single menu item. This class is not accessible to VBScript.| VBScripting |
| (Constructor)| Method| CallbackReference| index, reference| | Constructor | VBScripting |
| Index| Property| CallbackReference| | | This Index corresponds to the Index of a menuItem in the context menu. | VBScripting |
| Reference| Property| CallbackReference| | | COM object generated by VBScript's GetRef Function. | VBScripting |

## ProgressBar

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| ProgressBar| Type| | | | Supplies a progress bar to VBScript, for illustration purposes. | VBScripting |
| IProgressBar| Type| | | | Exposes the ProgressBar members to COM/VBScript. | VBScripting |
| PerformStep| Method| ProgressBar| | | Advances the progress bar one step. | VBScripting |
| FormSize| Method| ProgressBar| width, height| | Sets the size of the window. | VBScripting |
| PBarSize| Method| ProgressBar| width, height| | Sets the size of the progress bar. | VBScripting |
| FormLocation| Method| ProgressBar| x, y| | Sets the position of the window. | VBScripting |
| FormLocationByPercentage| Method| ProgressBar| x, y| | Sets the position of the window. | VBScripting |
| PBarLocation| Method| ProgressBar| x, y| | Sets the location of the progress bar within the window. | VBScripting |
| SuspendLayout| Method| ProgressBar| | | Suspends drawing of the window temporarily. | VBScripting |
| ResumeLayout| Method| ProgressBar| | | Resumes drawing the window. | VBScripting |
| SetIconByIcoFile| Method| ProgressBar| fileName| | Sets the icon given the filespec of an .ico file. Environment variables are allowed.| VBScripting |
| SetIconByDllFile| Method| ProgressBar| fileName, index| | Sets the icon given the filespec of a .dll or .exe file and an index. The index is an integer that identifies the icon. Environment variables are allowed.| VBScripting |
| Dispose| Method| ProgressBar| | | Disposes of the object's resources. | VBScripting |
| Visible| Property| ProgressBar| | | Gets or sets the progress bar's visibility. A boolean. The default is False.| VBScripting |
| Minimum| Property| ProgressBar| | | Gets or sets the value at which there is no apparent progress. An integer. The default is 0.| VBScripting |
| Maximum| Property| ProgressBar| | | Gets or sets the value at which the progress appears to be complete. An integer. The default is 100.| VBScripting |
| Value| Property| ProgressBar| | | Gets or sets the apparent progress. An integer. Should be at or above the minimum and at or below the maximum.| VBScripting |
| Step| Property| ProgressBar| | | Gets or sets the increment between steps. | VBScripting |
| Caption| Property| ProgressBar| | | Gets or sets the window title-bar text. | VBScripting |
| Debug| Property| ProgressBar| | | Gets or sets whether the type is under development. Affects the behavior of two methods, SetIconByIcoFile and SetIconByDllFile, if exceptions are thrown: when debugging, a message box is shown. Default is False.| VBScripting |
| BorderStyle| Property| ProgressBar| | a FormBorderStyleT| Provides an object useful in VBScript for setting FormBorderStyle. | VBScripting |
| FormBorderStyle| Property| ProgressBar| | | Sets the style of the window border. An integer. One of the BorderStyle property return values can be used: Fixed3D, FixedDialog, FixedSingle, FixedToolWindow, None, Sizable (default), or SizableToolWindow. VBScript example: <pre> pb.FormBorderStyle = pb.BorderStyle.Fixed3D </pre>
| VBScripting |
| FormBorderStyleT| Type| | | | Enumeration of border styles. This class is available to VBScript via the <tt>ProgressBar.BorderStyle</tt> property.| VBScripting |
| Fixed3D| Property| FormBorderStyleT| | 1|  | VBScripting |
| FixedDialog| Property| FormBorderStyleT| | 2|  | VBScripting |
| FixedSingle| Property| FormBorderStyleT| | 3|  | VBScripting |
| FixedToolWindow| Property| FormBorderStyleT| | 4|  | VBScripting |
| None| Property| FormBorderStyleT| | 5|  | VBScripting |
| Sizable| Property| FormBorderStyleT| | 6|  | VBScripting |
| SizableToolWindow| Property| FormBorderStyleT| | 7|  | VBScripting |

## SpeechSynthesis

| Member name | Kind | Member of | Parameters | Returns | Remarks | Namespace |
| :--------------------- | :------ | :---------------- | :----------------- | :----------- | :------------- | :----------------- |
| SpeechSynthesis| Type| | | | Provide a wrapper for the .Net speech synthesizer for VBScript, for demonstration purposes. Requires an assembly reference to <tt>%ProgramFiles(x86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll</tt>, which may not be available on older machines.| VBScripting |
| ISpeechSynthesis| Type| | | | The COM interface for <tt>VBScripting.SpeechSynthesis</tt>. | VBScripting |
| (Constructor)| Method| SpeechSynthesis| | | Constructor | VBScripting |
| Speak| Method| SpeechSynthesis| | | Convert text to speech. This method is synchronous.| VBScripting |
| SpeakAsync| Method| SpeechSynthesis| | | Convert text to speech. This method is asynchronous.| VBScripting |
| Pause| Method| SpeechSynthesis| | | Pause speech synthesis. | VBScripting |
| Resume| Method| SpeechSynthesis| | | Resume speech synthesis. | VBScripting |
| Voices| Method| SpeechSynthesis| | | Gets an array of the names of the installed, enabled voices. Each element of the array can be used to set <tt>Voice</tt>
| VBScripting |
| Dispose| Method| SpeechSynthesis| | | Disposes the SpeechSynthesis object's resources. | VBScripting |
| Voice| Property| SpeechSynthesis| | | Gets or sets the current voice by name. A string. One of the names from the <tt>Voices</tt> array.| VBScripting |
| SynthesizerState| Property| SpeechSynthesis| | | Gets the state of the SpeechSynthesizer. Read only. Returns an integer equal to one of the <tt>State</tt> method return values.| VBScripting |
| Volume| Property| SpeechSynthesis| | | Gets or sets the volume. An integer from 0 to 100.| VBScripting |
| State| Property| SpeechSynthesis| | a SynthersizerStateT| Gets an object whose properties (Ready, Paused, and Speaking) provide values useful for comparing to <tt>SynthesizerState</tt>. | VBScripting |
| SynthesizerStateT| Type| | | | Enumerates the synthesizer states. Not intended for use in VBScript. See <tt>SpeechSynthesis.State</tt>.| VBScripting |
| (Constructor)| Method| SynthesizerStateT| | | Constructor | VBScripting |
| Ready| Property| SynthesizerStateT| | 1|  | VBScripting |
| Paused| Property| SynthesizerStateT| | 2|  | VBScripting |
| Speaking| Property| SynthesizerStateT| | 3|  | VBScripting |
| Unexpected| Property| SynthesizerStateT| | 4|  | VBScripting |
