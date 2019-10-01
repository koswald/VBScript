
// Provides a user-friendly folder chooser for VBSxript

// adapted from https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043
// by EricE https://stackoverflow.com/users/57611/erike

using System;
using System.Reflection;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace VBScripting
{
    /// <summary> COM interface for VBScripting.FolderChooser </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-B1F8-495F-AB4D-6C61BD463EA4")]
    public interface IFolderChooser
    {
        /// <summary> </summary>
        string InitialDirectory { get; set; }
        /// <summary> </summary>
        string Title { get; set; }
        /// <summary> </summary>
        string FolderName { get; }
    }
    /// <summary> Present the Windows Vista-style open file dialog to select a folder. Fall back for older Windows Versions. </summary>
    /// <remarks> Adapted from <a title="stackoverflow.com" href="https://stackoverflow.com/questions/11767/browse-for-a-directory-in-c-sharp#33817043"> a stackoverflow post</a> by <a title="stackoverflow.com" href="https://stackoverflow.com/users/57611/erike"> EricE</a>. Uses <tt> System.Reflection</tt>. </remarks>
    [ProgId("VBScripting.FolderChooser"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-B0F8-495F-AB4D-6C61BD463EA4")]
    public class FolderChooser : IFolderChooser
    {
        private string _initialDirectory;
        private string _title;
        private string _fileName = "";

        /// <summary> Gets or sets the initial directory that the folder select dialog opens to. Environment variables are allowed. Relative paths are allowed. Optional. The default value is the current directory. </summary>
        public string InitialDirectory
        {
            get { return string.IsNullOrEmpty(_initialDirectory) ? Environment.CurrentDirectory : _initialDirectory; }
            set { _initialDirectory = value; }
        }
        /// <summary> Gets or sets the title/caption of the folder select dialog. Optional. The default value is "Select a folder". </summary>
        public string Title
        {
            get { return string.IsNullOrEmpty(_title) ? "Select a folder" : _title; }
            set { _title = value; }
        }
        /// <summary> Opens a dialog and returns the folder selected by the user. </summary>
        /// <returns> a path </returns>
        public string FolderName
        {
            get { return Show()? _fileName : string.Empty; }
        }

        private bool Show() { return Show(IntPtr.Zero); }

        // <param name="hWndOwner">Handle of the control or window to be the parent of the file dialog</param>
        // <returns>true if the user clicks OK</returns>
        private bool Show(IntPtr hWndOwner)
        {
            var dir1 = Path.GetFullPath(Environment.ExpandEnvironmentVariables(InitialDirectory));
            var result = Environment.OSVersion.Version.Major >= 6
                ? VistaDialog.Show(hWndOwner, dir1, Title)
                : ShowXpDialog(hWndOwner, dir1, Title);
            _fileName = result.FileName;
            return result.Result;
        }

        private struct ShowDialogResult
        {
            public bool Result { get; set; }
            public string FileName { get; set; }
        }

        private static ShowDialogResult ShowXpDialog(IntPtr ownerHandle, string initialDirectory, string title)
        {
            var folderBrowserDialog = new FolderBrowserDialog
            {
                Description = title,
                SelectedPath = initialDirectory,
                ShowNewFolderButton = false
            };
            var dialogResult = new ShowDialogResult();
            if (folderBrowserDialog.ShowDialog(new WindowWrapper(ownerHandle)) == DialogResult.OK)
            {
                dialogResult.Result = true;
                dialogResult.FileName = folderBrowserDialog.SelectedPath;
            }
            return dialogResult;
        }

        private static class VistaDialog
        {
            private const string c_foldersFilter = "Folders|\n";

            private const BindingFlags c_flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            private readonly static Assembly s_windowsFormsAssembly = typeof(FileDialog).Assembly;
            private readonly static Type s_iFileDialogType = s_windowsFormsAssembly.GetType("System.Windows.Forms.FileDialogNative+IFileDialog");
            private readonly static MethodInfo s_createVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("CreateVistaDialog", c_flags);
            private readonly static MethodInfo s_onBeforeVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("OnBeforeVistaDialog", c_flags);
            private readonly static MethodInfo s_getOptionsMethodInfo = typeof(FileDialog).GetMethod("GetOptions", c_flags);
            private readonly static MethodInfo s_setOptionsMethodInfo = s_iFileDialogType.GetMethod("SetOptions", c_flags);
            private readonly static uint s_fosPickFoldersBitFlag = (uint)s_windowsFormsAssembly
                .GetType("System.Windows.Forms.FileDialogNative+FOS")
                .GetField("FOS_PICKFOLDERS")
                .GetValue(null);
            private readonly static ConstructorInfo s_vistaDialogEventsConstructorInfo = s_windowsFormsAssembly
                .GetType("System.Windows.Forms.FileDialog+VistaDialogEvents")
                .GetConstructor(c_flags, null, new[] { typeof(FileDialog) }, null);
            private readonly static MethodInfo s_adviseMethodInfo = s_iFileDialogType.GetMethod("Advise");
            private readonly static MethodInfo s_unAdviseMethodInfo = s_iFileDialogType.GetMethod("Unadvise");
            private readonly static MethodInfo s_showMethodInfo = s_iFileDialogType.GetMethod("Show");

            public static ShowDialogResult Show(IntPtr ownerHandle, string initialDirectory, string title)
            {
                var openFileDialog = new OpenFileDialog
                {
                    AddExtension = false,
                    CheckFileExists = false,
                    DereferenceLinks = true,
                    Filter = c_foldersFilter,
                    InitialDirectory = initialDirectory,
                    Multiselect = false,
                    Title = title
                };

                var iFileDialog = s_createVistaDialogMethodInfo.Invoke(openFileDialog, new object[] { });
                s_onBeforeVistaDialogMethodInfo.Invoke(openFileDialog, new[] { iFileDialog });
                s_setOptionsMethodInfo.Invoke(iFileDialog, new object[] { (uint)s_getOptionsMethodInfo.Invoke(openFileDialog, new object[] { }) | s_fosPickFoldersBitFlag });
                var adviseParametersWithOutputConnectionToken = new[] { s_vistaDialogEventsConstructorInfo.Invoke(new object[] { openFileDialog }), 0U };
                s_adviseMethodInfo.Invoke(iFileDialog, adviseParametersWithOutputConnectionToken);

                try
                {
                    int retVal = (int)s_showMethodInfo.Invoke(iFileDialog, new object[] { ownerHandle });
                    return new ShowDialogResult
                    {
                        Result = retVal == 0,
                        FileName = openFileDialog.FileName
                    };
                }
                finally
                {
                    s_unAdviseMethodInfo.Invoke(iFileDialog, new[] { adviseParametersWithOutputConnectionToken[1] });
                }
            }
        }

        private class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _handle;
            public WindowWrapper(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }
    }
}