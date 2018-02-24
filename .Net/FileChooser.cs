
// wrap the OpenFileDialog class for VBScript

using System.Windows.Forms; // for OpenFileDialog
using System.Runtime.InteropServices;
using System.Reflection;
using System.Linq; // for Cast<object>
using System.IO; // for Path

namespace VBScripting
{
    /// <summary> Provides a file chooser dialog for VBScript. </summary>
    [ProgId("VBScripting.FileChooser"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-3AF8-495F-AB4D-6C61BD463EA4")]
    public class FileChooser : IFileChooser
    {
        private OpenFileDialog chooser;
        private string _initialDirectory;
        private string _expandedResolvedInitialDirectory;

        /// <summary> Constructor </summary>
        public FileChooser()
        {
            this.chooser = new OpenFileDialog();
            this.Filter = "All files (*.*)|*.*";
            this.FilterIndex = 1;
            this.Title = "Browse for a file";
            this.DereferenceLinks = false;
            this.DefaultExt = "txt";
            this.InitialDirectory = System.Environment.CurrentDirectory;
        }

        /// <summary> Opens a dialog enabling the user to browse for and choose a file. </summary>
        /// <remarks> Returns the filespec of the chosen file. Returns an empty string if the user cancels.  </remarks> 
        public string FileName
        {
            get
            {
                DialogResult result = chooser.ShowDialog();

                if (result == DialogResult.OK)
                {
                    return chooser.FileName;
                }
                else
                {
                    return string.Empty;
                }
            }
            set
            {
                chooser.FileName = value;
            }
        }

        /// <summary> Opens a dialog enabling the user to browse for and choose multiple files. </summary>
        /// <remarks> Gets a string array of filespecs. Returns an empty array if the user cancels. Requires Multiselect to have been set to True. </remarks>
        public object FileNames
        {
            get
            {
                DialogResult result = chooser.ShowDialog();

                if (result == DialogResult.OK)
                {
                    // convert C# array to VBScript array
                    return chooser.FileNames.Cast<object>().ToArray();
                }
                else
                {
                    return new string[] { }.Cast<object>().ToArray();
                }
            }
        }

        /// <summary> Opens a dialog enabling the user to browse for and choose multiple files. </summary>
        /// <remarks> Gets a string of filespecs delimited by a vertical bar (|). Returns an empty string if the user cancels. Requires Multiselect to have been set to True. </remarks>
        public string FileNamesString
        {
            get
            {
                DialogResult result = chooser.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string[] names = this.chooser.FileNames;
                    return string.Join("|", names, 0, names.Length);
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        /// <summary> Gets or sets directory at which the dialog opens. </summary>
        public string InitialDirectory
        {
            get
            {
                return this._initialDirectory;
            }
            set
            {
                this._initialDirectory = value;
                this.ExpandedResolvedInitialDirectory = value;
                this.chooser.InitialDirectory = ExpandedResolvedInitialDirectory;
            }
        }
        /// <summary> Gets the initial directory with relative path resolved and environment variables expanded. </summary>
        /// <remarks> Improves testability. </remarks>
        public string ExpandedResolvedInitialDirectory
        {
            get
            {
                return this._expandedResolvedInitialDirectory;
            }
            private set
            {
                this._expandedResolvedInitialDirectory = Path.GetFullPath(System.Environment.ExpandEnvironmentVariables(value));
            }
        }
        /// <summary> Gets or sets the selectable file types.  </summary>
        /// <remarks> 
        /// Examples: <pre> fc.Filter = "All files (*.*)|*.*" // the default <br /> fc.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*" <br /> fc.Filter = "Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*" </pre>
        /// </remarks> 
        public string Filter
        {
            get
            {
                return this.chooser.Filter;
            }
            set
            {
                this.chooser.Filter = value;
            }
        }

        /// <summary> Gets or sets the index controlling which filter item is initially selected. </summary>
        /// <remarks> An integer. The index is 1-based. The default is 1. </remarks>
        public int FilterIndex
        {
            get
            {
                return this.chooser.FilterIndex;
            }
            set
            {
                this.chooser.FilterIndex = value;
            }
        }

        /// <summary> Gets or sets the dialog titlebar text. </summary>
        /// <remarks> The default text is "Browse for a file." </remarks>
        public string Title
        {
            get
            {
                return this.chooser.Title;
            }
            set
            {
                this.chooser.Title = value;
            }
        }

        /// <summary> Gets or sets whether multiple files can be selected. </summary> 
        /// <remarks> The default is False. </remarks>
        public bool Multiselect
        {
            get
            {
                return this.chooser.Multiselect;
            }
            set
            {
                this.chooser.Multiselect = value;
            }
        }

        /// <summary> Indicates whether the returned file is the referenced file or the .lnk file itself. </summary>
        /// <remarks> Gets or sets, if the selected file is a .lnk file, whether the filespec returned refers to the .lnk file itself (False) or to the file that the .lnk file points to (True). The default is False. </remarks>
        public bool DereferenceLinks
        {
            get
            {
                return this.chooser.DereferenceLinks;
            }
            set
            {
                this.chooser.DereferenceLinks = value;
            }
        }

        /// <summary> Gets or sets the file extension name that is automatically supplied when one is not specified. </summary>
        /// <remarks> A string. The default is "txt". </remarks>
        public string DefaultExt
        {
            get
            {
                return this.chooser.DefaultExt;
            }
            set
            {
                this.chooser.DefaultExt = value;
            }
        }

        /// <summary> Gets or sets whether to validate the file name(s). </summary>
        public bool ValidateNames
        {
            get
            {
                return this.chooser.ValidateNames;
            }
            set
            {
                this.chooser.ValidateNames = value;
            }
        }

        /// <summary> Gets or sets whether to check that the file exists. </summary>
        public bool CheckFileExists
        {
            get
            {
                return this.chooser.CheckFileExists;
            }
            set
            {
                this.chooser.CheckFileExists = value;
            }
        }
    }

    /// <summary> The COM interface for FileChooser </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-3BF8-495F-AB4D-6C61BD463EA4")]
    public interface IFileChooser
    {
        /// <summary> </summary>
        [DispId(1)]
        string FileName { get; set; }

        /// <summary> </summary>
        [DispId(2)]
        object FileNames { get; }

        /// <summary> </summary>
        [DispId(3)]
        string FileNamesString { get; }

        /// <summary> </summary>
        [DispId(4)]
        string Filter { get; set; }

        /// <summary> </summary>
        [DispId(5)]
        int FilterIndex { get; set; }

        /// <summary> </summary>
        [DispId(6)]
        string Title { get; set; }

        /// <summary> </summary>
        [DispId(7)]
        bool Multiselect { get; set; }

        /// <summary> </summary>
        [DispId(8)]
        bool DereferenceLinks { get; set; }

        /// <summary> </summary>
        [DispId(9)]
        string DefaultExt { get; set; }

        /// <summary> </summary>
        [DispId(10)]
        bool ValidateNames { get; set; }

        /// <summary> </summary>
        [DispId(11)]
        bool CheckFileExists { get; set; }

        /// <summary> </summary>
        [DispId(12)]
        string InitialDirectory { get; set; }

        /// <summary> </summary>
        [DispId(12)]
        string ExpandedResolvedInitialDirectory { get; }
    }
}
