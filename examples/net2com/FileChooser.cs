
// wrap the OpenFileDialog class for VBScript

// OpenFileDialog class:
// https://msdn.microsoft.com/en-us/library/system.windows.forms.openfiledialog(v=vs.110).aspx

using System.Windows.Forms; // for OpenFileDialog
using System.Runtime.InteropServices;
using System.Reflection;
using System.Linq; // for FileNames Cast

[assembly:AssemblyKeyFileAttribute("FileChooser.snk")]

namespace FileChooser
{
    [Guid("2650C2AB-3AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("FileChooser")]
    public class FileChooser : IFileChooser
    {
        private OpenFileDialog chooser;

        // Constructor
        public FileChooser()
        {
            this.chooser = new OpenFileDialog();
            this.Filter = "All files (*.*)|*.*";
            this.FilterIndex = 1;
            this.Title = "Browse for a file";
            this.DereferenceLinks = false;
            this.DefaultExt = "txt";
        }

        // Property: FileName
        // Returns: a filespec string
        // Remark: Opens a dialog enabling the user to browse for and choose a local file. Returns an empty string if the user cancels.
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
            private set { }
        }

        // Property: FileNames
        // Returns: a string array of filespecs
        // Remark: Opens a dialog enabling the user to browse for and choose multiple local files. Requires Multiselect to have been set to `True`. Returns an empty array if the user cancels.
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
            private set { }
        }

        // Property: FileNamesString
        // Returns: a `|` delimited string of filespecs
        // Remark: Opens a dialog enabling the user to browse for and choose multiple local files. Requires Multiselect to have been set to `True`. Returns an empty string if the user cancels.
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
            private set { }
        }

        // Property: Filter
        // Parameter: a filter string
        // Remark: Sets or gets the selectable file type options. The default is `All files (*.*)|*.*` Example #1: `Text files (*.txt)|*.txt|All files (*.*)|*.*`  Example #2: `Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*`
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

        // Property: FilterIndex
        // Paramater: index (int)
        // Remark: Sets or gets the index controlling which filter item is initially selected. The index is 1-based. The default is 1.
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

        // Property: Title
        // Parameter: title (string)
        // Remark: Sets or gets the titlebar text of the dialog window. The default is `Browse for a file`.
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

        // Property: setMultiselect
        // Parameter: a boolean
        // Remark: Sets or gets whether multiple files can be selected. Default is `False`.
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

        // Property: DereferenceLinks
        // Parameter: a boolean
        // Remark: Sets or gets, if the selected file is a `.lnk` file, whether the filespec returned refers to the `.lnk` file itself (`False`) or to the file that the `.lnk` file points to (`True`). The default is `False`.
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

        // Property: DefaultExt
        // Parameter: ext (string)
        // Remark: Sets or gets the file extension name that is automatically supplied when one is not specified. The default is `txt`.
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
    }

    [Guid("2650C2AB-3BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IFileChooser
    {
        [DispId(1)]
        string FileName { get; }

        [DispId(2)]
        object FileNames { get; }

        [DispId(3)]
        string FileNamesString { get; }

        [DispId(4)]
        string Filter { get; set; }

        [DispId(5)]
        int FilterIndex { get; set; }

        [DispId(6)]
        string Title { get; set; }

        [DispId(7)]
        bool Multiselect { get; set; }

        [DispId(8)]
        bool DereferenceLinks { get; set; }

        [DispId(9)]
        string DefaultExt { get; set; }
    }
}
