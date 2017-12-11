
// wrap the OpenFileDialog class for VBScript

using System.Windows.Forms; // for OpenFileDialog
using System.Runtime.InteropServices;
using System.Reflection;
using System.Linq; // for FileNames Cast

namespace VBScripting
{
    /// <summary> Provides a file chooser dialog for VBScript. </summary>
    [ProgId("VBScripting.FileChooser"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-3AF8-495F-AB4D-6C61BD463EA4")]
    public class FileChooser : IFileChooser
    {
        private OpenFileDialog chooser;

        /// <summary> Constructor </summary>
        public FileChooser()
        {
            this.chooser = new OpenFileDialog();
            this.Filter = "All files (*.*)|*.*";
            this.FilterIndex = 1;
            this.Title = "Browse for a file";
            this.DereferenceLinks = false;
            this.DefaultExt = "txt";
        }

        /// <summary> Opens a dialog enabling the user to browse for and choose a file. <para>
        /// <returns> Returns the filespec of the chosen file. Returns an empty string if the user cancels.  </returns> </para> </summary>
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

        /// <summary> Opens a dialog enabling the user to browse for and choose multiple files. <para> 
        /// <returns> Returns a string array of filespecs. Returns an empty array if the user cancels. </returns> </para>
        /// <para> Requires Multiselect to have been set to True. </para> </summary>
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

        /// <summary> Opens a dialog enabling the user to browse for and choose multiple files. <para> 
        /// <returns> Returns a delimited string of filespecs ( delimited by | ). Returns an empty string if the user cancels. </returns> </para>
        /// <para> Requires Multiselect to have been set to True. </para> </summary>
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

        /// <summary> Gets or sets the selectable file types. The default is "All files (*.*)|*.*" <para>
        /// <example> Example #1: <code> fc.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*" </code> </example> </para>
        /// <example> Example #2: <code> fc.Filter = "Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*" </code> </example> </summary>
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

        /// <summary> Gets or sets the index controlling which filter item is initially selected. 
        /// <para> An integer. The index is 1-based. The default is 1. </para> </summary>
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

        /// <summary> Gets or sets the dialog titlebar text. 
        /// <para> The default text is "Browse for a file." </para> </summary>
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

        /// <summary> Gets or sets whether multiple files can be selected. 
        /// <para> A boolean. The default is False. </para> </summary>
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

        /// <summary> Indicates whether the returned file is the referenced file or the .lnk file itself.
        /// <para> Gets or sets, if the selected file is a .lnk file, whether the filespec returned refers to the .lnk file itself (False) or to the file that the .lnk file points to (True). </para>
        /// <para> A boolean. The default is False. </para> </summary>
        /// <remarks> The default of the wrapped property, OpenFileDialog.DereferenceLinks, is True. </remarks>
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

        /// <summary> Gets or sets the file extension name that is automatically supplied when one is not specified.
        /// <para> A string. The default is "txt". </para> </summary>
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

    /// <summary> The COM interface for <see cref="FileChooser"/>. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-3BF8-495F-AB4D-6C61BD463EA4")]
    public interface IFileChooser
    {
        /// <summary> COM interface member for <see cref="FileChooser.FileName"/>. </summary>
        [DispId(1)]
        string FileName { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.FileNames"/>. </summary>
        [DispId(2)]
        object FileNames { get; }

        /// <summary> COM interface member for <see cref="FileChooser.FileNamesString"/>. </summary>
        [DispId(3)]
        string FileNamesString { get; }

        /// <summary> COM interface member for <see cref="FileChooser.Filter"/>. </summary>
        [DispId(4)]
        string Filter { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.FilterIndex"/>. </summary>
        [DispId(5)]
        int FilterIndex { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.Title"/>. </summary>
        [DispId(6)]
        string Title { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.Multiselect"/>. </summary>
        [DispId(7)]
        bool Multiselect { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.DereferenceLinks"/>. </summary>
        [DispId(8)]
        bool DereferenceLinks { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.DefaultExt"/>. </summary>
        [DispId(9)]
        string DefaultExt { get; set; }

        /// <summary> COM interface for <see cref="FileChooser.ValidateNames"/> </summary>
        [DispId(10)]
        bool ValidateNames { get; set; }

        /// <summary> COM interface member for <see cref="FileChooser.CheckFileExists"/> </summary>
        [DispId(11)]
        bool CheckFileExists { get; set; }
    }
}
