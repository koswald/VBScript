
// Progress bar for VBScript

using System.Windows.Forms; 
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing;
using System;

[assembly:AssemblyKeyFileAttribute("VBSProgressBar.snk")]

namespace VBSProgressBar
{
    [Guid("2650C2AB-4AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("VBScript.ProgressBar")]
    public class VBSProgressBar : IVBSProgressBar
    {
        private Form form;
        private ProgressBar pbar;
        private int pctX = -1; // percentage of avail. screen
        private int pctY = -1;

        // Constructor
        public VBSProgressBar()
        {
            this.form = new System.Windows.Forms.Form();
            this.pbar = new ProgressBar();
            this.SuspendLayout();
            this.form.Controls.Add(this.pbar);
            this.ResumeLayout(false);
            this.debug = false;

        }
        public bool Visible
        {
            get
            {
                return this.pbar.Visible && this.form.Visible;
            }
            set
            {
                this.form.Visible = value;
                this.pbar.Visible = value;
            }
        }
        public void PerformStep()
        {
            this.pbar.PerformStep();
        }
        public void FormSize(int width, int height)
        {
            this.form.ClientSize = new Size(width, height);
            // if positioning window by percentage,
            // reposition with new size
            if (this.pctX > -1 && this.pctY > -1)
            {
                this.FormLocationByPercentage(this.pctX, this.pctY);
            }
        }
        public void PBarSize(int width, int height)
        {
            this.pbar.Size = new Size(width, height);
        }
        public int Minimum
        {
            get { return this.pbar.Minimum; }
            set { this.pbar.Minimum = value; }
        }
        public int Maximum
        {
            get { return this.pbar.Maximum; }
            set { this.pbar.Maximum = value; }
        }
        public int Value
        {
            get { return this.pbar.Value; }
            set { this.pbar.Value = value; }
        }
        public int Step
        {
            get { return this.pbar.Step; }
            set { this.pbar.Step = value; }
        }
        public string Caption
        {
            get { return this.form.Text; }
            set { this.form.Text = value;
            }
        }
        public void FormLocation(int x, int y)
        {
            // mark as not positioning by percentage
            this.pctX = -1; this.pctY = -1;
            // enable manual positioning
            this.form.StartPosition = FormStartPosition.Manual;
            // position (locate) the window
            this.form.Location = new Point(x, y);
        }
        // set the window position by percentage of available screen width and height
        public void FormLocationByPercentage(int x, int y)
        {
            // save percentages; if the form is resized later, 
            // this method will be run again to readjust the location.
            this.pctX = x; this.pctY = y;
            // enable manual positioning
            this.form.StartPosition = FormStartPosition.Manual;
            // get available screen width and height
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            // convert percentages to pixels
            int pxX = (int)((workingArea.Width - this.form.Width) * x * .0101);
            int pxY = (int)((workingArea.Height - this.form.Height) * y * .0101);
            // position (locate) the window
            this.form.Location = new Point(pxX, pxY);
        }
        public void PBarLocation(int x, int y)
        {
            this.pbar.Location = new Point(x, y);
        }
        public void SuspendLayout()
        {
            this.form.SuspendLayout();
        }
        public void ResumeLayout(bool performLayout)
        {
            this.form.ResumeLayout(performLayout);
        }
        // set the form's icon by .ico file
        public void SetIconByIcoFile(string fileName)
        {
            try
            {
                this.form.Icon = new Icon(System.Environment.ExpandEnvironmentVariables(fileName));
            }
            catch (System.IO.FileNotFoundException fnfe)
            {
                if (debug)
                {
                    MessageBox.Show(fnfe.Message, "Couldn't find icon file",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                if (debug) { MessageBox.Show(e.ToString()); }
            }
        }
        // set the form's icon by extracting from a .dll or .exe file
        public void SetIconByDllFile(string fileName, int index)
        {
            try
            {
                this.form.Icon = IconExtractor.Extract(System.Environment.ExpandEnvironmentVariables(fileName), index, false);
            }
            catch (System.IO.FileNotFoundException fnfe)
            {
                if (debug)
                {
                    MessageBox.Show(fnfe.Message, "Couldn't find icon file",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                if (debug)
                {
                    MessageBox.Show(string.Format(
                        " Error setting icon from file \n" +
                        " {0} \n with index {1} \n\n {2}",
                        fileName, index, e.ToString()
                    ));
                }
            }
        }

        public bool debug { get; set; }

        // returns an object that can be used to provide
        // the integers for setting the FormBorderStyle
        public PBFormBorderStyle BorderStyle
        {
            get { return new PBFormBorderStyle(); }
            private set { }
        }

        // set the form border style according to one of the
        // styles defined in the FormBorderStyle class
        public void FormBorderStyle(int style)
        {
            if (style == this.BorderStyle.Fixed3D)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            }
            else if (style == this.BorderStyle.FixedDialog)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            }
            else if (style == this.BorderStyle.FixedSingle)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            }
            else if (style == this.BorderStyle.FixedToolWindow)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            }
            else if (style == this.BorderStyle.None)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            }
            else if (style == this.BorderStyle.Sizable)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            }
            else if (style == this.BorderStyle.SizableToolWindow)
            {
                this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            }
        }
        public void Dispose()
        {
            this.pbar.Dispose();
            this.form.Dispose();
        }
    }

    // enumeration of FormBorderStyle values for VBScript
    public enum formBorderStyle : int
    { 
        Fixed3D = 1, 
        FixedDialog, 
        FixedSingle, 
        FixedToolWindow, 
        None, 
        Sizable, 
        SizableToolWindow 
    }

    // returns integers corresponding to FormBorderStyle methods
    public class PBFormBorderStyle
    {
        public PBFormBorderStyle() { } // constructor
        public int Fixed3D 
        {
            get { return (int)formBorderStyle.Fixed3D; }
            private set { }
        }
        public int FixedDialog 
        {
            get { return (int)formBorderStyle.FixedDialog; }
            private set { }
        }
        public int FixedSingle 
        {
            get { return (int)formBorderStyle.FixedSingle; }
            private set { }
        }
        public int FixedToolWindow 
        {
            get { return (int)formBorderStyle.FixedToolWindow; }
            private set { }
        }
        public int None
        {
            get { return (int)formBorderStyle.None; }
            private set { }
        }
        public int Sizable
        {
            get { return (int)formBorderStyle.Sizable; }
            private set { }
        }
        public int SizableToolWindow 
        {
            get { return (int)formBorderStyle.SizableToolWindow; }
            private set { }
        }
    }

    // interface
    [Guid("2650C2AB-4BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IVBSProgressBar
    {
        [DispId(1)]
        bool Visible { get; set; }

        [DispId(2)]
        void PerformStep();

        [DispId(3)]
        void FormSize(int width, int height);

        [DispId(4)]
        void PBarSize(int width, int height);

        [DispId(5)]
        int Minimum { get; set; }

        [DispId(6)]
        int Maximum { get; set; }

        [DispId(7)]
        int Value { get; set; }

        [DispId(8)]
        int Step { get; set; }

        [DispId(10)]
        string Caption { get; set; }

        [DispId(11)]
        void FormLocation(int x, int y);

        [DispId(12)]
        void PBarLocation(int x, int y);

        [DispId(13)]
        void SuspendLayout();

        [DispId(14)]
        void ResumeLayout(bool performLayout);

        [DispId(15)]
        void FormLocationByPercentage(int x, int y);

        [DispId(16)]
        void SetIconByDllFile(string fileName, int index);

        [DispId(17)]
        bool debug { get; set; }
        
        [DispId(18)]
        PBFormBorderStyle BorderStyle { get; }

        [DispId(19)]
        void FormBorderStyle(int style);

        [DispId(20)]
        void SetIconByIcoFile(string fileName);

        [DispId(21)]
        void Dispose();
    }

    // Extract an icon from a .dll or .exe file
    // thanks to https://stackoverflow.com/users/98713/thomas-levesque
    // https://stackoverflow.com/questions/6872957/how-can-i-use-the-images-within-shell32-dll-in-my-c-sharp-project
    public class IconExtractor
    {

        public static Icon Extract(string file, int number, bool largeIcon)
        {
            IntPtr large;
            IntPtr small;
            ExtractIconEx(file, number, out large, out small, 1);
            try
            {
                return Icon.FromHandle(largeIcon ? large : small);
            }
            catch(Exception e)
            {
                throw(e);
                //return null;
            }

        }
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

    }

}
