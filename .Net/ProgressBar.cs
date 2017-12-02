
// Progress bar for VBScript

using System.Windows.Forms; 
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing;
using System;

namespace VBScripting
{
    /// <summary> Supplies a progress bar to VBScript, for illustration purposes. </summary>
    [ProgId("VBScripting.ProgressBar"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-4AF8-495F-AB4D-6C61BD463EA4")]
    public class ProgressBar : IProgressBar
    {
        private Form form;
        private System.Windows.Forms.ProgressBar pbar;
        private int pctX = -1; // percentage of avail. screen
        private int pctY = -1;

        /// <summary> Constructor. </summary>
        public ProgressBar()
        {
            this.form = new System.Windows.Forms.Form();
            this.pbar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            this.form.Controls.Add(this.pbar);
            this.ResumeLayout(false);
            this.Debug = false;
            this.Minimum = 0;
            this.Maximum = 100;

        }

        /// <summary> Gets or sets the progress bar's visibility. 
        /// <para> A boolean. The default is False. </para> </summary>
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

        /// <summary> Advances the progress bar one step.
        /// <para> See <see cref="ProgressBar.Step"/>. </para></summary>
        public void PerformStep()
        {
            this.pbar.PerformStep();
        }

        /// <summary> Sets the size of the window. </summary>
        /// <param name="width"> Width of the window in pixels. </param>
        /// <param name="height"> Height of the window in pixels. </param>
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

        /// <summary> Sets the size of the progress bar. </summary>
        /// <param name="width"> Width of the progress bar in pixels. </param>
        /// <param name="height"> Height of the progress bar in pixels. </param>
        public void PBarSize(int width, int height)
        {
            this.pbar.Size = new Size(width, height);
        }

        /// <summary> Gets or sets the value (<see cref="ProgressBar.Value"/>) at which there is no apparent progress.
        /// <para> An integer. The default is 0. </para> </summary>
        public int Minimum
        {
            get { return this.pbar.Minimum; }
            set { this.pbar.Minimum = value; }
        }

        /// <summary> Gets or sets the value (<see cref="ProgressBar.Value"/>) at which the progress appears to be complete. 
        /// <para> An integer. The default is 100. </para></summary>
        public int Maximum
        {
            get { return this.pbar.Maximum; }
            set { this.pbar.Maximum = value; }
        }

        /// <summary> Gets or sets the apparent progress.
        /// <para> Should be at or above the minimum (<see cref="ProgressBar.Minimum"/>) and at or below the maximum (<see cref="ProgressBar.Maximum"/>). </para>
        /// An integer. </summary>
        public int Value
        {
            get { return this.pbar.Value; }
            set { this.pbar.Value = value; }
        }

        /// <summary> Gets or sets the increment between steps.
        /// <para> See <see cref="ProgressBar.PerformStep()"/> </para> </summary>
        public int Step
        {
            get { return this.pbar.Step; }
            set { this.pbar.Step = value; }
        }

        /// <summary> Gets or sets the window title-bar text. </summary>
        public string Caption
        {
            get { return this.form.Text; }
            set { this.form.Text = value;
            }
        }

        /// <summary> Sets the position of the window. </summary>
        /// <param name="x"> The distance of the window in pixels from the left edge of the screen. </param>
        /// <param name="y"> The distance of the window in pixels from the top edge of the screen. </param>
        public void FormLocation(int x, int y)
        {
            // mark as not positioning by percentage
            this.pctX = -1; this.pctY = -1;
            // enable manual positioning
            this.form.StartPosition = FormStartPosition.Manual;
            // position (locate) the window
            this.form.Location = new Point(x, y);
        }


        /// <summary> Sets the position of the window. </summary>
        /// <param name="x"> The horizontal position of the window from 0 (at the left) to 100 (at the right). </param>
        /// <param name="y"> The vertical position of the window from 0 (at the top) to 100 (at the bottom). </param>
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

        /// <summary> Sets the location of the progress bar. </summary>
        /// <param name="x"> Distance in pixels of the left edge of the progress bar from the left edge of the window. </param>
        /// <param name="y"> Distance in pixels of the top edge of the progress bar from the top edge of the window. </param>
        public void PBarLocation(int x, int y)
        {
            this.pbar.Location = new Point(x, y);
        }

        /// <summary> Suspend drawing of the window temporarily. </summary>
        public void SuspendLayout()
        {
            this.form.SuspendLayout();
        }

        /// <summary> Resume drawing the window. </summary>
        public void ResumeLayout(bool performLayout)
        {
            this.form.ResumeLayout(performLayout);
        }
        /// <summary> Sets the system tray <see cref="Icon"/> given an .ico file. </summary>
        /// <param name="fileName"> The filespec of the .ico file. Environment variables are allowed. </param>
        public void SetIconByIcoFile(string fileName)
        {
            try
            {
                this.form.Icon = new Icon(System.Environment.ExpandEnvironmentVariables(fileName));
            }
            catch (System.IO.FileNotFoundException fnfe)
            {
                if (Debug)
                {
                    MessageBox.Show(fnfe.Message, "Couldn't find icon file",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                if (Debug) { MessageBox.Show(e.ToString()); }
            }
        }
        /// <summary> Sets the system tray <see cref="Icon"/> from a .dll or .exe file. </summary>
        /// <param name="fileName"> The path and name of a .dll or .exe file that contains icons. </param>
        /// <param name="index"> The index of the icon. An integer. </param>
        public void SetIconByDllFile(string fileName, int index)
        {
            try
            {
                this.form.Icon = IconExtractor.Extract(System.Environment.ExpandEnvironmentVariables(fileName), index, false);
            }
            catch (System.IO.FileNotFoundException fnfe)
            {
                if (Debug)
                {
                    MessageBox.Show(fnfe.Message, "Couldn't find icon file",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                if (Debug)
                {
                    MessageBox.Show(string.Format(
                        " Error setting icon from file \n" +
                        " {0} \n with index {1} \n\n {2}",
                        fileName, index, e.ToString()
                    ));
                }
            }
        }

        /// <summary> Gets or sets whether the type is under development. 
        /// <para> Affects the behavior of two methods, if exceptions are thrown. See <see cref="SetIconByIcoFile(string)"/> and <see cref="SetIconByDllFile(string, int)"/> </para> </summary>
        public bool Debug { get; set; }

        /// <summary> Provides an object useful in VBScript for setting <see cref="FormBorderStyle(int)"/>. </summary>
        public FormBorderStyleT BorderStyle
        {
            get { return new FormBorderStyleT(); }
            private set { }
        }

        /// <summary> Sets the style of the window border. </summary>
        /// <param name="style"> An integer. Return values of one of the seven methods (Fixed3D, FixedDialog, FixedSingle, FixedToolWindow, None, Sizable (default), and SizableToolWindow) of <see cref="BorderStyle"/> can be used. </param>
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

        /// <summary> Disposes of the object's resources. </summary>
        public void Dispose()
        {
            this.pbar.Dispose();
            this.form.Dispose();
        }
    }

    /// <summary> C# enum not intended for use by VBScript. 
    /// <para> Corresponds to but not equivalent to System.Windows.Forms.FormBorderStyle. </para> </summary>
    [Guid("2650C2AB-4CF8-495F-AB4D-6C61BD463EA4")]
    public enum FormBorderStyle : int
    {
        /// <summary> Return value can be cast to an int: 1 </summary>
        Fixed3D = 1,
        /// <summary> Return value can be cast to an int: 2 </summary>
        FixedDialog,
        /// <summary> Return value can be cast to an int: 3 </summary>
        FixedSingle,
        /// <summary> Return value can be cast to an int: 4 </summary>
        FixedToolWindow,
        /// <summary> Return value can be cast to an int: 5 </summary>
        None,
        /// <summary> Return value can be cast to an int: 6 </summary>
        Sizable,
        /// <summary> Return value can be cast to an int: 7 </summary>
        SizableToolWindow
    }

    /// <summary> Supplies the type required by <see cref="BorderStyle"/>
    /// <para> Not intended for use in VBScript. </para> </summary>
    [Guid("2650C2AB-4DF8-495F-AB4D-6C61BD463EA4")]
    public class FormBorderStyleT
    {
        /// <returns> Returns 1 </returns>
        public int Fixed3D 
        {
            get { return (int)FormBorderStyle.Fixed3D; }
            private set { }
        }
        /// <returns> Returns 2 </returns>
        public int FixedDialog 
        {
            get { return (int)FormBorderStyle.FixedDialog; }
            private set { }
        }
        /// <returns> Returns 3 </returns>
        public int FixedSingle 
        {
            get { return (int)FormBorderStyle.FixedSingle; }
            private set { }
        }
        /// <returns> Returns 4 </returns>
        public int FixedToolWindow 
        {
            get { return (int)FormBorderStyle.FixedToolWindow; }
            private set { }
        }
        /// <returns> Returns 5 </returns>
        public int None
        {
            get { return (int)FormBorderStyle.None; }
            private set { }
        }
        /// <returns> Returns 6 </returns>
        public int Sizable
        {
            get { return (int)FormBorderStyle.Sizable; }
            private set { }
        }
        /// <returns> Returns 7 </returns>
        public int SizableToolWindow 
        {
            get { return (int)FormBorderStyle.SizableToolWindow; }
            private set { }
        }
    }

    /// <summary> The COM interface for <see cref="ProgressBar"/>. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-4BF8-495F-AB4D-6C61BD463EA4")]
    public interface IProgressBar
    {
        /// <summary> COM interface member for <see cref="Visible"/></summary>
        [DispId(1)]
        bool Visible { get; set; }

        /// <summary> COM interface member for <see cref="PerformStep()"/></summary>
        [DispId(2)]
        void PerformStep();

        /// <summary> COM interface member for <see cref="FormSize(int, int)"/></summary>
        [DispId(3)]
        void FormSize(int width, int height);

        /// <summary> COM interface member for <see cref="PBarLocation(int, int)"/></summary>
        [DispId(4)]
        void PBarSize(int width, int height);

        /// <summary> COM interface member for <see cref="Minimum"/></summary>
        [DispId(5)]
        int Minimum { get; set; }

        /// <summary> COM interface member for <see cref="Maximum"/></summary>
        [DispId(6)]
        int Maximum { get; set; }

        /// <summary> COM interface member for <see cref="Value"/></summary>
        [DispId(7)]
        int Value { get; set; }

        /// <summary> COM interface member for <see cref="Step"/></summary>
        [DispId(8)]
        int Step { get; set; }

        /// <summary> COM interface member for <see cref="Caption"/></summary>
        [DispId(10)]
        string Caption { get; set; }

        /// <summary> COM interface member for <see cref="FormLocation(int, int)"/></summary>
        [DispId(11)]
        void FormLocation(int x, int y);

        /// <summary> COM interface member for <see cref="PBarLocation(int, int)"/></summary>
        [DispId(12)]
        void PBarLocation(int x, int y);

        /// <summary> COM interface member for <see cref="SuspendLayout()"/></summary>
        [DispId(13)]
        void SuspendLayout();

        /// <summary> COM interface member for <see cref="ResumeLayout(bool)"/></summary>
        [DispId(14)]
        void ResumeLayout(bool performLayout);

        /// <summary> COM interface member for <see cref="FormLocationByPercentage(int, int)"/></summary>
        [DispId(15)]
        void FormLocationByPercentage(int x, int y);

        /// <summary> COM interface member for <see cref="SetIconByDllFile(string, int)"/></summary>
        [DispId(16)]
        void SetIconByDllFile(string fileName, int index);

        /// <summary> COM interface member for <see cref="Debug"/></summary>
        [DispId(17)]
        bool Debug { get; set; }

        /// <summary> COM interface member for <see cref="BorderStyle"/></summary>
        [DispId(18)]
        FormBorderStyleT BorderStyle { get; }

        /// <summary> COM interface member for <see cref="FormBorderStyle(int)"/></summary>
        [DispId(19)]
        void FormBorderStyle(int style);

        /// <summary> COM interface member for <see cref="SetIconByIcoFile(string)"/></summary>
        [DispId(20)]
        void SetIconByIcoFile(string fileName);

        /// <summary> COM interface member for <see cref="Dispose()"/></summary>
        [DispId(21)]
        void Dispose();
    }
}
