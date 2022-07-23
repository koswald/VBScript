// Progress bar for VBScript

using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing;
using System;

namespace VBScripting
{
    /// <summary> Supplies a progress bar to VBScript, for illustration purposes. </summary>
    [ProgId( "VBScripting.ProgressBar" ),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-4AF8-495F-AB4D-6C61BD463EA4")]
    public class ProgressBar : IProgressBar
    {
        private Form form;
        private System.Windows.Forms.ProgressBar pbar;
        private int _style = 0;
        private int pctX = -1; // percentage of avail. screen
        private int pctY = -1;

        /// Constructor <summary></summary>
        public ProgressBar()
        {
            this.form = new System.Windows.Forms.Form();
            this.pbar = new System.Windows.Forms.ProgressBar();
            this.form.Controls.Add(this.pbar);
            this.Debug = false;
            Application.EnableVisualStyles();
        }

        /// <summary> Gets or sets the progress bar's visibility.  </summary>
        /// <remarks> A boolean. The default is False. </remarks>
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

        /// <summary> Advances the progress bar one step.</summary>
        public void PerformStep()
        {
            this.pbar.PerformStep();
        }

        /// <summary> Sets the size of the window. </summary>
        /// <parameters> width, height </parameters>
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
        /// <parameters> width, height </parameters>
        public void PBarSize(int width, int height)
        {
            this.pbar.Size = new Size(width, height);
        }

        /// <summary> Gets or sets the value at which there is no apparent progress. </summary>
        /// <remarks> An integer. The default is 0. </remarks>
        public int Minimum
        {
            get { return this.pbar.Minimum; }
            set { this.pbar.Minimum = value; }
        }

        /// <summary> Gets or sets the value at which the progress appears to be complete. </summary>
        /// <remarks> An integer. The default is 100. </remarks>
        public int Maximum
        {
            get { return this.pbar.Maximum; }
            set { this.pbar.Maximum = value; }
        }

        /// <summary> Gets or sets the apparent progress. </summary>
        /// <remarks> An integer. Should be at or above the minimum and at or below the maximum. </remarks>
        public int Value
        {
            get { return this.pbar.Value; }
            set
            {
                if (value > pbar.Maximum)
                {
                    this.pbar.Value = pbar.Maximum;
                }
                else if (value < pbar.Minimum)
                {
                    this.pbar.Value = pbar.Minimum;
                }
                else
                {
                    this.pbar.Value = value;
                }
            }
        }

        /// <summary> Integer. Gets or sets the increment between steps. </summary>
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

        /// <summary> Sets the position of the window, in pixels. </summary>
        /// <parameters> x, y </parameters>
        public void FormLocation(int x, int y)
        {
            // mark as not positioning by percentage
            this.pctX = -1; this.pctY = -1;
            // enable manual positioning
            this.form.StartPosition = FormStartPosition.Manual;
            // position (locate) the window
            this.form.Location = new Point(x, y);
        }


        /// <summary> Sets the position of the window, as a percentage of screen width and height. </summary>
        /// <parameters> x, y </parameters>
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

        /// <summary> Sets the location of the progress bar within the window. </summary>
        /// <parameters> x, y </parameters>
        public void PBarLocation(int x, int y)
        {
            this.pbar.Location = new Point(x, y);
        }

        /// <summary> Suspends drawing of the window temporarily. </summary>
        public void SuspendLayout()
        {
            this.form.SuspendLayout();
        }

        /// <summary> Resumes drawing the window. </summary>
        public void ResumeLayout(bool performLayout)
        {
            this.form.ResumeLayout(performLayout);
        }
        /// <summary> Sets the icon given the filespec of an .ico file. </summary>
        /// <parameters> fileName </parameters>
        /// <remarks> Environment variables are allowed. </remarks>
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
        /// <summary> Sets the icon given the filespec of a .dll or .exe file and an index. </summary>
        /// <parameters> fileName, index </parameters>
        /// <remarks> The index is an integer that identifies the icon. Environment variables are allowed. </remarks>
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
                else throw;
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
                else throw;
            }
        }

        /// <summary> Gets or sets whether the type is under development.  </summary>
        /// <remarks> Affects the behavior of two methods, SetIconByIcoFile and SetIconByDllFile, if exceptions are thrown: when debugging, a message box is shown. Default is False. </remarks>
        public bool Debug { get; set; }

        /// <summary> Provides an object useful in VBScript for setting FormBorderStyle. </summary>
        /// <returns> a FormBorderStyleT </returns>
        public FormBorderStyleT BorderStyle
        {
            get { return new FormBorderStyleT(); }
            private set { }
        }

        /// <summary> Sets the style of the progress bar.  </summary>
        /// <remarks> Use 1 for continuous, and 2 for marquee. </remarks>
        public int Style
        {
            set
            {
                _style = value;
                this.pbar.Style = (ProgressBarStyle) value;
            }
            get { return _style; }
        }

        /// <summary> Sets the style of the window border. </summary>
        /// <remarks> An integer. One of the BorderStyle property return values can be used: Fixed3D, FixedDialog, FixedSingle, FixedToolWindow, None, Sizable (default), or SizableToolWindow. VBScript example: <pre> pb.FormBorderStyle = pb.BorderStyle.Fixed3D </pre></remarks>
        public int FormBorderStyle
        {
            set
            {
                if (value == this.BorderStyle.Fixed3D)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
                }
                else if (value == this.BorderStyle.FixedDialog)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                }
                else if (value == this.BorderStyle.FixedSingle)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
                }
                else if (value == this.BorderStyle.FixedToolWindow)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
                }
                else if (value == this.BorderStyle.None)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                }
                else if (value == this.BorderStyle.Sizable)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                }
                else if (value == this.BorderStyle.SizableToolWindow)
                {
                    this.form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
                }
            }
        }

        /// <summary> Disposes of the object's resources. </summary>
        public void Dispose()
        {
            this.pbar.Dispose();
            this.form.Dispose();
        }

    }

    /// <summary> Enumeration of border styles. </summary>
    /// <remarks> This class is available to VBScript via the <code>ProgressBar.BorderStyle</code> property. </remarks>
    [Guid("2650C2AB-4DF8-495F-AB4D-6C61BD463EA4")]
    public class FormBorderStyleT
    {
        /// <returns> 1 </returns>
        public int Fixed3D
        {
            get { return 1; }
            private set { }
        }
        /// <returns> 2 </returns>
        public int FixedDialog
        {
            get { return 2; }
            private set { }
        }
        /// <returns> 3 </returns>
        public int FixedSingle
        {
            get { return 3; }
            private set { }
        }
        /// <returns> 4 </returns>
        public int FixedToolWindow
        {
            get { return 4; }
            private set { }
        }
        /// <returns> 5 </returns>
        public int None
        {
            get { return 5; }
            private set { }
        }
        /// <returns> 6 </returns>
        public int Sizable
        {
            get { return 6; }
            private set { }
        }
        /// <returns> 7 </returns>
        public int SizableToolWindow
        {
            get { return 7; }
            private set { }
        }
    }

    /// <summary> Exposes the VBScripting.ProgressBar members to COM/VBScript. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-4BF8-495F-AB4D-6C61BD463EA4")]
    public interface IProgressBar
    {
        /// <summary> </summary>
        [DispId(1)]
        bool Visible { get; set; }

        /// <summary> </summary>
        [DispId(2)]
        void PerformStep();

        /// <summary> </summary>
        [DispId(3)]
        void FormSize(int width, int height);

        /// <summary> </summary>
        [DispId(4)]
        void PBarSize(int width, int height);

        /// <summary> </summary>
        [DispId(5)]
        int Minimum { get; set; }

        /// <summary> </summary>
        [DispId(6)]
        int Maximum { get; set; }

        /// <summary> </summary>
        [DispId(7)]
        int Value { get; set; }

        /// <summary> </summary>
        [DispId(8)]
        int Step { get; set; }

        /// <summary> </summary>
        [DispId(10)]
        string Caption { get; set; }

        /// <summary> </summary>
        [DispId(11)]
        void FormLocation(int x, int y);

        /// <summary> </summary>
        [DispId(15)]
        void FormLocationByPercentage(int x, int y);

        /// <summary> </summary>
        [DispId(12)]
        void PBarLocation(int x, int y);

        /// <summary> </summary>
        [DispId(13)]
        void SuspendLayout();

        /// <summary> </summary>
        [DispId(14)]
        void ResumeLayout(bool performLayout);

        /// <summary> </summary>
        [DispId(20)]
        void SetIconByIcoFile(string fileName);

        /// <summary> </summary>
        [DispId(16)]
        void SetIconByDllFile(string fileName, int index);

        /// <summary> </summary>
        [DispId(17)]
        bool Debug { get; set; }

        /// <summary> </summary>
        [DispId(18)]
        FormBorderStyleT BorderStyle { get; }

        /// <summary> </summary>
        [DispId(19)]
        int FormBorderStyle { set; }

        /// <summary> </summary>
        [DispId(21)]
        void Dispose();

        /// <summary> </summary>
        [DispId(22)]
        int Style { get; set; }
    }

}
