
// System tray icon for VBScript


// NotifyIcon class example
// https://msdn.microsoft.com/en-us/library/system.windows.forms.notifyicon.icon(v=vs.110).aspx

// See also
// https://stackoverflow.com/questions/7625421/minimize-app-to-system-tray

using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing;
using System;

[assembly: AssemblyKeyFileAttribute("VBSNotifyIcon.snk")]

namespace VBSNotifyIcon
{
    [Guid("2650C2AB-5AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("VBSNotifyIcon")]
    public class VBSNotifyIcon : IVBSNotifyIcon
    {
        private NotifyIcon notifyIcon;
        private ContextMenu contextMenu;
        private MenuItem menuItem1;

        // Constructor
        public VBSNotifyIcon()
        {
            this.contextMenu = new ContextMenu();
            this.menuItem1 = new MenuItem();

            // Initialize contextMenu
            this.contextMenu.MenuItems.AddRange(
                new MenuItem[] { this.menuItem1 });

            // Initialize menuItem1
            this.menuItem1.Index = 0;
            this.menuItem1.Text = "E&xit";
            this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);

            // Create the NotifyIcon
            this.notifyIcon = new NotifyIcon();

            this.notifyIcon.ContextMenu = this.contextMenu;

            this.notifyIcon.Click += new System.EventHandler(this.notifyIcon_Click);

            this.debug = false;
            this.BalloonTipLifetime = 5000;
        }

        public bool debug { get; set; }

        public string Text
        {
            get { return this.notifyIcon.Text; }

            set { this.notifyIcon.Text = value; }
        }

        public bool Visible
        {
            get { return this.notifyIcon.Visible; }
            set { this.notifyIcon.Visible = value; }
        }

        // set the form's icon by .ico file
        public void SetIconByIcoFile(string fileName)
        {
            try
            {
                this.notifyIcon.Icon = new Icon(System.Environment.ExpandEnvironmentVariables(fileName));
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
                this.notifyIcon.Icon = IconExtractor.Extract(System.Environment.ExpandEnvironmentVariables(fileName), index, false);
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

        public string BalloonTipTitle
        {
            get { return this.notifyIcon.BalloonTipTitle; }
            set { this.notifyIcon.BalloonTipTitle = value; }
        }

        public string BalloonTipText
        {
            get { return this.notifyIcon.BalloonTipText; }
            set { this.notifyIcon.BalloonTipText = value; }
        }

        public int BalloonTipLifetime { get; set; } // "Deprecated as of Windows Vista." Now the value is overridden by accessibility settings

        public ToolTipIconE ToolTipIcon
        {
            get { return new ToolTipIconE(); }
            private set { }
        }

        public void SetBalloonTipIcon(int type)
        {
            if (type == this.ToolTipIcon.Error)
            {
                this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Error;
            }
            else if (type == this.ToolTipIcon.Info)
            {
                this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            }
            else if (type == this.ToolTipIcon.None)
            {
                this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.None;
            }
            else if (type == this.ToolTipIcon.Warning)
            {
                this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning;
            }
        }

        public void Dispose()
        {
            this.notifyIcon.Dispose();
        }

        private void notifyIcon_Click(object Sender, EventArgs e)
        {
            this.notifyIcon.ShowBalloonTip(this.BalloonTipLifetime);
        }

        private void menuItem1_Click(object Sender, EventArgs e)
        {
            this.Dispose();
        }
    }

    // interface
    [Guid("2650C2AB-5BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IVBSNotifyIcon
    {
        [DispId(1)]
        string Text { get; set; }

        [DispId(2)]
        void Dispose();

        [DispId(3)]
        bool Visible { get; set; }

        [DispId(4)]
        bool debug { get; set; }

        [DispId(5)]
        void SetIconByIcoFile(string file);

        [DispId(6)]
        void SetIconByDllFile(string file, int index);

        [DispId(7)]
        string BalloonTipTitle { get; set; }

        [DispId(8)]
        string BalloonTipText { get; set; }

        [DispId(9)]
        int BalloonTipLifetime { get; set; } // milliseconds

        [DispId(10)]
        ToolTipIconE ToolTipIcon { get; }

        [DispId(11)]
        void SetBalloonTipIcon(int type);
    }

    // enum for VBScript 
    // corresponds to the System.Windows.Forms.ToolTipIcon enum
    public enum toolTipIcon : int
    {
        Error = 1,
        Info,
        None,
        Warning
    }

    // For returning an object to VBScript with ToolTipIcon types
    public class ToolTipIconE
    {
        public int Error
        {
            get { return (int)toolTipIcon.Error; }
            private set { }
        }
        public int Info
        {
            get { return (int)toolTipIcon.Info; }
            private set { }
        }
        public int None
        {
            get { return (int)toolTipIcon.None; }
            private set { }
        }
        public int Warning
        {
            get { return (int)toolTipIcon.Warning; }
            private set { }
        }
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
            }

        }
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

    }
}
