// System tray icon for VBScript

using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Collections.Generic;
using System;
using System.Reflection;

namespace VBScripting
{
    /// <summary> Provides a notification area (system tray) icon for VBScript, for illustration purposes. </summary>
    [ProgId( "VBScripting.NotifyIcon" ),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-5AF8-495F-AB4D-6C61BD463EA4")]
    public class NotifyIcon : INotifyIcon
    {
        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.ContextMenu contextMenu;
        private CallbackEventSettings settings;
        private int nextIndex;

        /// <summary> Constructor </summary>
        public NotifyIcon()
        {
            this.settings = new CallbackEventSettings();
            this.nextIndex = 0;
            this.contextMenu = new ContextMenu();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon();
            this.notifyIcon.MouseUp += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseUp);
            this.notifyIcon.BalloonTipClicked += new EventHandler(this.notifyIcon_BalloonTipClicked);
            this.notifyIcon.ContextMenu = this.contextMenu;
        }

        /// <summary> Gets or sets the text shown when the mouse hovers over the system tray icon. </summary>
        public string Text
        {
            get { return this.notifyIcon.Text; }
            set { this.notifyIcon.Text = value; }
        }

        /// <summary>  Disposes of the icon resources when it is no longer needed. </summary>
        /// <remarks> If this method is not called, the icon may persist in the system tray until the mouse hovers over it, even after the object instance has lost scope. </remarks>
        public void Dispose()
        {
            this.notifyIcon.Icon.Dispose();
            this.notifyIcon.Dispose();
        }

        /// <summary> Gets or sets the icon's visibility. A boolean. </summary>
        /// <remarks> Required. Set this property to True after initializing other settings. </remarks>
        public bool Visible
        {
            get { return this.notifyIcon.Visible; }
            set { this.notifyIcon.Visible = value; }
        }

        /// <summary> Sets the system tray icon given an .ico file. </summary>
        /// <parameters> fileName </parameters>
        /// <remarks> The parameter <code>fileName</code> specifies the filespec of the .ico file. Environment variables and relative paths are allowed. </remarks>
        public void SetIconByIcoFile(string fileName)
        {
            if (this.notifyIcon.Icon != null)
                this.notifyIcon.Icon.Dispose();
            try
            {
                this.notifyIcon.Icon = new Icon(Environment.ExpandEnvironmentVariables(fileName));
            }
            catch (Exception e)
            {
                Admin.Log(string.Format(
                    "Exception at VBScripting.NotifyIcon.SetIconbyIcoFile\nFile: {0}\n\n{1}",
                    fileName, e.ToString()
                ));
                throw new Exception(string.Format(
                    "Failed to set icon from file {0}", fileName), e);
            }
        }

        /// <summary> Sets the system tray icon from a .dll or .exe file. </summary>
        /// <parameters> fileName, index, largeIcon </parameters>
        /// <remarks> Parameters: <code>fileName</code> is the path and name of a .dll or .exe file that contains icons. <code>index</code> is an integer that specifies which icon to use. <code>largeIcon</code> is a boolean that specifies whether to use a large or small icon. </remarks>
        public void SetIconByDllFile(string fileName, int index, bool largeIcon)
        {
            if (this.notifyIcon.Icon != null)
                this.notifyIcon.Icon.Dispose();
            try
            {
                this.notifyIcon.Icon = VBScripting.IconExtractor.Extract(System.Environment.ExpandEnvironmentVariables(fileName), index, largeIcon);
            }
            catch (Exception e)
            {
                Admin.Log(string.Format(
                    "Exception at VBScripting.NotifyIcon.SetIconByDllFile\nFile: {0}\nIndex: {1}\n\n{2}",
                    fileName, index, e.ToString()
                ));
                throw new Exception(string.Format(
                    "Failed to set icon from file {0}; index: {1}; largeIcon: {2}.", fileName, index, largeIcon), e);
            }
        }

        /// <summary> Gets or sets the title of the "balloon tip" or notification. </summary>
        public string BalloonTipTitle
        {
            get { return this.notifyIcon.BalloonTipTitle; }
            set { this.notifyIcon.BalloonTipTitle = value; }
        }

        /// <summary> Gets or sets the text of the "balloon tip" or notification. </summary>
        public string BalloonTipText
        {
            get { return this.notifyIcon.BalloonTipText; }
            set { this.notifyIcon.BalloonTipText = value; }
        }

        /// <summary> Gets or sets the lifetime of the "balloon tip" or notification. An integer (milliseconds). Deprecated as of Windows Vista. The value is overridden by accessibility settings.  </summary>
        public int BalloonTipLifetime { get; set; }

        /// <summary> Gets an object useful in VBScript for selecting a ToolTipIcon type. The properties Error, Info, None, and Warning may be used with SetBalloonTipIcon. </summary>
        /// <returns> a ToolTipIconT </returns>
        /// <remarks> VBScript example: <pre> obj.SetBallonTipIcon obj.ToolTipIcon.Warning </pre></remarks>
        public ToolTipIconT ToolTipIcon
        {
            get { return new ToolTipIconT(); }
            private set { }
        }

        /// <summary> Sets the icon of the "balloon tip" or notification. </summary>
        /// <parameters> type </parameters>
        /// <remarks> The parameter <code>type</code> is an integer that specifies which icon to use: Return values of ToolTipIcon properties can be used: Error = 1, Info = 2, None = 3, Warning = 4. </remarks>
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

        /// <summary> Show the balloon tip. </summary>
        public void ShowBalloonTip()
        {
            this.notifyIcon.ShowBalloonTip(this.BalloonTipLifetime);
        }

        /// <summary> Add a menu item to the system tray icon's context menu. </summary>
        /// <parameters> menuText, callbackRef </parameters>
        /// <remarks> This method can be called only from VBScript. The parameter <code>menuText</code> is a string that specifies the text that appears in the menu. The parameter <code>callbackRef</code> is a VBScript object reference returned by the VBScript GetRef Function. </remarks>
        public void AddMenuItem(string menuText, object callbackRef)
        {
            this.settings.AddRef(new CallbackReference(this.nextIndex, callbackRef));

            MenuItem menuItem = new MenuItem();
            menuItem.Text = menuText;
            menuItem.Index = this.nextIndex;
            menuItem.Click += new System.EventHandler(this.CallbackEventHandler);
            this.contextMenu.MenuItems.Add(menuItem);
            this.nextIndex += 1;
        }

        // Invoke a callback reference from a List according to the menu item index
        private void CallbackEventHandler(object Sender, EventArgs e)
        {
            this.InvokeCallbackByIndex(((MenuItem)Sender).Index);
        }

        /// <summary> Provide callback testability from VBScript. </summary>
        public void InvokeCallbackByIndex(int index)
        {
            foreach (var reference in this.settings.Refs)
            {
                if (reference.Index == index)
                {
                    ComEvent.InvokeComCallback(reference.Reference);
                }
            }
        }

        /// <summary> Show the context menu. </summary>
        /// <remarks> Public in order to provide testability from VBScript. </remarks>
        public void ShowContextMenu()
        {
            MethodInfo mi = typeof(System.Windows.Forms.NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
            mi.Invoke(notifyIcon, null);
        }

        // show the context menu on left mouse click too
        private void notifyIcon_MouseUp(object Sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.ShowContextMenu();
            }
        }

        private void notifyIcon_BalloonTipClicked(object Sender, EventArgs e)
        {
            if (balloonTipCallback != null)
            {
                ComEvent.InvokeComCallback(balloonTipCallback);
            }
        }

        private object balloonTipCallback;

        /// <summary> Sets the VBScript callback Sub or Function reference invoked when the notification "balloon" is clicked. </summary>
        /// <remarks> VBScript example: <pre>    obj.SetBalloonTipCallback GetRef( "ProcedureName" ) </pre></remarks>
        public void SetBalloonTipCallback(object callbackRef)
        {
            this.balloonTipCallback = callbackRef;
        }

        /// <summary> Disables a menu item. </summary>
        /// <parameters> index </parameters>
        /// <remarks> Parameter <em> index </em> values are integers assigned automatically as each item is added to the menu, beginning with 0.</remarks>
        public void DisableMenuItem(int index)
        {
            this.contextMenu.MenuItems[index].Enabled = false;
        }

        /// <summary> Enables a menu item. </summary>
        /// <parameters> index </parameters>
        /// <remarks> Parameter <em> index </em> values are integers assigned automatically as each item is added to the menu, beginning with 0.</remarks>
        public void EnableMenuItem(int index)
        {
            this.contextMenu.MenuItems[index].Enabled = true;
        }
    }

    /// <summary> The COM interface for VBScripting.NotifyIcon </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-5BF8-495F-AB4D-6C61BD463EA4")]
    public interface INotifyIcon
    {
        /// <summary> </summary>
        [DispId(1)]
        string Text { get; set; }

        /// <summary> </summary>
        [DispId(2)]
        void Dispose();

        /// <summary> </summary>
        [DispId(3)]
        bool Visible { get; set; }

        /// <summary> </summary>
        [DispId(5)]
        void SetIconByIcoFile(string file);

        /// <summary> </summary>
        [DispId(6)]
        void SetIconByDllFile(string file, int index, bool largeIcon);

        /// <summary> </summary>
        [DispId(7)]
        string BalloonTipTitle { get; set; }

        /// <summary> </summary>
        [DispId(8)]
        string BalloonTipText { get; set; }

        /// <summary> </summary>
        [DispId(9)]
        int BalloonTipLifetime { get; set; } // milliseconds

        /// <summary> </summary>
        [DispId(10)]
        ToolTipIconT ToolTipIcon { get; }

        /// <summary> </summary>
        [DispId(11)]
        void SetBalloonTipIcon(int type);

        /// <summary> </summary>
        [DispId(12)]
        void ShowBalloonTip();

        /// <summary> </summary>
        [DispId(13)]
        void AddMenuItem(string menuText, object callbackRef);

        /// <summary> </summary>
        [DispId(14)]
        void InvokeCallbackByIndex(int index);

        /// <summary> </summary>
        [DispId(15)]
        void ShowContextMenu();

        /// <summary> </summary>
        [DispId(16)]
        void SetBalloonTipCallback(object callbackRef);

        /// <summary> </summary>
        [DispId(17)]
        void DisableMenuItem(int index);

        /// <summary> </summary>
        [DispId(18)]
        void EnableMenuItem(int index);
    }

    /// <summary> Supplies the type required by NotifyIcon.ToolTipIcon </summary>
    /// <remarks><strong> This class is accessible from VBScript via the <code>NotifyIcon.ToolTipIcon</code> property. </strong></remarks>
    [Guid("2650C2AB-5DF8-495F-AB4D-6C61BD463EA4")]
    public class ToolTipIconT
    {
        /// <returns> 1 </returns>
        public int Error
        {
            get { return (int)ToolTipIcon.Error; }
            private set { }
        }
        /// <returns> 2 </returns>
        public int Info
        {
            get { return (int)ToolTipIcon.Info; }
            private set { }
        }
        /// <returns> 3 </returns>
        public int None
        {
            get { return (int)ToolTipIcon.None; }
            private set { }
        }
        /// <returns> 4 </returns>
        public int Warning
        {
            get { return (int)ToolTipIcon.Warning; }
            private set { }
        }
    }

    /// <summary> Settings for saving VBScript method references. This class is not accessible from VBScript. </summary>
    [Guid("2650C2AB-5EF8-495F-AB4D-6C61BD463EA4")]
    public class CallbackEventSettings
    {
        /// <summary> Gets or sets a list of callback references. </summary>
        public List<CallbackReference> Refs { get; set; }

        /// <summary> Constructor </summary>
        public CallbackEventSettings()
        {
            this.Refs = new List<CallbackReference>();
        }

        /// <summary> Adds a CallbackReference instance reference to the list. </summary>
        /// <parameters> callbackRef </parameters>
        public void AddRef(CallbackReference callbackRef)
        {
            if (callbackRef != null && !(this.Refs.Contains(callbackRef)))
            {
                this.Refs.Add(callbackRef);
            }
        }
    }
    /// <summary> An orderly way to save the index and callback reference for a single menu item. </summary>
    /// <remarks><strong> This class is accessible to VBScript indirectly via the AddMenuItem and SetBalloonTipCallback methods. </strong></remarks>
    [Guid("2650C2AB-5FF8-495F-AB4D-6C61BD463EA4")]
    public class CallbackReference
    {
        /// <summary> The Index property corresponds to the index of a menu item in the notification area (system tray) context menu. </summary>
        public int Index { get; set; }

        /// <summary> COM object reference generated by VBScript's GetRef Function. </summary>
        public object Reference { get; set; }

        /// <summary> Constructor </summary>
        /// <parameters> index, reference </parameters>
        public CallbackReference(int index, object reference)
        {
            this.Index = index;
            this.Reference = reference;
        }
    }
}
