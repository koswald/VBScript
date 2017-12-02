
// System tray icon for VBScript

using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Collections.Generic;
using System;
using System.Reflection;

namespace VBScripting
{
    /// <summary> Provides a system tray icon for VBScript, for illustration purposes. </summary>
    [ProgId("VBScripting.NotifyIcon"),
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
            this.notifyIcon.ContextMenu = this.contextMenu;
            this.Debug = false;
        }

        /// <summary> 
        /// <para> Gets or sets whether the type is under development. </para>
        /// <para> Affects the behavior of two methods, if exceptions are thrown. See <see cref="SetIconByIcoFile(string)"/> and <see cref="SetIconByDllFile(string, int, bool)"/> </para>
        /// </summary>
        public bool Debug { get; set; }

        /// <summary> Gets or sets the text shown when the mouse hovers over the system tray icon. </summary>
        public string Text
        {
            get { return this.notifyIcon.Text; }
            set { this.notifyIcon.Text = value; }
        }

        /// <summary> Gets or sets the icon's visibility. A boolean. </summary>
        public bool Visible
        {
            get { return this.notifyIcon.Visible; }
            set { this.notifyIcon.Visible = value; }
        }

        /// <summary> Sets the system tray <see cref="Icon"/> given an .ico file. </summary>
        /// <param name="fileName"> The filespec of the .ico file. Environment variables are allowed. </param>
        public void SetIconByIcoFile(string fileName)
        {
            try
            {
                this.notifyIcon.Icon = new Icon(Environment.ExpandEnvironmentVariables(fileName));
            }
            catch (Exception e)
            {
                Admin.Log(string.Format(
                    "File: {0}\n\n{1}",
                    fileName, e.ToString()
                ));
                this.Dispose();
                throw;
            }
        }

        /// <summary> Sets the system tray <see cref="Icon"/> from a .dll or .exe file. </summary>
        /// <param name="fileName"> The path and name of a .dll or .exe file that contains icons. </param>
        /// <param name="index"> The index of the icon. An integer. </param>
        /// <param name="largeIcon"> A boolean: true extracts a loarge icon, false extracts a large icon. </param>
        public void SetIconByDllFile(string fileName, int index, bool largeIcon)
        {
            try
            {
                this.notifyIcon.Icon = VBScripting.IconExtractor.Extract(System.Environment.ExpandEnvironmentVariables(fileName), index, largeIcon);
            }
            catch (Exception e)
            {
                Admin.Log(string.Format(
                    "File: {0}\nIndex: {1}\n\n{2}",
                    fileName, index, e.ToString()
                ));
                this.Dispose();
                throw;
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

        /// <summary> Gets or sets the lifetime of the "balloon tip" or notification. An integer (milliseconds). Deprecated as of Windows Vista. Now the value is overridden by accessibility settings.  </summary>
        public int BalloonTipLifetime { get; set; }

        /// <summary> Provides an object useful in VBScript for selecting a ToolTipIcon type. The methods (Error, Info, None, Warning) may be used with <see cref="SetBalloonTipIcon(int)"/>. </summary>
        public ToolTipIconT ToolTipIcon
        {
            get { return new ToolTipIconT(); }
            private set { }
        }

        /// <summary> Sets the icon of the "balloon tip" or notification. </summary>
        /// <param name="type"> An integer. Return values of one of the four methods (Error, Info, None, Warning) of <see cref="ToolTipIcon"/> can be used. </param>
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

        /// <summary> 
        /// Disposes of the icon resources when it is no longer needed.
        /// <para> If this method is not called, the icon may persist in the system tray until the mouse hovers over it, even after the object instance has lost scope. </para>
        /// </summary>
        public void Dispose()
        {
            this.notifyIcon.Icon.Dispose();
            this.notifyIcon.Dispose();
        }

        /// <summary> Show the baloon tip. </summary>
        public void ShowBalloonTip()
        {
            this.notifyIcon.ShowBalloonTip(this.BalloonTipLifetime);
        }

        /// <summary>
        /// Add a menu item to the system tray icon's context menu.
        /// <para> This method can be called only from VBScript. </para>
        /// </summary>
        /// <param name="menuText"> string: The text that appears in the menu. </param>
        /// <param name="callbackRef"> object: The VBScript object reference returned by GetRef in VBScript. </param>
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

        // show the context menu on left mouse click too
        private void notifyIcon_MouseUp(object Sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.ShowContextMenu();
            }
        }
        
        /// <summary> Show the context menu. </summary>
        /// <remarks> Public in order to provide testability from VBScript. </remarks>
        public void ShowContextMenu()
        {
            MethodInfo mi = typeof(System.Windows.Forms.NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
            mi.Invoke(notifyIcon, null);
        }
    }

    /// <summary> The COM interface for <see cref="NotifyIcon"/>. </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-5BF8-495F-AB4D-6C61BD463EA4")]
    public interface INotifyIcon
    {
        /// <summary> COM interface member for <see cref="NotifyIcon.Text"/></summary>
        [DispId(1)]
        string Text { get; set; }

        /// <summary> COM interface member for <see cref="NotifyIcon.Dispose()"/></summary>
        [DispId(2)]
        void Dispose();

        /// <summary> COM interface member for <see cref="NotifyIcon.Visible"/></summary>
        [DispId(3)]
        bool Visible { get; set; }

        /// <summary> COM interface member for <see cref="NotifyIcon.Debug"/></summary>
        [DispId(4)]
        bool Debug { get; set; }

        /// <summary> COM interface member for <see cref="NotifyIcon.SetIconByIcoFile(string)"/></summary>
        [DispId(5)]
        void SetIconByIcoFile(string file);

        /// <summary> COM interface member for <see cref="NotifyIcon.SetIconByDllFile(string, int, bool)"/></summary>
        [DispId(6)]
        void SetIconByDllFile(string file, int index, bool largeIcon);

        /// <summary> COM interface member for <see cref="NotifyIcon.BalloonTipTitle"/></summary>
        [DispId(7)]
        string BalloonTipTitle { get; set; }

        /// <summary> COM interface member for <see cref="NotifyIcon.BalloonTipText"/></summary>
        [DispId(8)]
        string BalloonTipText { get; set; }

        /// <summary> COM interface member for <see cref="NotifyIcon.BalloonTipLifetime"/></summary>
        [DispId(9)]
        int BalloonTipLifetime { get; set; } // milliseconds

        /// <summary> COM interface member for <see cref="NotifyIcon.ToolTipIcon"/></summary>
        [DispId(10)]
        ToolTipIconT ToolTipIcon { get; }

        /// <summary> COM interface member for <see cref="NotifyIcon.SetBalloonTipIcon(int)"/></summary>
        [DispId(11)]
        void SetBalloonTipIcon(int type);

        /// <summary> COM interface member for <see cref="NotifyIcon.ShowBalloonTip"/></summary>
        [DispId(12)]
        void ShowBalloonTip();

        /// <summary> COM interface member for <see cref="AddMenuItem(string, object)"/> </summary>
        [DispId(13)]
        void AddMenuItem(string menuText, object callbackRef);

        /// <summary> COM interface member for <see cref="InvokeCallbackByIndex(int)"/> </summary>
        [DispId(14)]
        void InvokeCallbackByIndex(int index);

        /// <summary> COM interface member for <see cref="ShowContextMenu"/> </summary>
        [DispId(15)]
        void ShowContextMenu();
    }

    /// <summary> C# enum not intended for use by VBScript. 
    /// <para> Corresponds to but not equivalent to System.Windows.Forms.ToolTipIcon. </para> </summary>
    [Guid("2650C2AB-5CF8-495F-AB4D-6C61BD463EA4")]
    public enum ToolTipIcon : int
    {
        /// <summary> Return value can be cast to an int: 1 </summary>
        Error = 1,
        /// <summary> Return value can be cast to an int: 2 </summary>
        Info,
        /// <summary> Return value can be cast to an int: 3 </summary>
        None,
        /// <summary> Return value can be cast to an int: 4 </summary>
        Warning
    }

    /// <summary> Supplies the type required by <see cref="NotifyIcon.ToolTipIcon"/>
    /// <para> Not intended for use in VBScript. </para> </summary>
    [Guid("2650C2AB-5DF8-495F-AB4D-6C61BD463EA4")]
    public class ToolTipIconT
    {
        /// <returns> Returns 1 </returns>
        public int Error
        {
            get { return (int)ToolTipIcon.Error; }
            private set { }
        }
        /// <returns> Returns 2 </returns>
        public int Info
        {
            get { return (int)ToolTipIcon.Info; }
            private set { }
        }
        /// <returns> Returns 3 </returns>
        public int None
        {
            get { return (int)ToolTipIcon.None; }
            private set { }
        }
        /// <returns> Returns 4 </returns>
        public int Warning
        {
            get { return (int)ToolTipIcon.Warning; }
            private set { }
        }
    }
    /// <summary>
    /// Settings for saving VBScript method references.
    /// </summary>
    [Guid("2650C2AB-5EF8-495F-AB4D-6C61BD463EA4")]
    public class CallbackEventSettings
    {
        /// <summary> A List of callback references. </summary>
        public List<CallbackReference> Refs { get; set; }

        /// <summary> Constructor </summary>
        public CallbackEventSettings()
        {
            this.Refs = new List<CallbackReference>();
        }

        /// <summary>
        /// Adds a CallbackReference instance reference to the List.
        /// </summary>
        /// <param name="callbackRef"></param>
        public void AddRef(CallbackReference callbackRef)
        {
            if (callbackRef != null && !(this.Refs.Contains(callbackRef)))
            {
                this.Refs.Add(callbackRef);
            }
        }
    }
    /// <summary> An orderly way to save the index and callback reference for a single menu item. </summary>
    [Guid("2650C2AB-5FF8-495F-AB4D-6C61BD463EA4")]
    public class CallbackReference
    {
        /// <summary> This Index corresponds to the Index of a menuItem in the context menu. </summary>
        public int Index { get; set; }
        /// <summary> COM object generated by VBScript's GetRef Function. </summary>
        public object Reference { get; set; }

        /// <summary> Constructor </summary>
        /// <param name="index"><see cref="CallbackReference.Index"/></param>
        /// <param name="reference"> See <see cref="CallbackReference.Reference"/></param>
        public CallbackReference(int index, object reference)
        {
            this.Index = index;
            this.Reference = reference;
        }
    }
}
