// Extract an icon from a .dll or .exe file

using System.Drawing;
using System.Runtime.InteropServices;
using System;
using System.IO;
using System.Drawing.Imaging;

namespace VBScripting
{
    /// <summary> Extracts an icon from a .dll or .exe file. </summary>
    /// <remarks><span class="red"> Not all members of this class are accessible to VBScript. </span></remarks>
    [ProgId("VBScripting.IconExtractor"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-6AF8-495F-AB4D-6C61BD463EA4")]
    public class IconExtractor : IIconExtractor
    {
        private ImageFormat imageFormat;

        /// <summary> Constructor. </summary>
        public IconExtractor()
        {
            SetImageFormatBmp();
        }
        /// <summary> Extracts an icon from a .dll or .exe and saves it to a file. </summary>
        /// <parameters> resFile, index, icoFile, largeIcon </parameters>
        /// <remarks> Parameters: resFile is the .dll or .exe file; index selects the icon within the resource file; icoFile is the output file; largeIcon is a boolean: True if a large icon is to be extracted, False for a small icon. Environment variables and relative paths are allowed. </remarks>
        public void Save(string resFile, int index, string icoFile, bool largeIcon)
        {
            string inFile = Resolve(Expand(resFile));
            string outFile = Resolve(Expand(icoFile));
            int iconPointer = (int)IntPtr.Zero;
            FileStream stream = File.Create(outFile);
            Icon icon = null;
            try
            {
                iconPointer = GetPointer(inFile, index, largeIcon);
                icon = ExtractIcon(iconPointer);
                icon.ToBitmap().Save(stream, imageFormat);
            }
            catch (Exception e)
            {
                throw new ApplicationException(string.Format(
                        "VBScripting.IconExtractor.Save failed to save. \n e.Message: {0} \n\n resFile: {1} \n index: {2} \n icoFile: {3} \n largeIcon: {4} \n inFile: {5} \n outFile: {6} \n\n e.ToString(): {7}", 
                        e.Message, resFile, index, icoFile, largeIcon, inFile, outFile, e.ToString()
                    ), e);
            }
            finally
            {
                if (stream != null)
                    stream.Dispose();
                if (icon != null)
                    icon.Dispose();
            }
        }
        /// <summary> Change the image format to BMP. Default is BMP. </summary>
        public void SetImageFormatBmp()
        {
            imageFormat = ImageFormat.Bmp;
        }
        /// <summary> Change the image format to PNG. Default is BMP. </summary>
        public void SetImageFormatPng()
        {
            imageFormat = ImageFormat.Png;
        }
        // resolve relative path or no path => absolute path
        private string Resolve(string unresolved)
        {
            return System.IO.Path.GetFullPath(unresolved);
        }

        // expand environment strings
        private string Expand(string unexpanded)
        {
            return System.Environment.ExpandEnvironmentVariables(unexpanded);
        }

        [DllImport("Shell32", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        [DllImport("user32.dll", EntryPoint = "DestroyIcon", SetLastError=true)]
        private static extern int DestroyIcon(IntPtr pointer);

        /// <summary> Extracts an icon from the specified .dll or .exe file. <span class="red"> This method is static and so it is not directly available to VBScript. </span></summary>
        /// <parameters> file, index, largeIcon </parameters>
        /// <remarks> Other parameters: <tt>index</tt> is an integer that specifies the icon's index within the resource. <tt>largeIcon</tt> is a boolean that specifies whether the icon should be a large icon; if False, a small icon is extracted, if available. The icon must be disposed in order to free memory.</remarks>
        /// <returns> an icon </returns>
        public static Icon Extract(string file, int index, bool largeIcon)
        {
            IntPtr large = IntPtr.Zero;
            IntPtr small = IntPtr.Zero;
            try
            {
                ExtractIconEx(file, index, out large, out small, 1);
                return Icon.FromHandle(largeIcon ? large : small);
            }
            catch(Exception ex)
            {
                throw new ApplicationException(string.Format(
                        "Failed to extract icon from {0}", file
                    ), ex);
            }
        }
        /// <summary> Returns the number of icons in a .dll or .exe file. </summary>
        /// <parameters> filespec (.dll or .exe) </parameters>
        /// <returns> an int </returns>
        /// <remarks> A relative path or environmental variable is allowed. </remarks>
        public int IconCount(string file)
        {
            IntPtr largeIcons = IntPtr.Zero;
            IntPtr smallIcons = IntPtr.Zero;
            int count = -1;
            try
            {
                count = (int)ExtractIconEx(Resolve(Expand(file)), -1, out largeIcons, out smallIcons, 1);
            }
            finally
            {
                if (largeIcons != IntPtr.Zero)
                    DestroyIcon(largeIcons);

                if (smallIcons != IntPtr.Zero)
                    DestroyIcon(smallIcons);
            }
            return count;

        }
        /// <summary> Gets a pointer to an icon. </summary>
        /// <parameters> file, index, largeIcon</parameters>
        /// <returns> integer </returns>
        /// <remarks> Must be disposed with DisposeIcon(pointer) or Icon.Dispose(), in order to release memory. A relative path or environmental variable is allowed. </remarks>
        public int GetPointer(string file, int index, bool largeIcon)
        {
            IntPtr largeIcons = IntPtr.Zero;
            IntPtr smallIcons = IntPtr.Zero;
            ExtractIconEx(Resolve(Expand(file)), index, out largeIcons, out smallIcons, 1);
            return (int)(largeIcon ? largeIcons : smallIcons);
        }
        /// <summary> Gets an icon. </summary>
        /// <parameters> integer </parameters>
        /// <returns> Icon </returns>
        /// <remarks> Must be disposed with DisposeIcon(pointer) or Icon.Dispose(). </remarks>
        public Icon ExtractIcon(int pointer)
        {
            return Icon.FromHandle((IntPtr)pointer);
        }
        /// <summary> Dispose an icon by pointer (an int). </summary>
        /// <parameters> pointer </parameters>
        /// <returns> Returns true for success. </returns>
        public bool DisposeIcon(int pointer)
        {
            return DestroyIcon((IntPtr)pointer) != 0;
        }
    }

    /// <summary> A COM Interface for VBScripting.IconExtractor </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-6BF8-495F-AB4D-6C61BD463EA4")]
    public interface IIconExtractor
    {
        /// <summary> </summary>
        [DispId(0)]
        void Save(string resFile, int index, string icoFile, bool largeIcon);
        /// <summary> </summary>
        [DispId(1)]
        int IconCount(string file);
        /// <summary> </summary>
        [DispId(2)]
        void SetImageFormatBmp();
        /// <summary> </summary>
        [DispId(3)]
        void SetImageFormatPng();
    }

}