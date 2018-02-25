
// Extract an icon from a .dll or .exe file

using System.Drawing;
using System.Runtime.InteropServices;
using System;

namespace VBScripting
{
    /// <summary> Extracts an icon from a .dll or .exe file. </summary>
    /// <remarks><span class="red"> This class is not accessible to VBScript. </span></remarks>
    [Guid("2650C2AB-6AF8-495F-AB4D-6C61BD463EA4")]
    public class IconExtractor
    {
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        /// <summary> Extracts an icon from the specified .dll or .exe file. </summary>
        /// <parameters> file, number, largeIcon </parameters>
        /// <remarks> Other parameters: <tt>number</tt> is an integer that specifies the icon's index within the resource. <tt>largeIcon</tt> is a boolean that specifies whether the icon should be a large icon or small icon. </remarks>
        /// <returns> an icon </returns>
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

    }
}