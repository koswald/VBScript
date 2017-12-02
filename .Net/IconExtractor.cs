
// Extract an icon from a .dll or .exe file

using System.Drawing;
using System.Runtime.InteropServices;
using System;

namespace VBScripting
{
    /// <summary> Extracts an icon from a .dll or .exe file. </summary>
    [Guid("2650C2AB-6AF8-495F-AB4D-6C61BD463EA4")]
    public class IconExtractor
    {
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        /// <summary> Extracts an icon from a .dll or .exe file. </summary>
        /// <param name="file"> A filespec string. </param>
        /// <param name="number"> An integer. The icon's index within the resource. </param>
        /// <param name="largeIcon"> A boolean. True for a large icon, False for a small icon. </param>
        /// <returns> An icon: See <see cref="Icon"/>. </returns>
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