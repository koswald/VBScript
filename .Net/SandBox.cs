
using System;
using System.Diagnostics; // for EventLog
using System.Linq; // for Cast
using System.Collections.Generic; // for List
using Microsoft.Win32; // for RegistryKey
using System.IO; // for FileInfo
using System.Runtime.InteropServices;

namespace VBScripting
{
    /// <summary> Proof of concept testing. </summary>
    [ProgId("VBScripting.SandBox"),
        ClassInterface(ClassInterfaceType.None),
        Guid("2650C2AB-AAF8-495F-AB4D-6C61BD463EA4")]
    public class SandBox : ISandBox
    {
        public int AdHoc()
        {
            return 1;
        }
    }

    /// <summary> COM interface for <see cref="SandBox"/> </summary>
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
        Guid("2650C2AB-ABF8-495F-AB4D-6C61BD463EA4")]
    public interface ISandBox
    {
        /// <summary> </summary>
        int AdHoc();
    }
}
