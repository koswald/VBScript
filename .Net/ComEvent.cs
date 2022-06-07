using System.Runtime.InteropServices;
using System;
using System.Reflection;

namespace VBScripting
{
    /// <summary> Invokes VBS members from C#. <strong>This class is not accessible from VBScript. </strong> </summary>
    [Guid("2650C2AB-7AF8-495F-AB4D-6C61BD463EA4")]
    public class ComEvent
    {
        /// <summary> Invokes a VBScript method. </summary>
        /// <parameters> callbackRef </parameters>
        /// <remarks> The parameter <code> callbackRef</code> is a reference to a VBScript member returned by the VBScript Function GetRef. </remarks>
        public static void InvokeComCallback(object callbackRef)
        {
            try
            {
                callbackRef.GetType().InvokeMember("",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, callbackRef, null);
            }
            catch {}
        }
    }
}
