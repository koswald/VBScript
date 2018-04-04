
// Provide VBS method calling from C#

using System.Runtime.InteropServices;
using System;
using System.Reflection; 

namespace VBScripting
{
    /// <summary> Invokes VBS methods from C#. <span class="red"> This class is not callable from VBScript. </span></summary>
    [Guid("2650C2AB-7AF8-495F-AB4D-6C61BD463EA4")]
    public class ComEvent
    {
        /// <summary> Invokes a VBScript method. </summary>
        /// <remarks> The parameter <tt>callbackRef</tt> is an object reference to a VBScript member returned by the VBScript Function GetRef. </remarks>
        public static void InvokeComCallback(object callbackRef)
        {
            try
            {
                callbackRef.GetType().InvokeMember("",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, callbackRef, null);
            }
            catch (TargetInvocationException tie) 
           {
               // this is an expected exception
               // for .vbs and .wsf files
               // (but not for .hta files).
           }
            catch (Exception e)
            {
                Admin.Log(string.Format(
                    "VBScripting.ComEvent.InvokeComCallback:\n{0}\n\n{1}",
                    e.Message, e.ToString()
                ));
            }
        }
    }
}
