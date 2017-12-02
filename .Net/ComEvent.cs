
// Provide VBS method calling from C#

using System.Runtime.InteropServices;
using System;
using System.Reflection; 

namespace VBScripting
{
    /// <summary> Invoke VBS methods from C#. </summary>
    [Guid("2650C2AB-7AF8-495F-AB4D-6C61BD463EA4")]
    public class ComEvent
    {

        /// <summary> Invokes a VBScript method. </summary>
        /// <remarks> This method is not callable from VBScript. </remarks>
        /// <param name="callbackRef"> A reference to a VBScript method. </param> 
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
                string msg = string.Format(
                    "VBScripting.ComEvent.InvokeComCallback: {0}\n\n{1}",
                    tie.Message,
                    "This is an expected error when the invoked method is in  " +
                    "a .vbs or .wsf script (as opposed to an .hta script), " +
                    "but it does not appear to affect the success of invoking the method."
                );
                Admin.Log(msg);
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
