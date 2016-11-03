
using System.Speech.Synthesis;
using System.Runtime.InteropServices;
using System.Reflection;

[assembly:AssemblyKeyFileAttribute("Vox.snk")]

namespace vox
{
    [Guid("2650bE95-8A27-4ae6-B6DE-0542A0FC7039")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _Vox
    {
        [DispId(1)]
        void say(string myLine);
    }

    [Guid("2650b2AD-4BF8-495f-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("Vox")]
    public class Vox : _Vox
    {
        private SpeechSynthesizer ed = 
		new SpeechSynthesizer();

        public void say(string myLine)
        {
            ed.Speak(myLine);
        }
    }
}
