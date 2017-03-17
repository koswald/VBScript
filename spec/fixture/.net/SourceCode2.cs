// fixture for DotNetCompiler.spec.vbs

using System.Speech.Synthesis;
using System.Runtime.InteropServices;
using System.Reflection;

[assembly:AssemblyKeyFileAttribute("SourceCode2.snk")]

namespace vox
{
    [Guid("2650C2AD-2BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IVox
    {
        [DispId(1)]
        void say(string myLine);
    }

    [Guid("2650C2AD-2AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("SourceCode2")]
    public class Vox : IVox
    {
        private SpeechSynthesizer ed = 
		new SpeechSynthesizer();

        public void say(string myLine)
        {
            ed.Speak(myLine);
        }
    }
}
