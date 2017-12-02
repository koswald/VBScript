
// .Net speech synthesis library for VBScript

// requires an assembly reference to
// %ProgramFiles(x86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll

using System.Speech.Synthesis;
using System.Runtime.InteropServices;
using System.Linq; // for Cast<object>
using System.Collections.Generic; // for List
using System.Windows.Forms; // for MessageBox

namespace VBScripting
{
    /// <summary> Provide a wrapper for the .Net speech synthesizer 
    /// for VBScript, for demonstration purposes. </summary>
    [Guid("2650C2AB-2AF8-495F-AB4D-6C61BD463EA4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("VBScripting.SpeechSynthesis")]
    public class SpeechSynthesis : ISpeechSynthesis
    {
        private SpeechSynthesizer ss;
        private List<string> installedVoices;
        private List<string> voices; // installed & enabled

        /// <summary> Constructor </summary>
        public SpeechSynthesis()
        {
            this.ss = new SpeechSynthesizer();
            this.voices = new List<string>(); // installed, enabled voices
            this.installedVoices = new List<string>(); // installed voices
            string name;
            
            foreach (InstalledVoice voice in ss.GetInstalledVoices())
            {
                name = voice.VoiceInfo.Name;
                this.installedVoices.Add(name);
                if (voice.Enabled)
                {
                    this.voices.Add(name);
                }
            }
        }

        /// <summary> Convert text to speech. 
        /// <para> This method is synchronous. </para> </summary>
        public void Speak(string text)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                this.ss.Speak(text);
            }
        }

        /// <summary> Convert text to speech. 
        /// <para> This method is asynchronous. </para> </summary>
        public void SpeakAsync(string text)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                this.ss.SpeakAsync(text);
            }
        }

        /// <summary> Pause speech synthesis. </summary>
        public void Pause()
        {
            if (System.Speech.Synthesis.SynthesizerState.Speaking == this.ss.State)
            {
                this.ss.Pause();
            }
        }

        /// <summary> Resume speech synthesis. </summary>
        public void Resume()
        {
            if (System.Speech.Synthesis.SynthesizerState.Paused == this.ss.State)
            {
                this.ss.Resume();
            }
        }

        /// <summary> Gets an array of the names of the installed, enabled voides. 
        /// <para> Each element of the array can be used to set <see cref="Voice"/> </para> </summary>
        public object Voices()
        {
            return this.voices.Cast<object>().ToArray(); // convert to VBScript array
        }

        /// <summary> Gets or sets the current voice by name. 
        /// <para> A string. One of the names from the Voices array. </para> </summary>
        public string Voice
        {
            set
            {
                if (this.voices.Contains(value))
                {
                    this.ss.SelectVoice(value);
                }
                else if (this.installedVoices.Contains(value))
                {
                    string msg = string.Format(
                        "\"{0}\" is an installed voice but is not enabled.", value);
                    ShowInfoMessage(msg);
                }
                else
                {
                    string msg = string.Format(
                        "\"{0}\" is not an installed voice.", value);
                    ShowInfoMessage(msg);
                }
            }
            get
            {
                return this.ss.Voice.Name;
            }
        }

        // Shows a message box with the specified string
        private void ShowInfoMessage(string msg)
        {
            MessageBox.Show(msg, 
                   "SpeechSynthesis class",
                   MessageBoxButtons.OK, 
                   MessageBoxIcon.Information);
        }

        /// <summary> Disposes the SpeechSynthesis object's resources. </summary>
        public void Dispose()
        {
            if (this.ss != null)
            {
                this.ss.Dispose();
            }
        }

        /// <summary> Gets the state of the SpeechSynthesizer. 
        /// <para> Read only. Returns an integer equal to one of 
        /// the <see cref="State"/> method return values. </para> </summary>
        public int SynthesizerState
        {
            get
            {
                if (System.Speech.Synthesis.SynthesizerState.Ready == this.ss.State)
                {
                    return (int) VBScripting.SynthesizerState.Ready;
                }
                else if (System.Speech.Synthesis.SynthesizerState.Paused == this.ss.State)
                {
                    return (int) VBScripting.SynthesizerState.Paused;
                }
                else if (System.Speech.Synthesis.SynthesizerState.Speaking == this.ss.State)
                {
                    return (int) VBScripting.SynthesizerState.Speaking;
                }
                else
                {
                    return (int) VBScripting.SynthesizerState.Unexpected;
                }
            }
            private set { }
        }
        
        /// <summary> Gets or sets the volume. 
        /// <para> An integer from 0 to 100. </para> </summary>
        public int Volume
        {
            get
            {
                return this.ss.Volume;
            }
            set
            {
                this.ss.Volume = value;
            }
        }

        /// <summary> Provides an object whose methods (Ready, Paused, and Speaking) provide return values useful in VBScript for comparing to <see cref="SynthesizerState"/>. </summary>
        public SynthesizerStateT State
        {
            get { return new SynthesizerStateT(); }
            private set { }
        }
    }

    /// <summary> C# enum not intended for use by VBScript. 
    /// <para> Corresponds to but not equivalent to System.Speech.Synthesis.SynthesizerState. </para> </summary>
    [Guid("2650C2AB-2CF8-495F-AB4D-6C61BD463EA4")]
    public enum SynthesizerState : int
    {
        /// <summary> Return value can be cast to an int: 1 </summary>
        Ready = 1,
        /// <summary> Return value can be cast to an int: 2 </summary>
        Paused,
        /// <summary> Return value can be cast to an int: 3 </summary>
        Speaking,
        /// <summary> Return value can be cast to an int: 4 </summary>
        Unexpected
    }

    /// <summary> Supplies the type required by <see cref="SpeechSynthesis.State"/>
    /// <para> Not intended for use in VBScript. </para> </summary>
    [Guid("2650C2AB-2DF8-495F-AB4D-6C61BD463EA4")]
    public class SynthesizerStateT
    {
        /// <summary> Constructor </summary>
        public SynthesizerStateT() { }
        /// <summary> Returns an integer: 1 </summary>
        public int Ready { get { return (int) SynthesizerState.Ready; } private set { } }
        /// <summary> Returns an integer: 2 </summary>
        public int Paused { get { return (int) SynthesizerState.Paused; } private set { } }
        /// <summary> Returns an integer: 3 </summary>
        public int Speaking { get { return (int) SynthesizerState.Speaking; } private set { } }
        /// <summary> Returns an integer: 4 </summary>
        public int Unexpected { get { return (int) SynthesizerState.Unexpected; } private set { } }
    }

    /// <summary> The COM interface for <see cref="SpeechSynthesis"/>. </summary>
    [Guid("2650C2AB-2BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISpeechSynthesis
    {
        /// <summary> COM interface member for <see cref="Speak(string)"/></summary>
        [DispId(1)]
        void Speak(string text);

        /// <summary> COM interface member for <see cref="SpeakAsync(string)"/></summary>
        [DispId(2)]
        void SpeakAsync(string text);

        /// <summary> COM interface member for <see cref="Pause()"/></summary>
        [DispId(3)]
        void Pause();

        /// <summary> COM interface member for <see cref="Resume()"/></summary>
        [DispId(4)]
        void Resume();

        /// <summary> COM interface member for <see cref="Voices()"/></summary>
        [DispId(5)]
        object Voices(); // array

        /// <summary> COM interface member for <see cref="Voice"/></summary>
        [DispId(6)]
        string Voice { get; set; }

        /// <summary> COM interface member for <see cref="Dispose()"/></summary>
        [DispId(7)]
        void Dispose();

        /// <summary> COM interface member for <see cref="SynthesizerState"/></summary>
        [DispId(8)]
        int SynthesizerState { get; }

        /// <summary> COM interface member for <see cref="State"/></summary>
        [DispId(9)]
        SynthesizerStateT State { get; }

        /// <summary> COM interface member for <see cref="Volume"/></summary>
        [DispId(13)]
        int Volume { get; set; }
    }
}
