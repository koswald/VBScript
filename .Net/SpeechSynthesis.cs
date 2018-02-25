
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
    /// <summary> Provide a wrapper for the .Net speech synthesizer for VBScript, for demonstration purposes. </summary>
    /// <remarks> Requires an assembly reference to <tt>%ProgramFiles(x86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll</tt>, which may not be available on older machines. </remarks>
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

        /// <summary> Convert text to speech. </summary>
        /// <remarks> This method is synchronous. </remarks> 
        public void Speak(string text)
        {
            if (!string.IsNullOrWhiteSpace(text))
            {
                this.ss.Speak(text);
            }
        }

        /// <summary> Convert text to speech. </summary>
        /// <remarks> This method is asynchronous. </remarks>
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

        /// <summary> Gets an array of the names of the installed, enabled voices. </summary>
        /// <remarks> Each element of the array can be used to set <tt>Voice</tt> </remarks>
        public object Voices()
        {
            return this.voices.Cast<object>().ToArray(); // convert to VBScript array
        }

        /// <summary> Gets or sets the current voice by name. </summary>
        /// <remarks> A string. One of the names from the <tt>Voices</tt> array. </remarks>
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

        /// <summary> Gets the state of the SpeechSynthesizer. </summary>
        /// <remarks> Read only. Returns an integer equal to one of the <tt>State</tt> method return values. </remarks>
        public int SynthesizerState
        {
            get
            {
                if (System.Speech.Synthesis.SynthesizerState.Ready == this.ss.State)
                {
                    return 1;
                }
                else if (System.Speech.Synthesis.SynthesizerState.Paused == this.ss.State)
                {
                    return 2;
                }
                else if (System.Speech.Synthesis.SynthesizerState.Speaking == this.ss.State)
                {
                    return 3;
                }
                else
                {
                    return 4;
                }
            }
            private set { }
        }
        
        /// <summary> Gets or sets the volume. </summary>
        /// <remarks> An integer from 0 to 100. </remarks>
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

        /// <summary> Gets an object whose properties (Ready, Paused, and Speaking) provide values useful for comparing to <tt>SynthesizerState</tt>. </summary>
        /// <returns> a SynthersizerStateT </returns>
        public SynthesizerStateT State
        {
            get { return new SynthesizerStateT(); }
            private set { }
        }
    }

    /// <summary> Enumerates the synthesizer states. </summary>
    /// <remarks> Not intended for use in VBScript. See <tt>SpeechSynthesis.State</tt>. </remarks>
    [Guid("2650C2AB-2DF8-495F-AB4D-6C61BD463EA4")]
    public class SynthesizerStateT
    {
        /// <summary> Constructor </summary>
        public SynthesizerStateT() { }
        /// <returns> 1 </returns>
        public int Ready { get { return 1; } private set { } }
        /// <returns> 2 </returns>
        public int Paused { get { return 2; } private set { } }
        /// <returns> 3 </returns>
        public int Speaking { get { return 3; } private set { } }
        /// <returns> 4 </returns>
        public int Unexpected { get { return 4; } private set { } }
    }

    /// <summary> The COM interface for <tt>VBScripting.SpeechSynthesis</tt>. </summary>
    [Guid("2650C2AB-2BF8-495F-AB4D-6C61BD463EA4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISpeechSynthesis
    {
        /// <summary> </summary>
        [DispId(1)]
        void Speak(string text);

        /// <summary> </summary>
        [DispId(2)]
        void SpeakAsync(string text);

        /// <summary> </summary>
        [DispId(3)]
        void Pause();

        /// <summary> </summary>
        [DispId(4)]
        void Resume();

        /// <summary> </summary>
        [DispId(5)]
        object Voices(); // array

        /// <summary> </summary>
        [DispId(6)]
        string Voice { get; set; }

        /// <summary> </summary>
        [DispId(7)]
        void Dispose();

        /// <summary> </summary>
        [DispId(8)]
        int SynthesizerState { get; }

        /// <summary> </summary>
        [DispId(9)]
        SynthesizerStateT State { get; }

        /// <summary> </summary>
        [DispId(13)]
        int Volume { get; set; }
    }
}
