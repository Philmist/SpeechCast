using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech.Synthesis;
using System.Diagnostics;

namespace SpeechCast
{
    public class Speaker
    {

        /// <summary>
        /// SAPIなエンジンへアクセスするためのインスタンス。
        /// </summary>
        protected static SpeechSynthesizer synthesizer;

        private Speaker() {

            if (synthesizer == null)
            {
                synthesizer = new SpeechSynthesizer();
                synthesizer.Volume = 100;
            }

            if (InstalledVoices == null)
            {
                InstalledVoices = new List<InstalledVoice>();
            }

            // インストール済みのSAPIなボイスを列挙
            foreach (InstalledVoice voice in synthesizer.GetInstalledVoices())
            {
                // メンバ変数へ列挙されたボイスを追加。
                if (voice != null)
                {
                    InstalledVoices.Add(voice);
                }
            }

            synthesizer.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(SpeakingEnd);
            synthesizer.SpeakProgress += new EventHandler<SpeakProgressEventArgs>(ASPISpeakingSentence);
        }

        /// <summary>
        /// 自身のインスタンス。シングルトンなクラスであることに注意。
        /// </summary>
        private static Speaker instance;

        /// <summary>
        /// SAPIな音声合成エンジンの一覧。
        /// </summary>
        public List<InstalledVoice> InstalledVoices { get; protected set; } = new List<InstalledVoice>();

        /// <summary>
        /// このクラスのインスタンスへアクセスするためのアクセサ。
        /// </summary>
        public static Speaker Instance { get {
                if (instance == null)
                {
                    instance = new Speaker();
                }

                return instance;
            }
        }

        /// <summary>
        /// 現在の発声メソッドを表わすための型。
        /// </summary>
        public enum SynthesizerType
        {
            SAPI,
            CommandLine,
            None
        }

        /// <summary>
        /// 現在選択されている発声メソッド。
        /// </summary>
        public static SynthesizerType SpeakingType { get; protected set; } = SynthesizerType.None;

        /// <summary>
        /// SAPIな発声エンジンを設定するメソッド。
        /// </summary>
        /// <param name="voice">readonlyなメンバーinstalledVoiceの中のうちの1つ。</param>
        public bool SetSpeakingSAPIMethod(InstalledVoice voice)
        {
            try
            {
                synthesizer.SelectVoice(voice.VoiceInfo.Name);

                SpeakingType = SynthesizerType.SAPI;
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 現在設定されている外部発声プログラム。
        /// </summary>
        public static string SynthesizerProgram { get; protected set; } = "";

        /// <summary>
        /// 発声メソッドを外部プログラムに指定する。
        /// 指定できなかった場合は無指定となる。
        /// </summary>
        /// <param name="programName">外部プログラムへのパス。</param>
        public bool SetSpeakingProgram(string programName)
        {
            try
            {
                if (programName == "")
                {
                    ExecuteFileName = "";
                    SpeakingType = SynthesizerType.None;
                    return false;
                }
                Process.Start(programName, "");
                ExecuteFileName = programName;
                SpeakingType = SynthesizerType.CommandLine;
            }
            catch
            {
                return false;
            }
            return true;
        }


        private Process ExecutedProcess = null;
        private string ExecuteFileName = "";

        protected static string SpeakingSentence = "";

        public virtual void SpeakSentence(string sentence)
        {
            sentence = MMFrame.Text.Language.Japanese.ToKatakanaFromKatakanaHalf(sentence);
            switch (SpeakingType)
            {
                case SynthesizerType.CommandLine:
                    if (ExecuteFileName == "")
                    {
                        return;
                    }
                    ExecutedProcess.StartInfo = new ProcessStartInfo(ExecuteFileName, sentence);
                    ExecutedProcess.SynchronizingObject = null;
                    ExecutedProcess.EnableRaisingEvents = true;
                    ExecutedProcess.Exited += new EventHandler(SpeakingEnd);
                    IsSpeaking = true;
                    SpokenSentence spokenSentence = new SpokenSentence();
                    spokenSentence.Sentence = sentence;
                    SpeakingSentence = sentence;
                    ExecutedProcess.Start();
                    break;

                case SynthesizerType.SAPI:
                    IsSpeaking = true;
                    synthesizer.SpeakAsync(sentence);
                    break;

                case SynthesizerType.None:
                default:
                    break;
            }
        }

        protected void SpeakingEnd(object sender, EventArgs eventArgs)
        {
            IsSpeaking = false;
            SpeakingSentence = "";
            EventHandler eventHandler = SpeakingEndEvent;
            if (eventHandler != null)
            {
                eventHandler(this, eventArgs);
            }
            
        }

        protected void ASPISpeakingSentence(object sender, SpeakProgressEventArgs eventArgs)
        {
            IsSpeaking = true;

            int index = eventArgs.CharacterPosition + eventArgs.CharacterCount;
            if (index > 0)
            {
                index = index - 1;
            }

            if (index > SpeakingSentence.Length)
            {
                index = SpeakingSentence.Length;
            }

            string SpokenString = SpeakingSentence.Substring(0, index);
            SpokenSentence spokenSentence = new SpokenSentence();
            spokenSentence.Sentence = SpokenString;
            
        }

        protected virtual void OnSpokenSentenceEvent(SpokenSentence spokenSentence)
        {
            EventHandler<SpokenSentence> handler = SpokenSentenceEvent;
            if (handler != null)
            {
                handler(this, spokenSentence);
            }
        }

        public event EventHandler<SpokenSentence> SpokenSentenceEvent;
        public event EventHandler SpeakingEndEvent;

        /// <summary>
        /// 現在、発声中かのフラグ。
        /// </summary>
        public static bool IsSpeaking { get; protected set; } = false;
    }

    public class SpokenSentence : EventArgs
    {
        public string Sentence { get; set; }
    }

    public delegate void SpokenSentenceEventHandler(object sender, SpokenSentence eventArgs);

}
