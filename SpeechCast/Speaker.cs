using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech.Synthesis;
using System.Diagnostics;

namespace SpeechCast
{
    class Speaker
    {
        /// <summary>
        /// 自身のインスタンス。シングルトンなクラスであることに注意。
        /// </summary>
        private static readonly Speaker instance = new Speaker();

        /// <summary>
        /// SAPIなエンジンへアクセスするためのインスタンス。
        /// </summary>
        private SpeechSynthesizer synthesizer = new SpeechSynthesizer();

        private Speaker() {

            // インストール済みのSAPIなボイスを列挙
            foreach (InstalledVoice voice in synthesizer.GetInstalledVoices())
            {
                // メンバ変数へ列挙されたボイスを追加。
                InstalledVoices.Add(voice);
            }
        }

        /// <summary>
        /// SAPIな音声合成エンジンの一覧。
        /// </summary>
        public List<InstalledVoice> InstalledVoices { get; protected set; }

        /// <summary>
        /// 現在、発声中かのフラグ。
        /// </summary>
        public bool IsSpeaking { get; protected set; }

        /// <summary>
        /// このクラスのインスタンスへアクセスするためのアクセサ。
        /// </summary>
        public static Speaker Instance { get { return instance; } }

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
        public SynthesizerType SpeakingType { get; protected set; } = SynthesizerType.None;

        /// <summary>
        /// SAPIな発声エンジンを設定するメソッド。
        /// </summary>
        /// <param name="voice">readonlyなメンバーinstalledVoiceの中のうちの1つ。</param>
        public bool SetSpeakingSAPIMethod(InstalledVoice voice)
        {
            synthesizer.SelectVoice(voice.ToString());

            SpeakingType = SynthesizerType.SAPI;

            return true;
        }

        /// <summary>
        /// 現在設定されている外部発声プログラム。
        /// </summary>
        public string SynthesizerProgram { get; protected set; } = "";

        /// <summary>
        /// 発声メソッドを外部プログラムに指定する。
        /// 指定できなかった場合は無指定となる。
        /// </summary>
        /// <param name="programName">外部プログラムへのパス。</param>
        public bool SetSpeakingProgram(string programName)
        {
            try
            {
                Process.Start(programName, "");
            }
            catch
            {
                SpeakingType = SynthesizerType.None;
                return false;
            }
            return true;
        }

        /// <summary>
        /// 発声メソッドが呼び出されてから発音した音(文章)。
        /// </summary>
        public string SpokeSentence { get; protected set; }
        
    }
}
