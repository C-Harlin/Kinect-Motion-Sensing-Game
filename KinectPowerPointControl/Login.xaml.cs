using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Navigation;
using Microsoft.Kinect;
using Microsoft.Speech.Recognition;
using System.Threading;
using System.IO;
using Microsoft.Speech.AudioFormat;
using System.Diagnostics;
using System.Windows.Threading;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Windows.Media.Animation;
using System.ComponentModel;




namespace KinectPowerPointControl
{
    /// <summary>
    /// Login.xaml 的交互逻辑
    /// </summary>
    public partial class Login : Window
    {
        /// <summary>
        /// kinect传感器对象
        /// </summary>
        KinectSensor sensor;

        /// <summary>
        /// 语音识别对象
        /// </summary>
        SpeechRecognitionEngine speechRecognizer;
        private DispatcherTimer readyTimer;

        public Login()
        {
            InitializeComponent();
            button.Focus();
            this.Loaded += new RoutedEventHandler(Login_Loaded);
        }

        void Login_Loaded(object sender, RoutedEventArgs e)
        {
            sensor = KinectSensor.KinectSensors.FirstOrDefault();

            if (sensor == null)
            {
                MessageBox.Show("This application requires a Kinect sensor.");
                this.Close();
            }

            sensor.Start();
            //2017/12/21 by czl

            SpeechSynthesizer voice = new SpeechSynthesizer();   //创建语音实例
            voice.Rate = 1; //设置语速,[-10,10]
            voice.Volume = 100; //设置音量,[0,100]
            voice.SpeakAsync("Device connected");  //播放指定的字符串

            
            //2017/10/21 by zzl
            var parameters = new TransformSmoothParameters
            {
                Smoothing = 0.5f,
                Correction = 0.5f,
                Prediction = 0.5f,
                JitterRadius = 0.05f,
                MaxDeviationRadius = 0.04f
            };


            voice.Rate = 1; //设置语速,[-10,10]
            voice.Volume = 100; //设置音量,[0,100]
            voice.SpeakAsync("Thank you for playing this game");

            voice.Rate = 1; //设置语速,[-10,10]
            voice.Volume = 100; //设置音量,[0,100]
            voice.SpeakAsync("Please select the game mode");
            //2017/11/29 by zzl
            
            sensor.SkeletonStream.Enable(parameters);            
        }
        //------------------------------
        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        private void InitializeSpeechRecognition()
        {
            RecognizerInfo ri = GetKinectRecognizer();
            if (ri == null)
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                speechRecognizer = new SpeechRecognitionEngine(ri.Id);
            }
            catch
            {
                MessageBox.Show(
                    @"There was a problem initializing Speech Recognition.
Ensure you have the Microsoft Speech SDK installed and configured.",
                    "Failed to load Speech SDK",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }

            var phrases = new Choices();
            phrases.Add("play");



            var gb = new GrammarBuilder();
            //Specify the culture to match the recognizer in case we are running in a different culture.                                 
            gb.Culture = ri.Culture;
            gb.Append(phrases);

            // Create the actual Grammar instance, and then load it into the speech recognizer.
            var g = new Grammar(gb);

            speechRecognizer.LoadGrammar(g);
            speechRecognizer.SpeechRecognized += SreSpeechRecognized;
            speechRecognizer.SpeechHypothesized += SreSpeechHypothesized;
            speechRecognizer.SpeechRecognitionRejected += SreSpeechRecognitionRejected;

            this.readyTimer = new DispatcherTimer();
            this.readyTimer.Tick += this.ReadyTimerTick;
            this.readyTimer.Interval = new TimeSpan(0, 0, 4);
            this.readyTimer.Start();

        }

        private void ReadyTimerTick(object sender, EventArgs e)
        {
            this.StartSpeechRecognition();
            this.readyTimer.Stop();
            this.readyTimer.Tick -= ReadyTimerTick;
            this.readyTimer = null;
        }
        //------------------------------
        private void StartSpeechRecognition()
        {
            if (sensor == null || speechRecognizer == null)
                return;

            var audioSource = this.sensor.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();

            speechRecognizer.SetInputToAudioStream(
                    kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);

        }

        void SreSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            Trace.WriteLine("\nSpeech Rejected, confidence: " + e.Result.Confidence);
        }

        void SreSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            Trace.Write("\rSpeech Hypothesized: \t{0}", e.Result.Text);
        }

        void SreSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //This first release of the Kinect language pack doesn't have a reliable confidence model, so 
            //we don't use e.Result.Confidence here.
            if (e.Result.Confidence < 0.70)
            {
                Trace.WriteLine("\nSpeech Rejected filtered, confidence: " + e.Result.Confidence);
                return;
            }

            Trace.WriteLine("\nSpeech Recognized, confidence: " + e.Result.Confidence + ": \t{0}", e.Result.Text);
            if (e.Result.Text == "play")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent,button));//模拟触发button事件
                    //button_Click(sender,e);

                });
            }
        }

      

        private void button_Click(object sender, RoutedEventArgs e)
        {
            button.Focus();
            MainWindow window = new MainWindow();
            this.Close();
            window.Show();
        }
    }
}
