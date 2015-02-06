using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;

using Microsoft.Kinect;

using K4W.KinectDrone.Client.Enums;
using K4W.KinectDrone.Client.Extensions;

namespace K4W.KinectDrone.Client
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        /// <summary>
        /// Representation of our Kinect-sensor
        /// </summary>
        private KinectSensor _currentSensor = null;


        /// <summary>
        /// WritebleBitmap that will draw the Kinect video output
        /// </summary>
        private WriteableBitmap _cameraVision = null;

        /// <summary>
        /// Buffer to copy the pixel data to
        /// </summary>
        private byte[] _pixelData = new byte[0];


        /// <summary>
        /// The RecognitionEngine used to build our grammar and start recognizing
        /// </summary>
        private SpeechRecognitionEngine _recognizer;

        /// <summary>
        /// The KinectAudioSource that is used.
        /// Basicly gets the Audio from the microphone array
        /// </summary>
        private KinectAudioSource _audioSource;

        /// <summary>
        /// Timestamp of the last successfully recognized command
        /// </summary>
        private DateTime _lastCommand = DateTime.Now;

        /// <summary>
        /// A constant defining the delay between 2 successful voice recognitions.
        /// It will be dropped if there already is one recognized in this interval
        /// </summary>
        private const float _delayInSeconds = 2;

        /// <summary>
        /// All speech commands and actions
        /// </summary>
        private readonly Dictionary<string, object> _speechActions = new Dictionary<string, object>()
        {
            { "Take off", VoiceCommand.TakeOff },
            { "Land", VoiceCommand.Land },
            { "May day", VoiceCommand.EmergencyLanding }
        };


        /// <summary>
        /// Default CTOR
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            // Kinect init
            InitializeKinect();

            // General events
            Closing += OnClosing;
        }

        /// <summary>
        /// Stop Kinect when closing the application
        /// </summary>
        private void OnClosing(object sender, CancelEventArgs e)
        {
            if (_currentSensor != null && _currentSensor.Status == KinectStatus.Connected)
                _currentSensor.Stop();
        }


        #region Kinect Global
        /// <summary>
        /// Initialisation of the Kinect
        /// </summary>
        private void InitializeKinect()
        {
            // Get current running sensor
            KinectSensor sensor = KinectSensor.KinectSensors.FirstOrDefault(sens => sens.Status == KinectStatus.Connected);

            // Initialize sensor
            StartSensor(sensor);

            // Sub to Kinect StatusChanged-event
            KinectSensor.KinectSensors.StatusChanged += OnKinectStatusChanged;
        }

        /// <summary>
        /// Start a new sensor
        /// </summary>
        private void StartSensor(KinectSensor sensor)
        {
            // Avoid crashes
            if (sensor == null)
                return;

            // Save instance
            _currentSensor = sensor;

            // Initialize color & skeletal tracking
            _currentSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);

            // Sub to events
            _currentSensor.ColorFrameReady += OnColorFrameReadyHandler;

            // Start sensor
            _currentSensor.Start();

            // Save sensor status
            KinectStatus = _currentSensor.Status;

            // Initialize speech
            InitializeSpeech();
        }

        /// <summary>
        /// Process a Kinect status change
        /// </summary>
        private void OnKinectStatusChanged(object sender, StatusChangedEventArgs e)
        {
            if (_currentSensor == null || _currentSensor.DeviceConnectionId != e.Sensor.DeviceConnectionId)
                return;

            // Save new status
            KinectStatus = e.Sensor.Status;

            // More later
        }
        #endregion Kinect Global

        #region Kinect Speech
        /// <summary>
        /// Initialize speech
        /// </summary>
        private void InitializeSpeech()
        {
            // Check if vocabulary is specifief
            if (_speechActions == null || _speechActions.Count == 0)
                throw new ArgumentException("A vocabulary is required.");

            // Check sensor state
            if (_currentSensor.Status != KinectStatus.Connected)
                throw new Exception("Unable to initialize speech if sensor isn't connected.");

            // Get the RecognizerInfo of our Kinect sensor
            RecognizerInfo info = GetKinectRecognizer();

            // Let user know if there is none.
            if (info == null)
                throw new Exception("There was a problem initializing Speech Recognition.\nEnsure that you have the Microsoft Speech SDK installed.");

            // Create new speech-engine
            try
            {
                _recognizer = new SpeechRecognitionEngine(info.Id);

                if (_recognizer == null) throw new Exception();
            }
            catch (Exception ex)
            {
                throw new Exception("There was a problem initializing Speech Recognition.\nEnsure that you have the Microsoft Speech SDK installed.");
            }

            // Add our commands as "Choices"
            Choices cmds = new Choices();
            foreach (string key in _speechActions.Keys)
                cmds.Add(key);

            /*
             * The GrammarBuilder defines what the requisted "flow" is of the possible commands.
             * You can insert plain text, or a Choices object with all our values in it, in our case our commands
             * We also need to pass in our Culture so that it knows what language we're talking
             */
            GrammarBuilder cmdBuilder = new GrammarBuilder { Culture = info.Culture };
            cmdBuilder.Append("Drone");
            cmdBuilder.Append(cmds);

            // Create our speech grammar
            Grammar cmdGrammar = new Grammar(cmdBuilder);

            // Prevent crashes
            if (_currentSensor == null || _recognizer == null)
                return;

            // Load grammer into our recognizer
            _recognizer.LoadGrammar(cmdGrammar);

            // Hook into speech events
            _recognizer.SpeechRecognized += OnCommandRecognizedHandler;
            _recognizer.SpeechRecognitionRejected += OnCommandRejectedHandler;

            // Get the kinect audio stream
            _audioSource = _currentSensor.AudioSource;

            // Set the beamangle
            _audioSource.BeamAngleMode = BeamAngleMode.Adaptive;

            // Start the kinect audio
            Stream kinectStream = _audioSource.Start();

            // Assign the stream to the recognizer along with FormatInfo
            _recognizer.SetInputToAudioStream(kinectStream, new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));

            // Start recognizingand make sure to tell that the RecognizeMode is Multiple or it will stop after the first recognition
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);
        }

        /// <summary>
        /// Command recognized
        /// </summary>
        private void OnCommandRecognizedHandler(object sender, SpeechRecognizedEventArgs e)
        {
            TimeSpan interval = DateTime.Now.Subtract(_lastCommand);

            if (interval.TotalSeconds < _delayInSeconds)
                return;

            if (e.Result.Confidence < 0.80f)
                return;

            // Retrieve the DroneAction from the recognized result
            VoiceCommand invokedAction = GetDroneAction(e.Result);

            // Log action
            WriteLog("Command '" + invokedAction.GetDescription() + "' recognized.");

            /*
             * 
             * Coming in Part II
             * 
             * 
             */

        }

        /// <summary>
        /// Command rejected
        /// </summary>
        private void OnCommandRejectedHandler(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            WriteLog("Unkown command");
        }

        /// <summary>
        /// Get the first RecognizerInfo-object that is a Kinect & has the English pack
        /// </summary>
        private RecognizerInfo GetKinectRecognizer()
        {
            /* Create a function that checks if the additioninfo contains a key called "Kinect" and if it's true.
             * Also check if the culture is en-US so that we're using the English pack
            */
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);

                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) &&
                       "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };

            return SpeechRecognitionEngine.InstalledRecognizers().FirstOrDefault(matchingFunc);
        }

        /// <summary>
        /// Convert recognized command to an enumeration representation
        /// </summary>
        private VoiceCommand GetDroneAction(RecognitionResult recogResult)
        {
            if (recogResult == null || string.IsNullOrEmpty(recogResult.Text))
                return VoiceCommand.Unknown;

            // Seperate 'Drone' from command grammar
            Match m = Regex.Match(recogResult.Text, "^Drone (.*)$");

            // Check if it matches
            if (m.Success)
            {
                // Get command from object
                KeyValuePair<string, object> cmd =
                   _speechActions.FirstOrDefault(action => action.Key.ToLower() == m.Groups[1].ToString().ToLower());

                if (cmd.Value != null)
                {
                    return (VoiceCommand)cmd.Value;
                }
                return VoiceCommand.Unknown;
            }
            else return VoiceCommand.Unknown;
        }
        #endregion Kinect Speech

        #region Kinect camera
        /// <summary>
        /// Process color data
        /// </summary>
        private void OnColorFrameReadyHandler(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame colorFrame = e.OpenColorImageFrame())
            {
                if (colorFrame == null)
                    return;

                // Initialize variables
                if (_pixelData.Length == 0)
                {
                    // Create buffer
                    _pixelData = new byte[colorFrame.PixelDataLength];

                    // Create output rep
                    _cameraVision = new WriteableBitmap(colorFrame.Width,
                                                        colorFrame.Height,

                                                        // DPI
                                                        96, 96,

                                                        // Current pixel format
                                                        PixelFormats.Bgr32,

                                                        // Bitmap palette
                                                        null);

                    // Hook image to Image-control
                    KinectImage.Source = _cameraVision;
                }

                // Copy data from frame to buffer
                colorFrame.CopyPixelDataTo(_pixelData);

                // Update bitmap
                _cameraVision.WritePixels(

                    // Image size
                    new Int32Rect(0, 0, colorFrame.Width, colorFrame.Height),

                    // Buffer
                    _pixelData,

                    // Stride
                    colorFrame.Width * colorFrame.BytesPerPixel,

                    // Buffer offset
                    0);
            }
        }
        #endregion Kinect camera


        #region UI Properties
        private KinectStatus _kinectStatus = KinectStatus.Error;
        public KinectStatus KinectStatus
        {
            get { return _kinectStatus; }
            set
            {
                if (_kinectStatus != value)
                {
                    _kinectStatus = value;
                    OnPropertyChanged("KinectStatus");
                }
            }
        }

        private string _log = "> Welcome!";
        public string Log
        {
            get { return _log; }
            set
            {
                if (_log != value)
                {
                    _log = value;
                    OnPropertyChanged("Log");
                }
            }
        }
        #endregion UI Properties

        #region Internal Methods/Events & UI Properties
        private void WriteLog(string output)
        {
            Log = "> " + output;
        }
        #endregion Internal Methods

        #region Internal events
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion Internal Events
    }
}
