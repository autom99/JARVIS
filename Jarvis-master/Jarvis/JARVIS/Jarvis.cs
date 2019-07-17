using System;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
//using iTunesLib;
//using Microsoft.Office.Interop.Word;
using System.Drawing;

namespace JARVIS
{
    public partial class Jarvis : Form
    {
        public static WelcomeAndGoodbye wAndG;
        public static Email email;
        public static Speak speak = new Speak();
        public static Calender calender;
        public static Weather weather;
        public static string choice;
        //public static iTunesApp app;
        public static int volume = 50;
        public static SpeechRecognitionEngine myVoice = new SpeechRecognitionEngine();
        System.Windows.Forms.Timer stopListeningTimer = new System.Windows.Forms.Timer();

        public Jarvis()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.None;

            //Set output, load grammar and set up speech recognisition event handler
            myVoice.SetInputToDefaultAudioDevice();

            string[] phrases = GetPhrases();
            myVoice.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(phrases))));

            myVoice.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(MyVoice_SpeechRecognized);
            myVoice.RecognizeAsync(RecognizeMode.Multiple);

            LoadUpWelcome();

            stopListeningTimer.Tick += new EventHandler(Time_Tick);
            stopListeningTimer.Interval = 1000;
        }

        private void MyVoice_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string input = e.Result.Text;
            Console.WriteLine(input);

            switch (input.ToUpper())
            {
                case ("HOW IS THE WEATHER"):
                case ("WEATHER"):
                case ("HOWS THE WEATHER"):
                case ("HOW'S THE WEATHER"):
                    GetWeather();
                    break;

                case ("TOMORROWS FORECAST"):
                case ("TOMORROWS WEATHER"):
                case ("HOWS TOMORROWS WEATHER"):
                case ("HOW IS TOMORROWS WEATHER"):
                case ("WHATS TOMORROW LIKE"):
                case ("WHATS IT LIKE TOMORROW"):
                case ("FORECAST"):
                    GetForecast();
                    break;

                case ("MAIL"):
                case ("EMAIL"):
                case ("EMAILS"):
                case ("OPEN MAIL"):
                case ("OPEN EMAIL"):
                case ("OUTLOOK"):
                case ("SHOW MAIL"):
                case ("SHOW EMAIL"):
                case ("SHOW EMAILS"):
                    OpenMail();
                    break;

                case ("SEND MAIL"):
                case ("SEND EMAIL"):
                case ("SEND EMAILS"):
                case ("SEND AN EMAIL"):
                case ("SEND MESSAGE"):
                    SendMail();
                    break;
                    
                case ("READ MAIL"):
                case ("READ EMAIL"):
                case ("READ EMAILS"):
                    ReadMail();
                    break;

                case ("SEARCH"):
                case ("SEARCH FOR"):
                case ("FIND"):
                    Search();
                    break;

                //Works with no appointments - Haven't checked when appointment is present
                case ("CALENDER"):
                case ("CHECK CALENDER"):
                case ("APPOINTMENTS"):
                case ("TASKS"):
                    Console.WriteLine("Calender");
                    CheckCalender();
                    break;

                case ("GOOGLE"):
                    GoogleSearch();
                    break;

                case ("SHOW TIME"):
                case ("TIME"):
                case ("CURRENT TIME"):
                case ("TELL TIME"):
                case ("SAY TIME"):
                    CurrentTime();
                    break;

                case ("SHOW DAY"):
                case ("DAY"):
                case ("CURRENT DAY"):
                case ("TELL DAY"):
                case ("SAY DAY"):
                case ("SHOW DATE"):
                case ("DATE"):
                case ("CURRENT DATE"):
                case ("TELL DATE"):
                case ("SAY DATE"):
                    CurrentDate();
                    break;

                case ("HELLO"):
                case ("HEY JARVIS"):
                case ("HEY"):
                case ("SUP"):
                case ("GOOD MORNING"):
                case ("GOOD AFTERNOON"):
                case ("GOOD EVENING"):
                    HelloResponse();
                    break;

                case ("PLAY"):
                case ("PLAY ITUNES"):
                case ("PLAY SONG"):
                    //app = new iTunesApp();
                    //app.Play();
                    break;

                case ("PAUSE"):
                case ("PAUSE ITUNES"):
                case ("PAUSE SONG"):
                    //app = new iTunesApp();
                    //app.Pause();
                    break;

                case ("NEXT"):
                case ("NEXT SONG"):
                    //app = new iTunesApp();
                    //app.NextTrack();
                    //app.Play();
                    break;

                case ("PREVIOUS"):
                case ("PREVIOUS SONG"):
                case ("LAST SONG"):
                    //app = new iTunesApp();
                    //app.PreviousTrack();
                    //app.Play();
                    break;

                case ("SHUFFLE"):
                case ("SHUFFLE ITUNES"):
                    //app = new iTunesApp();
                    //app.CurrentPlaylist.Shuffle = true;
                    //app.Play();
                    break;

                case ("TURN DOWN VOLUME"):
                case ("TURN DOWN ITUNES"):
                case ("TURN DOWN SONG VOLUME"):
                case ("TURN DOWN THE VOLUME"):
                case ("TURN DOWN THE SONG VOLUME"):
                case ("TURN DOWN"):
                case ("VOLUME DOWN"):
                    TurnDownItunes();
                    break;

                case ("TURN UP VOLUME"):
                case ("TURN UP ITUNES"):
                case ("TURN UP SONG VOLUME"):
                case ("TURN UP THE VOLUME"):
                case ("TURN UP THE SONG VOLUME"):
                case ("TURN UP"):
                case ("VOLUME UP"):
                    TurnUpItunes();
                    break;

                case ("LEAGUE OF LEGENDS"):
                case ("LOL"):
                case ("LEAGUE"):
                    Process.Start(@"C:\Riot Games\League of Legends\lol.launcher.exe");
                    break;

                case ("LOLESPORTS"):
                case ("LOL ESPORTS"):
                case ("WATCH LEAGUE"):
                    Process.Start("http://euw.lolesports.com/");
                    Thread loadLeague = new Thread(new ThreadStart(() => speak.loading()));
                    loadLeague.IsBackground = true;
                    loadLeague.Start();
                    break;

                case ("JARVIS QUIET"):
                case ("JARVIS SH"):
                case ("JARVIS VOLUME DOWN"):
                    jarvisVolume(true);
                    break;

                case ("JARVIS LOUD"):
                case ("I CANT HEAR YOU"):
                case ("I CANT HEAR YOU JARVIS"):
                case ("JARVIS VOLUME UP"):
                    jarvisVolume(false);
                    break;

                case ("JARVIS MUTE"):
                case ("MUTE"):
                    JarvisMute(true);
                    break;

                case ("JARVIS SPEAK"):
                case ("UNDO MUTE"):
                    JarvisMute(false);
                    break;

                case ("STOP LISTENING"):
                    StopListening();
                    break;

                case ("EXIT CHROME"):
                case ("CLOSE CHROME"):
                    ExitChromeWindows();
                    break;

                case ("OPEN CHROME"):
                    OpenChromeWindow();
                    break;

                case ("ALL PROCESSES"):
                case("SHOW PROCESSES"):
                    ShowProcesses();
                    break;

                case ("CLOSE MAIL"):
                case ("EXIT MAIL"):
                case ("CLOSE OUTLOOK"):
                case ("EXIT OUTLOOK"):
                    CloseEmail();
                    break;

                case ("CLOSE ITUNES"):
                case ("EXIT ITUNES"):
                    EndProcess("itunes");
                    break;

                case ("MOVIES"):
                    LoadMovies();
                    break;

                case ("MINIMIZE"):
                case ("JARVIS SMALL"):
                    Minimize();
                    break;

                case ("JARVIS COME BACK"):
                case ("JARVIS NORMAL"):
                case ("NORMAL"):
                    NormalSize();
                    break;

                case ("QUIT"):
                case ("Q"):
                case ("STOP"):
                case ("END"):
                    EndProgram(this);
                    break;

                default:
                    NoCommand();
                    break;
            }
        }

        //..........................................Welcome/Goodbye/Hello Response messages................................

        public static void LoadUpWelcome()
        {
            wAndG = new WelcomeAndGoodbye();
            Thread welcomeThread = new Thread(new ThreadStart(wAndG.welcome));
            welcomeThread.IsBackground = true;
            welcomeThread.Start();
        }


        public static void EndProgram(Form thisForm)
        {
            wAndG = new WelcomeAndGoodbye();
            Thread goodbyeThread = new Thread(new ThreadStart(wAndG.goodbye));
            goodbyeThread.IsBackground = true;
            goodbyeThread.Start();
            goodbyeThread.Join();
            thisForm.Close();
        }

        public static void HelloResponse()
        {

            Thread hello = new Thread(new ThreadStart(speak.hello));
            hello.IsBackground = true;
            hello.Start();
        }

        //....................................................................................................

        //........................................Weather.....................................................
        public static void GetWeather()
        {
            //Get the weather conditions
            weather = new Weather();
            string[] foundConditions = weather.getWeather();
            ThreadStart weatherThreadStart = new ThreadStart(() => speak.sayWeather(foundConditions));
            Thread weatherThread = new Thread(weatherThreadStart);
            weatherThread.IsBackground = true;
            weatherThread.Start();
        }

        public static void GetForecast()
        {
            weather = new Weather();
            string[] foundConditions = weather.getWeather();
            Thread forecast = new Thread(new ThreadStart(() => speak.sayForecast(foundConditions)));
            forecast.IsBackground = true;
            forecast.Start();
        }

        //.....................................................................................................

        //......................................Email..........................................................
        public static void OpenMail()
        {
            email = new Email();
            email.openMail();
        }

        public static void SendMail()
        {
            email = new Email();
            email.sendMail();
        }

        private static void ReadMail()
        {
            email = new Email();
            email.readMail();
        }
        //.....................................................................................................

        //.....................................Search Files....................................................

        public static Search getSearchFile = new Search();
        public static SpeechRecognitionEngine whatToSearch = new SpeechRecognitionEngine();
        public string getFile = "No Data";

        public void Search()
        {
            getSearchFile.Show();
            myVoice.RecognizeAsyncCancel();
            whatToSearch.SetInputToDefaultAudioDevice();
            whatToSearch.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(GetPhrases()))));
            whatToSearch.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(WhatToSearch_SpeechRecognized);
            whatToSearch.RecognizeAsync(RecognizeMode.Multiple);
        }

        private void WhatToSearch_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text.ToUpper().Equals("SEARCH"))
            {
                Console.WriteLine("Getting Input");
                this.GetFileMethod = getSearchFile.inputFound();
                Console.WriteLine("Searching");
                try
                {
                    myVoice.RecognizeAsync(RecognizeMode.Multiple);
                }
                catch
                {
                }
                whatToSearch.RecognizeAsyncStop();
                FindFile(this.GetFileMethod);
            }
        }

        public string GetFileMethod
        {
            get
            {
                return getFile;
            }
            set
            {
                getFile = value;
            }
        }


        public static void FindFile(string filename)
        {
            getSearchFile.Close();
            string[] files = null;

            try
            {
                files = Directory.GetFiles("C:\\Users\\Alex\\Downloads", filename + ".*", SearchOption.AllDirectories);
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine(e.ToString());
            }

            if (files != null)
            {
                foreach (string file in files)
                {
                    Console.WriteLine(file);

                    try
                    {
                        ProcessStartInfo psi = new ProcessStartInfo(file);
                        Process proc = new Process();
                        proc.StartInfo = psi;
                        if (proc.Start())
                        {
                            Thread loading = new Thread(new ThreadStart(() => speak.loading()));
                            loading.IsBackground = true;
                            loading.Start();
                        }
                        else
                        {
                            Console.WriteLine("Error");
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Error");
                    }
                }
            }
            else
            {
                Console.WriteLine("Error");
            }
        }
        //.....................................................................................................

        //.....................................Check Calender..................................................
        public static void CheckCalender()
        {
            calender = new Calender();
            calender.CalenderAppointments();
        }
        //.....................................................................................................

        //........................................Google search................................................

        public static SpeechRecognitionEngine whatToGoogle = new SpeechRecognitionEngine();
        public static googleOther google = new googleOther();
        public static void GoogleSearch()
        {
            google.Show();

            Thread searchForThread = new Thread(new ThreadStart(() => speak.searchFor()));
            searchForThread.IsBackground = true;
            searchForThread.Start();

            myVoice.RecognizeAsyncCancel();

            whatToGoogle.SetInputToDefaultAudioDevice();
            whatToGoogle.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(GetPhrases()))));
            whatToGoogle.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(WhatToGoogle_SpeechRecognized);
            whatToGoogle.RecognizeAsync(RecognizeMode.Multiple);
        }

        private static void WhatToGoogle_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text.ToUpper().Equals("SEARCH"))
            {
                google.searchGoogle();

                Thread loading = new Thread(new ThreadStart(() => speak.loading()));
                loading.IsBackground = true;
                loading.Start();
                try
                {
                    myVoice.RecognizeAsync(RecognizeMode.Multiple);
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                whatToGoogle.RecognizeAsyncStop();
            }

        }
        //...................................................................................................

        //........................................Get Time and Date..........................................
        public static void CurrentTime()
        {

            speak.currentTime();
        }

        public static void CurrentDate()
        {

            speak.currentDate();
        }
        //..................................................................................................

        //....................................................Itunes Volume.................................

        public static void TurnUpItunes()
        {
            //app = new iTunesApp();
            volume += 20;
            //app.SoundVolume = volume;
        }

        public static void TurnDownItunes()
        {
            //app = new iTunesApp();
            volume -= 20;
            //app.SoundVolume = volume;
        }

        //..................................................................................................

        //....................................................Other.........................................             

        public static void jarvisVolume(bool volumeDown)
        {

            speak.jarvisVol(volumeDown);
        }

        public static void JarvisMute(bool mute)
        {

            speak.jarvisMute(mute);
        }

        public static string[] GetPhrases()
        {
            string[] phrases = File.ReadAllLines(@"C:\Users\Tester2\Desktop\TestingPurpose\VS2019-Testing\Jarvis-master\Jarvis\JARVIS\Grammar.txt"); //C:\Users\Tester2\Desktop\TestingPurpose\VS2019-Testing\Jarvis-master\Jarvis\JARVIS
            int index = 0;
            foreach (string phrase in phrases)
            {
                if (phrase == string.Empty)
                {
                    phrases[index] = "Empty";
                }
                index++;
            }
            return phrases;
        }

        int time = 60;
        public void StopListening()
        {
            time = 60;
            myVoice.RecognizeAsyncStop();
            Console.WriteLine("Not Listening");
            stopListeningTimer.Start();
        }

        private void Time_Tick(object sender, EventArgs e)
        {
            time -= 1;
            Console.WriteLine(time.ToString());
            if (time == 0)
            {
                myVoice.RecognizeAsync(RecognizeMode.Multiple);
                Console.WriteLine("You may speak");
                stopListeningTimer.Stop();
            }
        }
        //..................................................................................................

        private void Jarvis_Click(object sender, EventArgs e)
        {
            time = 1;
        }

        private void ExitChromeWindows()
        {
            EndProcess("chrome");
        }

        private void OpenChromeWindow()
        {
            Process.Start("chrome.exe");
        }

        private void NoCommand()
        {
            Thread noOptAvail = new Thread(new ThreadStart(() => speak.noOptionAvailable()));
            noOptAvail.IsBackground = true;
            noOptAvail.Start();
        }

        private void LoadMovies()
        {
            try
            {
                Process.Start(@"C:\Users\Alex\Documents\Visual Studio 2013\Projects\Movies\Movies\bin\Debug\Movies.exe");
            }
            catch (System.Exception exc)
            {
                Console.WriteLine(exc.ToString());
            }
        }

        private void ShowProcesses()
        {
            try
            {
                Process[] allProcesses = Process.GetProcesses();
                foreach (Process item in allProcesses)
                {
                    Console.WriteLine(item.ProcessName.ToString());
                }
            }
            catch (System.Exception exce)
            {
                Console.WriteLine(exce.ToString());
            }
        }

        private void EndProcess(string process)
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process currentProcess in processes)
            {
                if (currentProcess.ProcessName.ToString().ToUpper().Contains(process.ToUpper()))
                {
                    Console.WriteLine("Process: {0} ID: {1}", currentProcess.ProcessName, currentProcess.Id);
                    currentProcess.Kill();
                }
            }
        }

        private void CloseEmail()
        {
            EndProcess("outlook");
        }


        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void Jarvis_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void Minimize()
        {
            this.WindowState = FormWindowState.Minimized;
            JarvisApp.Icon = SystemIcons.Application;
            JarvisApp.BalloonTipTitle = "Jarvis";
            JarvisApp.BalloonTipText = "Running in background";
            JarvisApp.ShowBalloonTip(1500);
            JarvisApp.Visible = true;
        }
            
        private void NormalSize()
        {
            this.WindowState = FormWindowState.Normal;
            JarvisApp.Visible = false;
        }

        private void JarvisApp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            NormalSize();
            JarvisApp.Visible = false;
        }

        private void Jarvis_FormClosing(object sender, FormClosingEventArgs e)
        {
            speak.closing();
        }
    }
}
