using CoreAudioApi;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Fix_My_Microphone_
{
    /// <summary>
    /// Proctorio Inc, Open Source Initiative 2015 https://proctorio.com
    /// 
    /// Mute/UnMute originally built by Matt Palmerlee November 2010 
    /// For muting and unmuting the microphone using C# on Windows XP, Vista, and Windows 7
    /// Uses parts of Gustavo Franco's MixerNative AudioLib source for Windows XP and older from here:
    /// http://www.codeguru.com/csharp/csharp/cs_graphics/sound/article.php/c10931
    /// And uses Ray Molenkamp's C# managed wrapper for accessing the Vista Core Audio API (for Windows Vista and newer)
    /// http://www.codeproject.com/KB/vista/CoreAudio.aspx?msg=2489276
    /// Other references:
    /// http://stackoverflow.com/questions/2078970/how-to-mute-the-microphone-c
    /// http://stackoverflow.com/questions/154089/mute-windows-volume-using-c
    /// http://stackoverflow.com/questions/3046668/how-to-mute-microphone-in-windows-7-with-c-c
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // remove the window icon, clean is key
        protected override void OnSourceInitialized(EventArgs e) { IconHelper.RemoveIcon(this); }

        // delegates for ui update outside of render thread
        public delegate void UpdateProgressCallback(double v);
        public delegate void ShowCheckCallback();
        public delegate void ShowXCallback();
        public delegate void ShowChromeCallback();
        public delegate void HideChromeCallback();
        public delegate void HideCheckCallback();
        public delegate void HideMicCallback();
        public delegate void UpdateTitleCallback(string m);

        // thread to animate and unmute microphones
        private void UIThread()
        {
            // hang tight
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("hang tight") });

            // sleep a bit
            Thread.Sleep(500);

            // find me some microphones
            P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { ("finding microphones to fix...") });

            // placeholder for device setting set
            bool setDevice = false;

            // first try the new way, otherwise fallback in the catch
            try
            {
                // get the devices connected
                MMDeviceEnumerator DevEnum = new MMDeviceEnumerator();
                MMDeviceCollection devices = DevEnum.EnumerateAudioEndPoints(EDataFlow.eCapture, EDeviceState.DEVICE_STATEMASK_ALL);

                // show how many devices we found
                P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { string.Format("found {0} possible devices", devices.Count) });

                // holder for progress spinner
                int t = 0;

                // itterate over devices
                for (int i = 0; i < devices.Count; i++)
                {
                    // itterate over progressbar
                    for (int j = t; j <= 100; j++)
                    {
                        t = j;
                        double d = (((double)(i + 1) / devices.Count) * 100);
                        if (d <= j)
                            break;

                        Thread.Sleep(35);
                        P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { j });
                    }

                    // dont spin too fast
                    Thread.Sleep(1000);

                    // extract device data
                    MMDevice deviceAt = devices[i];
                    string low_name = deviceAt.FriendlyName.ToLower();

                    // skip not present devices
                    if (deviceAt.State == EDeviceState.DEVICE_STATE_NOTPRESENT)
                    {
                        // not present
                        P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { string.Format("skipping {0}, device not present", low_name) });
                        continue;
                    }

                    // skip not plugged in devices
                    if (deviceAt.State == EDeviceState.DEVICE_STATE_UNPLUGGED)
                    {
                        // not plugged in
                        P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { string.Format("skipping {0}, device unplugged", low_name) });
                        continue;
                    }

                    // try to unmute and set volume on this device
                    try
                    {
                        deviceAt.AudioEndpointVolume.Mute = false;
                        deviceAt.AudioEndpointVolume.MasterVolumeLevelScalar = 1;
                        P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { string.Format("{0} : unmute, volume (100%)", low_name) });

                        // mark as passed this section if name is microphone
                        if (low_name.Contains("microphone")) setDevice = true;
                    }
                    catch { }
                }

                // did we even find any devices?
                if (devices.Count == 0)
                {
                    // failure, can't continue
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "no microphones found" });

                    // reset progressbar
                    P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });

                    // hide the microphone icon
                    Mic.Dispatcher.Invoke(new HideMicCallback(this.HideMic));

                    // show failure X
                    X.Dispatcher.Invoke(new ShowXCallback(this.ShowX));
                    return;
                }
            }
            catch
            {
                // fallback option
                MixerNativeLibrary.MicInterface.MuteOrUnMuteAllMics(false);

                // i dunno, always set this to true
                setDevice = true;
            }

            // hide the microphone icon
            Mic.Dispatcher.Invoke(new HideMicCallback(this.HideMic));

            // did we do some good?
            if (!setDevice)
            {
                // failure, can't continue
                P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "all valid microphones unplugged or disabled" });

                // reset progressbar
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });

                // show failure X
                X.Dispatcher.Invoke(new ShowXCallback(this.ShowX));
            }
            else
            {
                // finsh out the progress bar
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });

                // show the checkmark
                CheckMark.Dispatcher.Invoke(new ShowCheckCallback(this.ShowCheck));

                // done
                P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "done with microphone(s)" });

                // Zzzz
                Thread.Sleep(2000);

                // reset progressbar
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 0 });

                // hide the check mark
                CheckMark.Dispatcher.Invoke(new HideCheckCallback(this.HideCheck));

                // show the pulsing chrome icon
                CheckMark.Dispatcher.Invoke(new ShowChromeCallback(this.ShowChrome));

                // figure out how many chrome processes are open and running
                Process[] chromeInstances = Process.GetProcessesByName("chrome");
                int total = chromeInstances.Length;

                // case where no chrome windows open
                if (total <= 0)
                {
                    // indicate chrome restart
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "opening chrome..." });

                    // open chrome
                    Process.Start(@"chrome.exe");
                }
                else
                {
                    // indicate chrome restart
                    P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "restarting chrome..." });

                    // restart all instances of chrome, wait for them to all close
                    Process.Start(@"chrome.exe", "http://restart.chrome/");

                    // stopwatch for give up plan
                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    while (true)
                    {
                        chromeInstances = Process.GetProcessesByName("chrome");

                        // wait til we reach 2 or less chrome instances
                        // also give up after 45 seconds
                        if (chromeInstances.Length <= 2 || sw.Elapsed.TotalSeconds > 45)
                        {
                            // done
                            break;
                        }
                        else
                        {
                            // make sure the "progress" donut doesnt show less progress over time
                            total = Math.Max(total, chromeInstances.Length);

                            // update th eprogress bar
                            P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { Math.Ceiling((((double)total - (double)chromeInstances.Length) / (double)total) * 100) });
                        }

                        // dont spin the cpu
                        Thread.Sleep(100);
                    }
                }

                // set to 100% for visual clue
                P.Dispatcher.Invoke(new UpdateProgressCallback(this.UpdateProgress), new object[] { 100 });

                // hide the pulsing chrome icon
                CheckMark.Dispatcher.Invoke(new HideChromeCallback(this.HideChrome));

                // show the check mark
                CheckMark.Dispatcher.Invoke(new ShowCheckCallback(this.ShowCheck));

                // set done and good luck messaging
                P.Dispatcher.Invoke(new UpdateTitleCallback(this.UpdateTitle), new object[] { "done, good luck on your exam!" });

                // let them read it and wait
                Thread.Sleep(5000);

                // kill this app
                Environment.Exit(0);
            }
        }

        // update the progressbar
        private void UpdateProgress(double v) { P.Value = v; }

        // show the checkmark
        private void ShowCheck() { CheckMark.Visibility = Visibility.Visible; }

        // show the x
        private void ShowX() { X.Visibility = Visibility.Visible; }

        // show the chrome icon
        private void ShowChrome() { Chrome.Visibility = Visibility.Visible; }

        // hide the chrome icon
        private void HideChrome() { Chrome.Visibility = Visibility.Hidden; }

        // hide the checkmark
        private void HideCheck() { CheckMark.Visibility = Visibility.Hidden; }

        // hide the mic icon
        private void HideMic() { Mic.Visibility = Visibility.Hidden; }

        // title change for progress of events
        private void UpdateTitle(string m) { this.Title = m; }

        // load it up
        private void OnWindowLoaded(object sender, RoutedEventArgs e) { Thread UI = new Thread(new ThreadStart(UIThread)); UI.Start(); }
    }
}
