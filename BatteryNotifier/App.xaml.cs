using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Resources;
using Windows.Devices.Power;
using Windows.System.Power;
using Microsoft.VisualBasic.Devices;
using NAudio.Wave;
using Application = System.Windows.Application;

namespace BatteryNotifier
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private StreamResourceInfo _ico_gray;
        private StreamResourceInfo _ico_yellow;
        private StreamResourceInfo _ico_green;

        private StreamResourceInfo _connectPower;

        private MainWindow _window;

        private WaveOutEvent _outputDevice;
        private WaveFileReader _waveFileReader;

        private Icon GetIconFromStream(Stream stream)
        {
            stream.Seek(0, SeekOrigin.Begin); // Streamの位置を先頭に戻す
            return new System.Drawing.Icon(stream);
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            _window = new MainWindow();
            _window.Closing += (o, args) =>
            {
                _window.Hide();
                args.Cancel = true;
            };

            _ico_gray = GetResourceStream(new Uri("Resources/icon.ico", UriKind.Relative));
            _ico_yellow = GetResourceStream(new Uri("Resources/icon_yellow.ico", UriKind.Relative));
            _ico_green = GetResourceStream(new Uri("Resources/icon_green.ico", UriKind.Relative));

            _connectPower = GetResourceStream(new Uri("Resources/connect_power.wav", UriKind.Relative));

            _outputDevice = new WaveOutEvent();
            _outputDevice.NumberOfBuffers = 40;
            _outputDevice.DesiredLatency = 200;

            var notifyIcon = new NotifyIcon
            {
                Icon = GetIconFromStream(_ico_gray.Stream),
                Visible = true,
                Text = "Battery Notifier"
            };
            notifyIcon.MouseDoubleClick += NotifyIcon_OnMouseDoubleClick;
            notifyIcon.ContextMenuStrip = new ContextMenuStrip();
            notifyIcon.ContextMenuStrip.Items.Add("Exit").Click += ExitMenuItem_OnClick;

            Task.Run(() =>
            {
                // watch battery status and changed to charging, play sound
                var battery = Battery.AggregateBattery;
                var lastStatus = battery.GetReport().Status;
                battery.ReportUpdated += async (o, args) =>
                {
                    var report = o.GetReport();


                    String status = report.Status switch
                    {
                        BatteryStatus.Idle => "Idle",
                        BatteryStatus.Charging => "Charging",
                        BatteryStatus.Discharging => "Discharging",
                        BatteryStatus.NotPresent => "Not Present",
                        _ => "Unknown"
                    };
                    notifyIcon.Text = $"Battery Notifier\n" +
                                      $"{status} {report.ChargeRateInMilliwatts / 1000}W\n" +
                                      $"{Math.Round((float)report.RemainingCapacityInMilliwattHours! / (float)report.FullChargeCapacityInMilliwattHours! * 100, 2)}%";
                    Console.WriteLine(
                        $"Battery status changed to {status} {report.ChargeRateInMilliwatts / 1000}W {report.RemainingCapacityInMilliwattHours} / {report.FullChargeCapacityInMilliwattHours} {(float)report.RemainingCapacityInMilliwattHours / (float)report.FullChargeCapacityInMilliwattHours * 100}%");


                    notifyIcon.Icon = report.Status switch
                    {
                        BatteryStatus.Idle => GetIconFromStream(_ico_green.Stream),
                        BatteryStatus.Charging => GetIconFromStream(_ico_yellow.Stream),
                        _ => GetIconFromStream(_ico_gray.Stream)
                    };

                    if (lastStatus == report.Status) return;
                    if (lastStatus == BatteryStatus.Charging ||
                        lastStatus == BatteryStatus.Idle)
                    {
                        lastStatus = report.Status;
                        return;
                    }

                    lastStatus = report.Status;

                    if (report.Status is BatteryStatus.Charging or BatteryStatus.Idle)
                    {
                        _connectPower.Stream.Position = 0;
                        if (_waveFileReader == null)
                        {
                            _waveFileReader = new WaveFileReader(_connectPower.Stream);
                        }

                        _outputDevice.Init(_waveFileReader);

                        _outputDevice.PlaybackStopped += (sender, eventArgs) =>
                        {
                            _outputDevice.Dispose();
                            _waveFileReader = null;
                        };
                        _outputDevice.Play();
                    }
                };
            });
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            Environment.Exit(0);
        }

        private void NotifyIcon_OnMouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (_window.IsVisible)
                _window.Activate();
            else
                _window.Show();
        }

        private void ExitMenuItem_OnClick(object sender, EventArgs e)
        {
            Shutdown();
        }
    }
}