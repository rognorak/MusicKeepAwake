using Microsoft.Win32;
using System;
using System.Drawing;
using System.Timers;
using System.Windows.Forms;
using CSCore.CoreAudioAPI;

namespace MusicKeepAwake
{
    class TrayApplicationContext : ApplicationContext
    {
        private const string IconOff = "dialog-block.ico";
        private const string IconOn = "dialog-apply.ico";
        private const string Tooltip = "Keep Computer From Sleeping";

        private System.ComponentModel.IContainer components;    // a list of components to dispose when the context is disposed

        [UsedImplicitly]
        private NotifyIcon _notifyIcon;
        private Icon _onIcon;
        private Icon _offIcon;

        private AboutBox _mainWindow;
        private System.Timers.Timer _timer;
        private bool _enabled;
        private ToolStripMenuItem _status;
        private string _statusData = "";

        public TrayApplicationContext()
        {
            _timer = new System.Timers.Timer(10000);
            _timer.Elapsed += OnTimedEvent;
            _timer.AutoReset = true;

            _onIcon = new Icon(IconOn);
            _offIcon = new Icon(IconOff);

            CheckAudioStatus();

            InitializeContext();
        }

        private void InitializeContext()
        {
            components = new System.ComponentModel.Container();

            _notifyIcon = new NotifyIcon(components)
            {
                ContextMenuStrip = new ContextMenuStrip(),
                Icon = _onIcon,
                Text = Tooltip,
                Visible = true
            };

            _mainWindow = new AboutBox();
            components.Add(_mainWindow);

            BuildMenu();
            _notifyIcon.MouseDown += _notifyIcon_MouseDown;

            //notifyIcon.ContextMenuStrip.Opening += ContextMenuStrip_Opening;
            //notifyIcon.DoubleClick += notifyIcon_DoubleClick;
            //notifyIcon.MouseUp += notifyIcon_MouseUp;

            _enabled = true;
            _timer.Start();
        }

        private void BuildMenu()
        {
            _notifyIcon.ContextMenuStrip.Opening += ContextMenuStripOpening;

            _notifyIcon.ContextMenuStrip.Items.Clear();
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            ToolStripMenuItem temp;

            temp = new ToolStripMenuItem(String.Format("Status: {0}", _statusData));
            _status = temp;
            _notifyIcon.ContextMenuStrip.Items.Add(temp);

            temp = new ToolStripMenuItem("&Start");
            temp.Click += StartClick;
            _notifyIcon.ContextMenuStrip.Items.Add(temp);

            temp = new ToolStripMenuItem("&Stop");
            temp.Click += StopClick;
            _notifyIcon.ContextMenuStrip.Items.Add(temp);

            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            temp = new ToolStripMenuItem("&Exit");
            temp.Click += exit_Click;
            _notifyIcon.ContextMenuStrip.Items.Add(temp);

        }

        private void ContextMenuStripOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CheckAudioStatus();
            _status.Text = String.Format("Status: {0}", _statusData);
        }

        private void exit()
        {
            Application.Exit();
        }

        void exit_Click(object sender, EventArgs e)
        {
            exit();
        }

        private void StartClick(object sender, EventArgs e)
        {
            if (_enabled)
                return;

            _notifyIcon.Icon = _onIcon;
            _timer.Start();
            _enabled = true;
        }

        private void StopClick(object sender, EventArgs e)
        {
            if (!_enabled)
                return;

            _notifyIcon.Icon = _offIcon;
            _timer.Stop();
            _status.Text = "Status:";
            _enabled = false;
        }

        void _notifyIcon_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (!_mainWindow.Visible)
                    _mainWindow.Show();

                if (_mainWindow.WindowState == FormWindowState.Minimized)
                    _mainWindow.WindowState = FormWindowState.Normal;

                _mainWindow.Activate();
            }
        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            CheckAudioStatus();
        }

        private void CheckAudioStatus()
        {

            if (!_enabled)
            {
                _statusData = "";
                return;
            }
                

            bool status = IsAudioPlaying(GetDefaultRenderDevice());
            _statusData = status.ToString();

            if (status)
            {
                NativeCalls.KeepAlive(false);
            }
        }

        private static MMDevice GetDefaultRenderDevice()
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                return enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
            }
        }

        private static bool IsAudioPlaying(MMDevice device)
        {
            using (var meter = AudioMeterInformation.FromDevice(device))
            {
                return meter.PeakValue > 0;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
        }

        protected override void ExitThreadCore()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Dispose();
            }

            base.ExitThreadCore();
        }

    }
}