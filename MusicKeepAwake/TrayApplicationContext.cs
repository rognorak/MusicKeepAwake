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
        private const string Tooltip = "Prevent Sleep when music is playing";
        private const string Message = "Got Music? : ";

        private System.ComponentModel.IContainer components;    // a list of components to dispose when the context is disposed

        [UsedImplicitly]
        private NotifyIcon _notifyIcon;
        private Icon _onIcon;
        private Icon _offIcon;

        private AboutBox _mainWindow;
        private System.Timers.Timer _timer;
        private bool _enabled;
        
        private ToolStripMenuItem _status;
        private ToolStripMenuItem _on;
        private ToolStripMenuItem _off;

        private string _statusData = "";

        public TrayApplicationContext()
        {
            _timer = new System.Timers.Timer(10000);
            _timer.Elapsed += OnTimedEvent;
            _timer.AutoReset = true;

            _onIcon = new Icon(IconOn);
            _offIcon = new Icon(IconOff);

            CheckAudio();

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
            SetChecks();
            _timer.Start();
        }

        private void BuildMenu()
        {
            _notifyIcon.ContextMenuStrip.Opening += ContextMenuStripOpening;

            _notifyIcon.ContextMenuStrip.Items.Clear();
            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());
            ToolStripMenuItem temp;

            _status = new ToolStripMenuItem();
            SetStatus();
            _notifyIcon.ContextMenuStrip.Items.Add(_status);

            _on = new ToolStripMenuItem("&On");
            _on.Click += StartClick;
            _notifyIcon.ContextMenuStrip.Items.Add(_on);

            _off = new ToolStripMenuItem("&Off");
            _off.Click += StopClick;
            _notifyIcon.ContextMenuStrip.Items.Add(_off);

            _notifyIcon.ContextMenuStrip.Items.Add(new ToolStripSeparator());

            temp = new ToolStripMenuItem("&Exit");
            temp.Click += exit_Click;
            _notifyIcon.ContextMenuStrip.Items.Add(temp);

        }

        private void SetStatus()
        {
            _status.Text = String.Format("{0}{1}", Message, _statusData);
        }

        private void ContextMenuStripOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CheckAudio();
            SetStatus();
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
            SetChecks();
        }

        private void StopClick(object sender, EventArgs e)
        {
            if (!_enabled)
                return;

            _notifyIcon.Icon = _offIcon;
            _timer.Stop();
            _status.Text = Message;
            _enabled = false;
            SetChecks();
        }

        private void SetChecks()
        {
            _on.Checked = _enabled;
            _off.Checked = !_enabled;
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
            CheckAudio();
        }

        private void CheckAudio()
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