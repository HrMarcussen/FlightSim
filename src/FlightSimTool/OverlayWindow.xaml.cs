using System;
using System.Windows;
using System.Windows.Threading;
using FlightSimTool.Core;

namespace FlightSimTool
{
    public partial class OverlayWindow : Window
    {
        private readonly IntPtr _targetHandle;
        private readonly LayoutEntry _config;
        private readonly DispatcherTimer _timer;

        public OverlayWindow(IntPtr targetHandle, LayoutEntry config)
        {
            InitializeComponent();
            _targetHandle = targetHandle;
            _config = config;

            UpdateThickness();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(50) };
            _timer.Tick += OnTick;
            
            this.Loaded += (s, e) => 
            {
                NativeMethods.SetClickThrough(new System.Windows.Interop.WindowInteropHelper(this).Handle);
                _timer.Start();
            };
        }

        private void UpdateThickness()
        {
            int t = Math.Max(1, _config.BorderThickness);
            int tx = Math.Max(0, _config.BorderTopExtra);
            
            BorderTop.Height = t + tx;
            BorderBottom.Height = t;
            BorderLeft.Width = t;
            BorderRight.Width = t;

            // Adjust margins so side borders don't overlap corner areas if we were using a single border
            Thickness sideMargin = new Thickness(0, t + tx, 0, t);
            BorderLeft.Margin = sideMargin;
            BorderRight.Margin = sideMargin;
        }

        private void OnTick(object? sender, EventArgs e)
        {
            // Verify target exists
            if (!NativeMethods.IsWindowVisible(_targetHandle))
            {
                this.Close();
                return;
            }

            NativeMethods.GetWindowRect(_targetHandle, out var r);
            if (r.Width <= 0 || r.Height <= 0) return;

            // Calculate overlay bounds
            // We want the overlay borders to INTRUDE into the window by 'Cover' amount
            // But usually this means placing the overlay EXACTLY over the window
            // and relying on the transparent center.
            
            // XAML Window positioning
            // Ideally matches target rect
            this.Left = r.Left;
            this.Top = r.Top;
            this.Width = r.Width;
            this.Height = r.Height;
        }

        public void Stop()
        {
            _timer.Stop();
            this.Close();
        }
    }
}
