using System;
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Forms = System.Windows.Forms;

namespace PlayFlowMIDI
{
    public partial class NotificationWindow : Window
    {
        private static NotificationWindow? _current;

        private const int MarginFromEdge = 20;
        private const double DisplayDuration = 3.0;
        private const double FadeInDuration = 0.25;
        private const double FadeOutDuration = 0.35;

        private DispatcherTimer? _hideTimer;

        public NotificationWindow()
        {
            InitializeComponent();
            this.Opacity = 0;
        }

        /// <summary>Show a song notification.</summary>
        public static void ShowSong(string title, string duration, string position, IntPtr targetHwnd)
        {
            Show("Now Playing", $"{title}\n{duration}", position, targetHwnd);
        }

        /// <summary>Show a playback-mode notification.</summary>
        public static void ShowMode(string modeName, string position, IntPtr targetHwnd)
        {
            Show("Playback Mode", modeName, position, targetHwnd);
        }

        private static void Show(string line1, string line2, string position, IntPtr targetHwnd)
        {
            _current?.CloseNow();

            var win = new NotificationWindow();
            _current = win;

            win.LineOne.Text = line1;
            win.LineTwo.Text = line2;
            win.LineTwo.Visibility = string.IsNullOrEmpty(line2) ? Visibility.Collapsed : Visibility.Visible;

            win.ContentRendered += (s, e) => win.ApplyPosition(position, targetHwnd);

            win.Show();
            win.FadeIn();
        }

        private void ApplyPosition(string position, IntPtr targetHwnd)
        {
            Forms.Screen screen;
            try
            {
                screen = targetHwnd != IntPtr.Zero
                    ? Forms.Screen.FromHandle(targetHwnd)
                    : Forms.Screen.PrimaryScreen ?? Forms.Screen.AllScreens[0];
            }
            catch
            {
                screen = Forms.Screen.PrimaryScreen ?? Forms.Screen.AllScreens[0];
            }

            var area = screen.WorkingArea;
            double w = this.ActualWidth;
            double h = this.ActualHeight;

            double dpiScaleX = 1.0;
            double dpiScaleY = 1.0;
            try
            {
                var source = PresentationSource.FromVisual(this);
                if (source?.CompositionTarget != null)
                {
                    dpiScaleX = source.CompositionTarget.TransformToDevice.M11;
                    dpiScaleY = source.CompositionTarget.TransformToDevice.M22;
                }
            }
            catch { }

            double left, top;
            double areaLeft   = area.Left   / dpiScaleX;
            double areaTop    = area.Top    / dpiScaleY;
            double areaRight  = area.Right  / dpiScaleX;
            double areaBottom = area.Bottom / dpiScaleY;

            switch (position)
            {
                case "TopLeft":
                    left = areaLeft   + MarginFromEdge;
                    top  = areaTop    + MarginFromEdge;
                    break;
                case "TopRight":
                    left = areaRight  - w - MarginFromEdge;
                    top  = areaTop    + MarginFromEdge;
                    break;
                case "BottomLeft":
                    left = areaLeft   + MarginFromEdge;
                    top  = areaBottom - h - MarginFromEdge;
                    break;
                case "BottomRight":
                default:
                    left = areaRight  - w - MarginFromEdge;
                    top  = areaBottom - h - MarginFromEdge;
                    break;
            }

            this.Left = left;
            this.Top  = top;
        }

        private void FadeIn()
        {
            var anim = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(FadeInDuration))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            anim.Completed += (s, e) => ScheduleHide();
            this.BeginAnimation(OpacityProperty, anim);
        }

        private void ScheduleHide()
        {
            _hideTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(DisplayDuration) };
            _hideTimer.Tick += (s, e) =>
            {
                _hideTimer.Stop();
                FadeOut();
            };
            _hideTimer.Start();
        }

        private void FadeOut()
        {
            var anim = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(FadeOutDuration))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
            };
            anim.Completed += (s, e) => CloseNow();
            this.BeginAnimation(OpacityProperty, anim);
        }

        private void CloseNow()
        {
            _hideTimer?.Stop();
            if (_current == this) _current = null;
            try { this.Close(); } catch { }
        }
    }
}
