using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace OutlookNotifier
{
    // Borderless, non-activating acrylic card: next meeting of the day above the latest unread
    // message. Stays until clicked or closed. In edit mode (settings open) it shows sample content
    // and can be dragged and resized.
    class NotificationWindow : Window
    {
        public event Action OpenMailRequested;
        public event Action OpenCalendarRequested;
        public event Action Edited;
        public event Action Hidden;

        static readonly Color[] AvatarColors =
        {
            Color.FromRgb(0x4F, 0x6B, 0xED), Color.FromRgb(0x1F, 0x8A, 0x7E), Color.FromRgb(0xB0, 0x52, 0x7A),
            Color.FromRgb(0xC2, 0x6E, 0x2A), Color.FromRgb(0x5E, 0x6E, 0x80), Color.FromRgb(0x82, 0x5C, 0xCF),
            Color.FromRgb(0x2D, 0x7F, 0xC1), Color.FromRgb(0x5A, 0x8A, 0x3C)
        };

        const double ScreenGap = 12;
        const double SnapDistance = 28;

        readonly Border host;
        readonly ScaleTransform zoom = new ScaleTransform(1, 1);
        readonly TranslateTransform slide = new TranslateTransform();
        readonly Grid overlay;
        readonly Thumb widthGrip, sizeGrip;
        readonly Border unreadBadge, closeButton, meetingRow, meetingMore, divider, mailRow, avatar;
        readonly TextBlock unreadText, meetingWhen, meetingTitle, meetingDetail, meetingMoreText,
                           mailTime, mailSender, mailSubject, mailPreview, avatarText, emptyText;

        Settings settings = new Settings();
        OutlookState state = new OutlookState();
        IntPtr hwnd;
        bool hiding;
        bool editMode;
        bool visibleBeforeEdit;
        Point dragCursor;
        double dragWidth, dragScale;
        readonly DispatcherTimer watchOthers = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
        readonly DispatcherTimer glide = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(15) };
        Point glideFrom, glideTo;
        DateTime glideStart;

        public bool IsShown { get { return IsVisible && !hiding; } }
        public bool EditMode { get { return editMode; } }

        public NotificationWindow()
        {
            Title = "Outlook Notifier";
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            ShowActivated = false;
            Topmost = true;
            SizeToContent = SizeToContent.Height;
            Background = Brushes.Transparent;
            UseLayoutRounding = true;

            var view = Ui.Load("Popup.xaml");
            Ui.Find<Image>(view, "Logo").Source = MailIcon.Create();
            unreadBadge = Ui.Find<Border>(view, "UnreadBadge");
            unreadText = Ui.Find<TextBlock>(view, "UnreadText");
            closeButton = Ui.Find<Border>(view, "CloseButton");
            meetingRow = Ui.Find<Border>(view, "MeetingRow");
            meetingWhen = Ui.Find<TextBlock>(view, "MeetingWhen");
            meetingTitle = Ui.Find<TextBlock>(view, "MeetingTitle");
            meetingDetail = Ui.Find<TextBlock>(view, "MeetingDetail");
            meetingMore = Ui.Find<Border>(view, "MeetingMore");
            meetingMoreText = Ui.Find<TextBlock>(view, "MeetingMoreText");
            divider = Ui.Find<Border>(view, "Divider");
            mailRow = Ui.Find<Border>(view, "MailRow");
            avatar = Ui.Find<Border>(view, "Avatar");
            avatarText = Ui.Find<TextBlock>(view, "AvatarText");
            mailTime = Ui.Find<TextBlock>(view, "MailTime");
            mailSender = Ui.Find<TextBlock>(view, "MailSender");
            mailSubject = Ui.Find<TextBlock>(view, "MailSubject");
            mailPreview = Ui.Find<TextBlock>(view, "MailPreview");
            emptyText = Ui.Find<TextBlock>(view, "EmptyText");

            host = new Border { Child = view, LayoutTransform = zoom, RenderTransform = slide, Background = Brushes.Transparent };
            host.MouseLeftButtonDown += OnHostMouseDown;

            WireRow(meetingRow, () => Raise(OpenCalendarRequested));
            WireRow(mailRow, () => Raise(OpenMailRequested));
            closeButton.MouseEnter += delegate { closeButton.SetResourceReference(Border.BackgroundProperty, "ControlHover"); };
            closeButton.MouseLeave += delegate { closeButton.Background = Brushes.Transparent; };
            closeButton.MouseLeftButtonDown += (s, e) => e.Handled = true;
            closeButton.MouseLeftButtonUp += (s, e) => { e.Handled = true; Dismiss(); };

            // Edit-mode overlay: dashed outline, width handle on the free edge, size handle in the free corner.
            var outline = new Rectangle
            {
                StrokeThickness = 1.5, StrokeDashArray = new DoubleCollection { 4, 3 },
                RadiusX = 8, RadiusY = 8, Margin = new Thickness(1), IsHitTestVisible = false
            };
            outline.SetResourceReference(Shape.StrokeProperty, "Accent");
            widthGrip = new Thumb
            {
                Width = 12, Height = 40, Cursor = Cursors.SizeWE, VerticalAlignment = VerticalAlignment.Center, ToolTip = "Largeur",
                Template = Ui.Template("Thumb", "<Border Background=\"Transparent\"><Border Width=\"4\" Height=\"28\" CornerRadius=\"2\" Background=\"{DynamicResource Accent}\"/></Border>")
            };
            sizeGrip = new Thumb
            {
                Width = 14, Height = 14, Margin = new Thickness(3), ToolTip = "Taille",
                Template = Ui.Template("Thumb", "<Ellipse Fill=\"{DynamicResource Accent}\" Stroke=\"White\" StrokeThickness=\"2\"/>")
            };
            widthGrip.DragStarted += delegate { dragCursor = CursorDip(); dragWidth = settings.Width; };
            widthGrip.DragDelta += delegate
            {
                settings.Width = Settings.Clamp(dragWidth + InwardDelta() / settings.Scale, 260, 600);
                Resize();
            };
            sizeGrip.DragStarted += delegate { dragCursor = CursorDip(); dragScale = settings.Scale; dragWidth = settings.Width * settings.Scale; };
            sizeGrip.DragDelta += delegate
            {
                settings.Scale = Settings.Clamp(dragScale * (dragWidth + InwardDelta()) / dragWidth, 0.7, 1.6);
                Resize();
            };
            overlay = new Grid { Visibility = Visibility.Collapsed };
            overlay.Children.Add(outline);
            overlay.Children.Add(widthGrip);
            overlay.Children.Add(sizeGrip);

            var root = new Grid();
            root.Children.Add(host);
            root.Children.Add(overlay);
            Content = root;

            SizeChanged += delegate { Place(ActualHeight); };
            glide.Tick += OnGlide;
            watchOthers.Tick += delegate { if (IsShown && !editMode && settings.AvoidOthers) Place(ActualHeight); };
            SourceInitialized += OnSourceInitialized;
            ApplySettings(settings);
        }

        void OnSourceInitialized(object s, EventArgs e)
        {
            hwnd = new WindowInteropHelper(this).Handle;
            HwndSource.FromHwnd(hwnd).CompositionTarget.BackgroundColor = Colors.Transparent;
            int ex = Native.GetWindowLong(hwnd, Native.GWL_EXSTYLE);
            Native.SetWindowLong(hwnd, Native.GWL_EXSTYLE, ex | Native.WS_EX_TOOLWINDOW | Native.WS_EX_NOACTIVATE);
            ApplyBackdrop();
        }

        public void ApplySettings(Settings s)
        {
            settings = s;
            Theme.From(s).ApplyTo(Resources);
            zoom.ScaleX = zoom.ScaleY = s.Scale;
            Width = Math.Round(s.Width * s.Scale);
            ApplyBackdrop();
            Render();
            UpdateGrips();
            if (IsVisible) Place(MeasureHeight());
        }

        void ApplyBackdrop()
        {
            if (hwnd == IntPtr.Zero) return;
            var theme = Theme.From(settings);
            Native.ApplyBackdrop(hwnd, theme.Dark, theme.Tint, theme.Border, settings.Blur, settings.Rounded);
        }

        public void ShowState(OutlookState s)
        {
            state = s ?? new OutlookState();
            Render();
            Present();
        }

        public void UpdateState(OutlookState s)
        {
            state = s ?? new OutlookState();
            if (IsShown) Render();
        }

        public void EnterEditMode()
        {
            visibleBeforeEdit = IsShown;
            editMode = true;
            overlay.Visibility = Visibility.Visible;
            closeButton.Visibility = Visibility.Collapsed;
            host.Cursor = Cursors.SizeAll;
            Render();
            Present();
        }

        public void ExitEditMode()
        {
            editMode = false;
            overlay.Visibility = Visibility.Collapsed;
            closeButton.Visibility = Visibility.Visible;
            host.Cursor = null;
            if (visibleBeforeEdit && state.LatestUnread != null) Render();
            else Dismiss();
        }

        void Render()
        {
            OutlookState s = editMode ? Demo.Fill(state) : state;

            bool showMeeting = settings.ShowMeetings && s.NextMeeting != null;
            meetingRow.Visibility = Visible(showMeeting);
            divider.Visibility = Visible(showMeeting);
            if (showMeeting)
            {
                MeetingInfo m = s.NextMeeting;
                meetingTitle.Text = m.Subject;
                meetingWhen.Text = When(m);
                meetingDetail.Text = m.Start.ToString("HH:mm") + " – " + m.End.ToString("HH:mm") +
                                     (string.IsNullOrWhiteSpace(m.Location) ? "" : " · " + m.Location);
                meetingMore.Visibility = Visible(s.OtherMeetings > 0);
                meetingMoreText.Text = s.OtherMeetings == 1 ? "1 autre réunion aujourd'hui"
                                     : s.OtherMeetings + " autres réunions aujourd'hui";
            }

            MailInfo mail = s.LatestUnread;
            mailRow.Visibility = Visible(mail != null);
            emptyText.Visibility = Visible(mail == null);
            emptyText.Text = s.Connected ? "Aucun message non lu" : "Outlook n'est pas ouvert";
            if (mail != null)
            {
                mailSender.Text = mail.Sender ?? "Expéditeur inconnu";
                mailSubject.Text = mail.Subject;
                mailPreview.Text = mail.Preview;
                mailPreview.Visibility = Visible(settings.ShowPreview && !string.IsNullOrEmpty(mail.Preview));
                mailTime.Text = mail.Received.Date == DateTime.Today ? mail.Received.ToString("HH:mm") : mail.Received.ToString("dd/MM");
                avatarText.Text = Initials(mail.Sender);
                avatar.Background = new SolidColorBrush(AvatarColors[(int)(StableHash(mail.Sender) % (uint)AvatarColors.Length)]);
            }

            unreadBadge.Visibility = Visible(s.Unread > 0);
            unreadText.Text = s.Unread > 99 ? "99+" : s.Unread.ToString();
            unreadBadge.ToolTip = s.Unread == 1 ? "1 message non lu" : s.Unread + " messages non lus";
        }

        static string When(MeetingInfo m)
        {
            DateTime now = DateTime.Now;
            if (m.Start <= now) return "En cours";
            int minutes = (int)Math.Ceiling((m.Start - now).TotalMinutes);
            if (minutes < 60) return "Dans " + minutes + " min";
            int hours = minutes / 60, rest = minutes % 60;
            return "Dans " + hours + " h" + (rest > 0 ? " " + rest.ToString("00") : "");
        }

        void Present()
        {
            bool appearing = !IsShown;
            hiding = false;
            if (appearing) glide.Stop();
            Place(MeasureHeight());
            if (!IsVisible) Show();
            watchOthers.Start();
            Topmost = false;
            Topmost = true;

            if (appearing)
            {
                var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
                host.BeginAnimation(OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(220)) { EasingFunction = ease });
                slide.BeginAnimation(TranslateTransform.YProperty, new DoubleAnimation(10, 0, TimeSpan.FromMilliseconds(320)) { EasingFunction = ease });
            }
        }

        public void Dismiss()
        {
            if (!IsVisible || hiding || editMode) return;
            hiding = true;
            var ease = new CubicEase { EasingMode = EasingMode.EaseIn };
            var fade = new DoubleAnimation(0, TimeSpan.FromMilliseconds(150)) { EasingFunction = ease };
            fade.Completed += delegate
            {
                if (!hiding) return;
                hiding = false;
                watchOthers.Stop();
                glide.Stop();
                Hide();
                Raise(Hidden);
            };
            host.BeginAnimation(OpacityProperty, fade);
            slide.BeginAnimation(TranslateTransform.YProperty, new DoubleAnimation(6, TimeSpan.FromMilliseconds(150)) { EasingFunction = ease });
        }

        double MeasureHeight()
        {
            var root = (FrameworkElement)Content;
            root.Measure(new Size(Width, double.PositiveInfinity));
            return root.DesiredSize.Height;
        }

        void Place(double height)
        {
            Rect area = SystemParameters.WorkArea;
            double w = Width, h = height;
            double left, top;
            if (double.IsNaN(w) || h <= 0) return;
            switch (settings.Anchor)
            {
                case "TopLeft": left = area.Left + ScreenGap; top = area.Top + ScreenGap; break;
                case "TopRight": left = area.Right - w - ScreenGap; top = area.Top + ScreenGap; break;
                case "BottomLeft": left = area.Left + ScreenGap; top = area.Bottom - h - ScreenGap; break;
                case "Custom": left = settings.CustomX; top = settings.CustomFromTop ? settings.CustomY : settings.CustomY - h; break;
                default: left = area.Right - w - ScreenGap; top = area.Bottom - h - ScreenGap; break;
            }
            double vl = SystemParameters.VirtualScreenLeft, vt = SystemParameters.VirtualScreenTop;
            left = Math.Max(vl, Math.Min(left, vl + SystemParameters.VirtualScreenWidth - w));
            top = Math.Max(vt, Math.Min(top, vt + SystemParameters.VirtualScreenHeight - h));

            if (settings.AvoidOthers && !editMode)
            {
                Point free = StepAside(new Rect(left, top, w, h));
                left = free.X;
                top = free.Y;
            }
            MoveTo(left, top, IsShown && !editMode);
        }

        // Moves the card off other notifications: stacked above them (below them for top
        // positions), or beside them when there is no room vertically.
        Point StepAside(Rect home)
        {
            Rect area = SystemParameters.WorkArea;
            double scale = Dpi.Scale;
            var others = OtherNotifications.Find()
                .Select(r => new Rect(r.X / scale, r.Y / scale, r.Width / scale, r.Height / scale))
                .ToList();

            Rect card = home;
            for (int pass = 0; pass < 10; pass++)
            {
                Rect hit = Rect.Empty;
                foreach (Rect o in others)
                {
                    Rect zone = o;
                    zone.Inflate(ScreenGap / 2, ScreenGap / 2);
                    if (zone.IntersectsWith(card)) { hit = o; break; }
                }
                if (hit.IsEmpty) return card.TopLeft;

                card.Y = GrowsDown ? hit.Bottom + ScreenGap : hit.Top - ScreenGap - card.Height;
                if (card.Top < area.Top || card.Bottom > area.Bottom)
                {
                    card = home;
                    bool goLeft = home.Left + home.Width / 2 >= hit.Left + hit.Width / 2;
                    card.X = goLeft ? hit.Left - ScreenGap - card.Width : hit.Right + ScreenGap;
                    return card.Left >= area.Left && card.Right <= area.Right ? card.TopLeft : home.TopLeft;
                }
            }
            return card.TopLeft;
        }

        void MoveTo(double left, double top, bool animate)
        {
            if (!animate || (Math.Abs(left - Left) < 1 && Math.Abs(top - Top) < 1))
            {
                glide.Stop();
                if (animate) return;
                Left = left;
                Top = top;
                return;
            }
            if (glide.IsEnabled && Math.Abs(glideTo.X - left) < 1 && Math.Abs(glideTo.Y - top) < 1) return;
            glideFrom = new Point(Left, Top);
            glideTo = new Point(left, top);
            glideStart = DateTime.UtcNow;
            glide.Start();
        }

        void OnGlide(object sender, EventArgs e)
        {
            double t = Math.Min(1, (DateTime.UtcNow - glideStart).TotalMilliseconds / 240);
            double eased = 1 - Math.Pow(1 - t, 3);
            Left = glideFrom.X + (glideTo.X - glideFrom.X) * eased;
            Top = glideFrom.Y + (glideTo.Y - glideFrom.Y) * eased;
            if (t >= 1) glide.Stop();
        }

        bool GrowsLeft { get { return settings.Anchor == "TopRight" || settings.Anchor == "BottomRight"; } }
        bool GrowsDown { get { return settings.Anchor == "TopLeft" || settings.Anchor == "TopRight" || (settings.Anchor == "Custom" && settings.CustomFromTop); } }

        void UpdateGrips()
        {
            var side = GrowsLeft ? HorizontalAlignment.Left : HorizontalAlignment.Right;
            widthGrip.HorizontalAlignment = side;
            sizeGrip.HorizontalAlignment = side;
            sizeGrip.VerticalAlignment = GrowsDown ? VerticalAlignment.Bottom : VerticalAlignment.Top;
            sizeGrip.Cursor = GrowsLeft == GrowsDown ? Cursors.SizeNESW : Cursors.SizeNWSE;
        }

        void Resize()
        {
            zoom.ScaleX = zoom.ScaleY = settings.Scale;
            Width = Math.Round(settings.Width * settings.Scale);
            Raise(Edited);
        }

        double InwardDelta()
        {
            double dx = CursorDip().X - dragCursor.X;
            return GrowsLeft ? -dx : dx;
        }

        Point CursorDip()
        {
            var source = PresentationSource.FromVisual(this);
            Point p = Native.CursorPixels();
            return source == null ? p : source.CompositionTarget.TransformFromDevice.Transform(p);
        }

        void OnHostMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!editMode || e.ChangedButton != MouseButton.Left) return;
            try { DragMove(); } catch (InvalidOperationException) { return; }

            // Snap to a corner when dropped close to one, otherwise keep the free position.
            Rect area = SystemParameters.WorkArea;
            double w = ActualWidth, h = ActualHeight;
            bool left = Math.Abs(Left - (area.Left + ScreenGap)) < SnapDistance;
            bool right = Math.Abs(Left + w - (area.Right - ScreenGap)) < SnapDistance;
            bool top = Math.Abs(Top - (area.Top + ScreenGap)) < SnapDistance;
            bool bottom = Math.Abs(Top + h - (area.Bottom - ScreenGap)) < SnapDistance;
            if ((left || right) && (top || bottom))
            {
                settings.Anchor = (top ? "Top" : "Bottom") + (left ? "Left" : "Right");
            }
            else
            {
                settings.Anchor = "Custom";
                settings.CustomFromTop = Top + h / 2 < area.Top + area.Height / 2;
                settings.CustomX = Left;
                settings.CustomY = settings.CustomFromTop ? Top : Top + h;
            }
            UpdateGrips();
            Place(h);
            Raise(Edited);
        }

        void WireRow(Border row, Action click)
        {
            bool pressed = false;
            row.MouseEnter += delegate { if (!editMode) row.SetResourceReference(Border.BackgroundProperty, "Hover"); };
            row.MouseLeave += delegate { row.Background = Brushes.Transparent; pressed = false; };
            row.MouseLeftButtonDown += (s, e) =>
            {
                if (editMode) return;
                pressed = true;
                row.SetResourceReference(Border.BackgroundProperty, "Pressed");
                e.Handled = true;
            };
            row.MouseLeftButtonUp += (s, e) =>
            {
                if (editMode || !pressed) return;
                pressed = false;
                e.Handled = true;
                row.Background = Brushes.Transparent;
                click();
                Dismiss();
            };
        }

        static void Raise(Action action)
        {
            if (action != null) action();
        }

        static Visibility Visible(bool visible) { return visible ? Visibility.Visible : Visibility.Collapsed; }

        static string Initials(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "?";
            int at = name.IndexOf('@');
            if (at > 0) name = name.Substring(0, at);
            var parts = name.Split(new[] { ' ', '.', '_', '-', ',', '(', ')', '"', '\'' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return "?";
            string first = parts[0].Substring(0, 1);
            string second = parts.Length > 1 ? parts[parts.Length - 1].Substring(0, 1) : "";
            return (first + second).ToUpperInvariant();
        }

        static uint StableHash(string s)
        {
            uint h = 2166136261;
            foreach (char c in s ?? "") { h ^= c; h *= 16777619; }
            return h;
        }
    }
}
