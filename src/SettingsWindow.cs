using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace OutlookNotifier
{
    // Edits a working copy of the settings. Every change is pushed live through Preview;
    // Saved fires on "Enregistrer", and closing any other way discards the changes.
    class SettingsWindow : Window
    {
        public event Action<Settings> Preview;
        public event Action<Settings> Saved;

        static readonly string[,] Accents =
        {
            { "#1450F5", "Bleu Outlook" }, { "#0F6CBD", "Bleu" }, { "#038387", "Sarcelle" }, { "#107C10", "Vert" },
            { "#8764B8", "Violet" }, { "#C239B3", "Orchidée" }, { "#D13438", "Rouge" }, { "#CA5010", "Orange" },
            { "#5D5A58", "Graphite" }, { "system", "Couleur d'accent de Windows" }
        };

        static readonly string[,] Tints =
        {
            { "", "Neutre" }, { "accent", "Couleur d'accent" }, { "#1450F5", "Bleu" }, { "#038387", "Sarcelle" },
            { "#107C10", "Vert" }, { "#8764B8", "Violet" }, { "#C239B3", "Orchidée" }, { "#CA5010", "Ambre" }
        };

        readonly Settings working;
        readonly FrameworkElement view;
        readonly WrapPanel accentSwatches, tintSwatches;
        readonly RadioButton themeAuto, themeLight, themeDark, posTopLeft, posTopRight, posBottomLeft, posBottomRight;
        readonly Slider opacitySlider, scaleSlider, widthSlider;
        readonly TextBlock opacityValue, scaleValue, widthValue, positionHint;
        readonly CheckBox blurSwitch, roundedSwitch, meetingsSwitch, previewSwitch, startupShowSwitch, launchSwitch,
                          avoidSwitch, reminderSwitch, autoOutlookSwitch;
        readonly FrameworkElement reminderIntervalRow;
        static readonly int[] ReminderChoices = { 5, 10, 15, 30, 60 };
        readonly RadioButton[] reminderButtons;
        IntPtr hwnd;
        bool loading;

        public Settings Working { get { return working; } }

        public SettingsWindow(Settings current)
        {
            working = current.Clone();

            Title = "Paramètres · Outlook Notifier";
            Width = 500;
            SizeToContent = SizeToContent.Height;
            MaxHeight = SystemParameters.WorkArea.Height - 40;
            ResizeMode = ResizeMode.CanMinimize;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            UseLayoutRounding = true;
            try { Icon = Ui.AppIcon(); } catch { }

            view = Ui.Load("Settings.xaml");
            Content = view;
            Ui.Find<Image>(view, "Logo").Source = MailIcon.Create();

            accentSwatches = Find<WrapPanel>("AccentSwatches");
            tintSwatches = Find<WrapPanel>("TintSwatches");
            themeAuto = Find<RadioButton>("ThemeAuto");
            themeLight = Find<RadioButton>("ThemeLight");
            themeDark = Find<RadioButton>("ThemeDark");
            posTopLeft = Find<RadioButton>("PosTopLeft");
            posTopRight = Find<RadioButton>("PosTopRight");
            posBottomLeft = Find<RadioButton>("PosBottomLeft");
            posBottomRight = Find<RadioButton>("PosBottomRight");
            opacitySlider = Find<Slider>("OpacitySlider");
            scaleSlider = Find<Slider>("ScaleSlider");
            widthSlider = Find<Slider>("WidthSlider");
            opacityValue = Find<TextBlock>("OpacityValue");
            scaleValue = Find<TextBlock>("ScaleValue");
            widthValue = Find<TextBlock>("WidthValue");
            positionHint = Find<TextBlock>("PositionHint");
            blurSwitch = Find<CheckBox>("BlurSwitch");
            roundedSwitch = Find<CheckBox>("RoundedSwitch");
            meetingsSwitch = Find<CheckBox>("MeetingsSwitch");
            previewSwitch = Find<CheckBox>("PreviewSwitch");
            startupShowSwitch = Find<CheckBox>("StartupShowSwitch");
            launchSwitch = Find<CheckBox>("LaunchSwitch");
            avoidSwitch = Find<CheckBox>("AvoidSwitch");
            reminderSwitch = Find<CheckBox>("ReminderSwitch");
            autoOutlookSwitch = Find<CheckBox>("AutoOutlookSwitch");
            reminderIntervalRow = Find<FrameworkElement>("ReminderIntervalRow");
            reminderButtons = ReminderChoices.Select(m => Find<RadioButton>("Remind" + m)).ToArray();

            if (!Native.IsWindows11)
            {
                Find<FrameworkElement>("RoundedRow").Visibility = Visibility.Collapsed;
                Find<FrameworkElement>("RoundedSeparator").Visibility = Visibility.Collapsed;
            }

            OnChecked(themeAuto, () => working.ThemeMode = "auto");
            OnChecked(themeLight, () => working.ThemeMode = "light");
            OnChecked(themeDark, () => working.ThemeMode = "dark");
            OnChecked(posTopLeft, () => working.Anchor = "TopLeft");
            OnChecked(posTopRight, () => working.Anchor = "TopRight");
            OnChecked(posBottomLeft, () => working.Anchor = "BottomLeft");
            OnChecked(posBottomRight, () => working.Anchor = "BottomRight");
            OnSlide(opacitySlider, v => working.Opacity = v / 100);
            OnSlide(scaleSlider, v => working.Scale = v / 100);
            OnSlide(widthSlider, v => working.Width = v);
            OnToggle(blurSwitch, v => working.Blur = v);
            OnToggle(roundedSwitch, v => working.Rounded = v);
            OnToggle(meetingsSwitch, v => working.ShowMeetings = v);
            OnToggle(previewSwitch, v => working.ShowPreview = v);
            OnToggle(startupShowSwitch, v => working.ShowOnStartup = v);
            OnToggle(launchSwitch, v => working.LaunchAtStartup = v);
            OnToggle(avoidSwitch, v => working.AvoidOthers = v);
            OnToggle(reminderSwitch, v => working.Reminder = v);
            OnToggle(autoOutlookSwitch, v => working.AutoLaunchOutlook = v);
            for (int i = 0; i < ReminderChoices.Length; i++)
            {
                int minutes = ReminderChoices[i];
                OnChecked(reminderButtons[i], () => working.ReminderMinutes = minutes);
            }

            Find<Button>("ResetButton").Click += delegate
            {
                working.CopyFrom(new Settings());
                LoadValues();
                Changed();
            };
            Find<Button>("CancelButton").Click += delegate { Close(); };
            Find<Button>("SaveButton").Click += delegate
            {
                if (Saved != null) Saved(working);
                Close();
            };
            PreviewKeyDown += (s, e) => { if (e.Key == Key.Escape) Close(); };

            SourceInitialized += delegate
            {
                hwnd = new WindowInteropHelper(this).Handle;
                ApplyTheme();
            };
            LoadValues();
            ApplyTheme();
        }

        // Called when the notification itself was dragged or resized.
        public void RefreshFromPreview()
        {
            LoadValues();
        }

        T Find<T>(string name) where T : class { return Ui.Find<T>(view, name); }

        void LoadValues()
        {
            loading = true;
            themeAuto.IsChecked = working.ThemeMode == "auto";
            themeLight.IsChecked = working.ThemeMode == "light";
            themeDark.IsChecked = working.ThemeMode == "dark";
            opacitySlider.Value = Math.Round(working.Opacity * 100);
            scaleSlider.Value = Math.Round(working.Scale * 100);
            widthSlider.Value = Math.Round(working.Width);
            blurSwitch.IsChecked = working.Blur;
            roundedSwitch.IsChecked = working.Rounded;
            meetingsSwitch.IsChecked = working.ShowMeetings;
            previewSwitch.IsChecked = working.ShowPreview;
            startupShowSwitch.IsChecked = working.ShowOnStartup;
            launchSwitch.IsChecked = working.LaunchAtStartup;
            avoidSwitch.IsChecked = working.AvoidOthers;
            reminderSwitch.IsChecked = working.Reminder;
            autoOutlookSwitch.IsChecked = working.AutoLaunchOutlook;
            int closest = ReminderChoices.OrderBy(m => Math.Abs(m - working.ReminderMinutes)).First();
            for (int i = 0; i < ReminderChoices.Length; i++) reminderButtons[i].IsChecked = ReminderChoices[i] == closest;
            posTopLeft.IsChecked = working.Anchor == "TopLeft";
            posTopRight.IsChecked = working.Anchor == "TopRight";
            posBottomLeft.IsChecked = working.Anchor == "BottomLeft";
            posBottomRight.IsChecked = working.Anchor == "BottomRight";
            loading = false;
            UpdateLabels();
            BuildSwatches();
        }

        void UpdateLabels()
        {
            opacityValue.Text = Math.Round(working.Opacity * 100) + " %";
            scaleValue.Text = Math.Round(working.Scale * 100) + " %";
            widthValue.Text = Math.Round(working.Width) + " px";
            reminderIntervalRow.IsEnabled = working.Reminder;
            reminderIntervalRow.Opacity = working.Reminder ? 1 : 0.45;
            switch (working.Anchor)
            {
                case "TopLeft": positionHint.Text = "En haut à gauche"; break;
                case "TopRight": positionHint.Text = "En haut à droite"; break;
                case "BottomLeft": positionHint.Text = "En bas à gauche"; break;
                case "BottomRight": positionHint.Text = "En bas à droite"; break;
                default: positionHint.Text = "Position libre"; break;
            }
        }

        void Changed()
        {
            if (loading) return;
            working.Normalize();
            UpdateLabels();
            BuildSwatches();
            ApplyTheme();
            if (Preview != null) Preview(working);
        }

        void ApplyTheme()
        {
            var theme = Theme.From(working);
            theme.ApplyTo(Resources);
            bool mica = hwnd != IntPtr.Zero && Native.ApplyMica(hwnd, theme.Dark);
            if (mica)
            {
                Background = Brushes.Transparent;
                HwndSource.FromHwnd(hwnd).CompositionTarget.BackgroundColor = Colors.Transparent;
            }
            else
            {
                SetResourceReference(BackgroundProperty, "WinBg");
            }
        }

        void BuildSwatches()
        {
            accentSwatches.Children.Clear();
            for (int i = 0; i < Accents.GetLength(0); i++)
            {
                string value = Accents[i, 0];
                Color color = value == "system" ? Theme.SystemAccent() : Theme.Parse(value, Theme.Brand);
                string glyph = value == "system" ? "" : null;
                accentSwatches.Children.Add(Swatch(color, glyph, Accents[i, 1], working.Accent == value, () => working.Accent = value));
            }

            tintSwatches.Children.Clear();
            var theme = Theme.From(working);
            for (int i = 0; i < Tints.GetLength(0); i++)
            {
                string value = Tints[i, 0];
                Color color = value == "" ? (theme.Dark ? Color.FromRgb(0x2B, 0x2B, 0x2B) : Color.FromRgb(0xF4, 0xF4, 0xF4))
                            : value == "accent" ? theme.Accent
                            : Theme.Parse(value, Theme.Brand);
                string glyph = value == "accent" ? "" : null;
                tintSwatches.Children.Add(Swatch(color, glyph, Tints[i, 1], working.Tint == value, () => working.Tint = value));
            }
        }

        FrameworkElement Swatch(Color color, string glyph, string tooltip, bool selected, Action select)
        {
            var fill = new Border
            {
                CornerRadius = new CornerRadius(12), Background = new SolidColorBrush(color),
                BorderThickness = new Thickness(1)
            };
            fill.SetResourceReference(Border.BorderBrushProperty, "ControlStroke");
            if (glyph != null)
            {
                bool lightFill = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255 > 0.62;
                fill.Child = new TextBlock
                {
                    Text = glyph, FontFamily = new FontFamily("Segoe Fluent Icons, Segoe MDL2 Assets"), FontSize = 11,
                    Foreground = lightFill ? Brushes.Black : Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center
                };
            }
            var ring = new Border
            {
                Width = 32, Height = 32, CornerRadius = new CornerRadius(16), Padding = new Thickness(2.5),
                BorderThickness = new Thickness(2), Margin = new Thickness(3), Child = fill,
                Background = Brushes.Transparent, Cursor = Cursors.Hand, ToolTip = tooltip
            };
            if (selected) ring.SetResourceReference(Border.BorderBrushProperty, "Fg1");
            else ring.BorderBrush = Brushes.Transparent;
            ring.MouseEnter += delegate { if (!selected) ring.SetResourceReference(Border.BorderBrushProperty, "TrackRail"); };
            ring.MouseLeave += delegate { if (!selected) ring.BorderBrush = Brushes.Transparent; };
            ring.MouseLeftButtonUp += delegate { select(); Changed(); };
            return ring;
        }

        void OnChecked(RadioButton button, Action apply)
        {
            button.Checked += delegate { if (loading) return; apply(); Changed(); };
        }

        void OnSlide(Slider slider, Action<double> apply)
        {
            slider.ValueChanged += delegate { if (loading) return; apply(Math.Round(slider.Value)); Changed(); };
        }

        void OnToggle(CheckBox box, Action<bool> apply)
        {
            box.Checked += delegate { if (loading) return; apply(true); Changed(); };
            box.Unchecked += delegate { if (loading) return; apply(false); Changed(); };
        }
    }
}
