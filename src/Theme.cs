using System;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace OutlookNotifier
{
    // Colors derived from the settings. Exposed to the XAML views as dynamic resources.
    class Theme
    {
        public static readonly Color Brand = Color.FromRgb(0x14, 0x50, 0xF5);

        public bool Dark;
        public Color Accent;
        public uint Tint;     // ARGB laid over the acrylic
        public uint Border;   // RGB window border (Windows 11)

        public static Theme From(Settings s)
        {
            bool dark = s.ThemeMode == "dark" || (s.ThemeMode != "light" && SystemIsDark());
            var t = new Theme { Dark = dark };
            t.Accent = s.Accent == "system" ? SystemAccent() : Parse(s.Accent, Brand);

            Color surface = dark ? Rgb(0x1C, 0x1C, 0x1C) : Rgb(0xF4, 0xF4, 0xF4);
            if (!string.IsNullOrEmpty(s.Tint))
            {
                Color tint = s.Tint == "accent" ? t.Accent : Parse(s.Tint, surface);
                surface = Mix(surface, tint, dark ? 0.28 : 0.14);
            }
            byte alpha = (byte)Math.Round(255 * Settings.Clamp(s.Opacity, 0.15, 1));
            t.Tint = ((uint)alpha << 24) | ((uint)surface.R << 16) | ((uint)surface.G << 8) | surface.B;
            Color border = dark ? Mix(surface, Colors.White, 0.13) : Mix(surface, Colors.Black, 0.15);
            t.Border = ((uint)border.R << 16) | ((uint)border.G << 8) | border.B;
            return t;
        }

        public void ApplyTo(ResourceDictionary r)
        {
            Color ink = Dark ? Colors.White : Colors.Black;
            r["Fg1"] = Solid(Dark ? Colors.White : Rgb(0x1A, 0x1A, 0x1A));
            r["Fg2"] = Solid(Alpha(ink, Dark ? 0xC8 : 0xA8));
            r["Fg3"] = Solid(Alpha(ink, Dark ? 0x8C : 0x78));
            r["Hover"] = Solid(Alpha(ink, Dark ? 0x12 : 0x0A));
            r["Pressed"] = Solid(Alpha(ink, Dark ? 0x08 : 0x05));
            r["ControlHover"] = Solid(Alpha(ink, Dark ? 0x18 : 0x10));
            r["Divider"] = Solid(Alpha(ink, Dark ? 0x18 : 0x14));
            r["Chip"] = Solid(Alpha(ink, Dark ? 0x14 : 0x0C));

            r["Accent"] = Solid(Accent);
            r["AccentFg"] = Solid(Luma(Accent) > 0.62 ? Colors.Black : Colors.White);
            r["AccentText"] = Solid(Dark ? Mix(Accent, Colors.White, 0.38) : Mix(Accent, Colors.Black, 0.12));
            r["AccentSoft"] = Solid(Alpha(Accent, Dark ? 0x40 : 0x24));

            // Settings window
            r["WinBg"] = Solid(Dark ? Rgb(0x20, 0x20, 0x20) : Rgb(0xF3, 0xF3, 0xF3));
            r["CardBg"] = Solid(Dark ? Alpha(Colors.White, 0x0D) : Alpha(Colors.White, 0xB3));
            r["CardStroke"] = Solid(Dark ? Alpha(Colors.Black, 0x33) : Alpha(Colors.Black, 0x10));
            r["ControlBg"] = Solid(Dark ? Alpha(Colors.White, 0x0F) : Alpha(Colors.White, 0xB3));
            r["ControlStroke"] = Solid(Dark ? Alpha(Colors.White, 0x18) : Alpha(Colors.Black, 0x1A));
            r["TrackStroke"] = Solid(Alpha(ink, Dark ? 0x9A : 0x72));
            r["TrackRail"] = Solid(Alpha(ink, Dark ? 0x50 : 0x40));
            r["ThumbBg"] = Solid(Dark ? Rgb(0x45, 0x45, 0x45) : Colors.White);
        }

        public static bool SystemIsDark()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    object v = key == null ? null : key.GetValue("SystemUsesLightTheme");
                    return !(v is int && (int)v == 1);
                }
            }
            catch { return true; }
        }

        public static Color SystemAccent()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\DWM"))
                {
                    object v = key == null ? null : key.GetValue("AccentColor");
                    if (v is int)
                    {
                        uint abgr = unchecked((uint)(int)v);
                        return Color.FromRgb((byte)abgr, (byte)(abgr >> 8), (byte)(abgr >> 16));
                    }
                }
            }
            catch { }
            return Brand;
        }

        public static Color Parse(string hex, Color fallback)
        {
            uint v;
            if (hex != null && hex.Length == 7 && hex[0] == '#' &&
                uint.TryParse(hex.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v))
                return Color.FromRgb((byte)(v >> 16), (byte)(v >> 8), (byte)v);
            return fallback;
        }

        public static Color Mix(Color a, Color b, double t)
        {
            return Color.FromRgb((byte)(a.R + (b.R - a.R) * t), (byte)(a.G + (b.G - a.G) * t), (byte)(a.B + (b.B - a.B) * t));
        }

        static double Luma(Color c) { return (0.299 * c.R + 0.587 * c.G + 0.114 * c.B) / 255; }
        static Color Rgb(byte r, byte g, byte b) { return Color.FromRgb(r, g, b); }
        static Color Alpha(Color c, int a) { return Color.FromArgb((byte)a, c.R, c.G, c.B); }

        static Brush Solid(Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();
            return b;
        }
    }
}
