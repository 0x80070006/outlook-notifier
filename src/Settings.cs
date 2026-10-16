using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace OutlookNotifier
{
    // User preferences, stored per user in %APPDATA%\OutlookNotifier\settings.ini.
    class Settings
    {
        public string ThemeMode = "auto";       // auto | light | dark
        public string Accent = "#1450F5";       // #RRGGBB or "system"
        public string Tint = "";                // "" (neutral), "accent" or #RRGGBB
        public double Opacity = 0.82;           // background opacity, 0.15 to 1
        public bool Blur = true;
        public bool Rounded = true;
        public double Scale = 1.0;              // 0.7 to 1.6
        public double Width = 340;              // card width at 100 %, 260 to 600
        public string Anchor = "BottomRight";   // TopLeft | TopRight | BottomLeft | BottomRight | Custom
        public double CustomX;
        public double CustomY;                  // top edge if CustomFromTop, bottom edge otherwise
        public bool CustomFromTop;
        public bool AvoidOthers = true;         // step aside from other notifications
        public bool Reminder = true;            // show again while unread messages remain
        public double ReminderMinutes = 15;
        public bool AutoLaunchOutlook = true;   // open Outlook at startup, and again after ReminderMinutes if closed
        public bool ShowMeetings = true;
        public bool ShowPreview = true;
        public bool ShowOnStartup = true;
        public bool LaunchAtStartup = true;

        static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookNotifier", "settings.ini");

        static FieldInfo[] Fields
        {
            get { return typeof(Settings).GetFields(BindingFlags.Public | BindingFlags.Instance); }
        }

        public Settings Clone() { return (Settings)MemberwiseClone(); }

        public void CopyFrom(Settings other)
        {
            foreach (var f in Fields) f.SetValue(this, f.GetValue(other));
        }

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                if (!File.Exists(FilePath)) return s;
                foreach (string line in File.ReadAllLines(FilePath))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim().TrimStart('﻿');
                    var field = Fields.FirstOrDefault(f => f.Name == key);
                    if (field == null) continue;
                    string value = line.Substring(eq + 1).Trim();
                    if (field.FieldType == typeof(string)) field.SetValue(s, value);
                    else if (field.FieldType == typeof(bool)) field.SetValue(s, value == "true");
                    else if (field.FieldType == typeof(double))
                    {
                        double d;
                        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) field.SetValue(s, d);
                    }
                }
            }
            catch (Exception e) { Log.Write("Settings load: " + e.Message); }
            s.Normalize();
            return s;
        }

        public void Save()
        {
            Normalize();
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                var lines = Fields.Select(f =>
                {
                    object v = f.GetValue(this);
                    string text = v is bool ? ((bool)v ? "true" : "false")
                                : v is double ? ((double)v).ToString("R", CultureInfo.InvariantCulture)
                                : Convert.ToString(v);
                    return f.Name + "=" + text;
                });
                File.WriteAllLines(FilePath, lines);
            }
            catch (Exception e) { Log.Write("Settings save: " + e.Message); }
        }

        public void Normalize()
        {
            Opacity = Clamp(Opacity, 0.15, 1);
            Scale = Clamp(Scale, 0.7, 1.6);
            Width = Clamp(Width, 260, 600);
            ReminderMinutes = Clamp(ReminderMinutes, 1, 240);
            if (new[] { "TopLeft", "TopRight", "BottomLeft", "BottomRight", "Custom" }.All(a => a != Anchor)) Anchor = "BottomRight";
            if (new[] { "auto", "light", "dark" }.All(m => m != ThemeMode)) ThemeMode = "auto";
        }

        public static double Clamp(double v, double min, double max)
        {
            return double.IsNaN(v) ? min : Math.Max(min, Math.Min(max, v));
        }
    }

    // "Launch with Windows". The MSI registers the app for every user (HKLM Run, with --autostart);
    // turning it off is then a per-user setting checked at launch. Without the MSI entry the app
    // manages its own HKCU Run value.
    static class Startup
    {
        const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
        const string ValueName = "OutlookNotifier";

        public static void Apply(bool enabled)
        {
            bool machineWide = false;
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var key = hklm.OpenSubKey(RunKey))
                    machineWide = key != null && key.GetValue(ValueName) != null;
            }
            catch { }

            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RunKey))
                {
                    if (enabled && !machineWide)
                        key.SetValue(ValueName, "\"" + Assembly.GetExecutingAssembly().Location + "\" --autostart");
                    else
                        key.DeleteValue(ValueName, false);
                }
            }
            catch (Exception e) { Log.Write("Startup: " + e.Message); }
        }
    }
}
