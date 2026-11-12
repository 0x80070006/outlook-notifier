using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using WinForms = System.Windows.Forms;

[assembly: AssemblyTitle("Outlook Notifier")]
[assembly: AssemblyProduct("Outlook Notifier")]
[assembly: AssemblyCompany("Outlook Notifier")]
[assembly: AssemblyVersion("1.3.1.0")]
[assembly: AssemblyFileVersion("1.3.1.0")]

namespace OutlookNotifier
{
    static class Program
    {
        const string InstanceMutex = @"Local\OutlookNotifier.Instance";
        const string QuitEvent = @"Global\OutlookNotifier.Quit";

        [STAThread]
        static int Main(string[] args)
        {
            Func<string, bool> has = flag => args.Any(a => a.Equals(flag, StringComparison.OrdinalIgnoreCase));

            if (has("--quit")) { StopRunningInstances(); return 0; }

            var settings = Settings.Load();
            if (has("--autostart") && !settings.LaunchAtStartup) return 0;

            bool created;
            using (var mutex = new Mutex(true, InstanceMutex, out created))
            {
                if (!created) return 0;
                new Notifier(settings, has("--demo"), has("--settings")).Run();
            }
            return 0;
        }

        // Used by the installer before uninstalling or upgrading: ask politely, then force.
        static void StopRunningInstances()
        {
            try
            {
                EventWaitHandle signal;
                if (EventWaitHandle.TryOpenExisting(QuitEvent, out signal)) using (signal) signal.Set();
            }
            catch { }

            int self = Process.GetCurrentProcess().Id;
            DateTime deadline = DateTime.UtcNow.AddSeconds(4);
            while (DateTime.UtcNow < deadline)
            {
                if (!Process.GetProcessesByName("OutlookNotifier").Any(p => p.Id != self)) return;
                Thread.Sleep(150);
            }
            foreach (var p in Process.GetProcessesByName("OutlookNotifier"))
                if (p.Id != self) try { p.Kill(); } catch { }
        }

        class Notifier
        {
            readonly bool demo;
            readonly bool openSettings;
            Settings settings;
            Application app;
            NotificationWindow popup;
            SettingsWindow settingsWindow;
            OutlookWatcher watcher;
            WinForms.NotifyIcon tray;
            OutlookState last = new OutlookState();
            bool startupHandled;
            bool outlookChecked;
            DateTime? outlookClosedAt;   // null while Outlook is running

            public Notifier(Settings settings, bool demo, bool openSettings)
            {
                this.settings = settings;
                this.demo = demo;
                this.openSettings = openSettings;
            }

            public void Run()
            {
                WinForms.Application.EnableVisualStyles();
                app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
                app.SessionEnding += delegate { app.Shutdown(); };
                app.DispatcherUnhandledException += (s, e) =>
                {
                    Log.Write("UI: " + e.Exception);
                    e.Handled = true;
                };

                popup = new NotificationWindow();
                popup.ApplySettings(settings);
                popup.OpenMailRequested += OpenInbox;
                popup.OpenCalendarRequested += OpenCalendar;
                popup.Edited += delegate { if (settingsWindow != null) settingsWindow.RefreshFromPreview(); };

                watcher = new OutlookWatcher();
                watcher.StateChanged += (state, arrived) => app.Dispatcher.BeginInvoke(new Action(() => OnState(state, arrived)));

                // Reminder: bring the card back while unread messages remain after it was closed.
                DateTime hiddenAt = DateTime.Now;
                popup.Hidden += delegate { hiddenAt = DateTime.Now; };
                var reminder = new DispatcherTimer { Interval = TimeSpan.FromSeconds(20) };
                reminder.Tick += delegate
                {
                    KeepOutlookOpen();
                    if (!settings.Reminder || popup.IsShown || popup.EditMode || last.LatestUnread == null) return;
                    if (DateTime.Now - hiddenAt >= TimeSpan.FromMinutes(settings.ReminderMinutes)) popup.ShowState(last);
                };
                reminder.Start();
                var firstCheck = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
                firstCheck.Tick += delegate { firstCheck.Stop(); KeepOutlookOpen(); };
                firstCheck.Start();

                CreateTray();
                WatchQuitSignal();
                watcher.Start();
                if (demo) app.Dispatcher.BeginInvoke(new Action(() => popup.ShowState(Demo.Fill(last))), DispatcherPriority.ApplicationIdle);
                if (openSettings) app.Dispatcher.BeginInvoke(new Action(OpenSettings), DispatcherPriority.ApplicationIdle);

                app.Run();

                watcher.Stop();
                tray.Visible = false;
                tray.Dispose();
            }

            // Opens Outlook right away when the app starts without it, then, if it gets closed,
            // again once the reminder interval has passed (so closing it on purpose is respected for a while).
            void KeepOutlookOpen()
            {
                bool running = OutlookWatcher.IsRunning("OUTLOOK");
                bool firstLook = !outlookChecked;
                outlookChecked = true;
                if (running) { outlookClosedAt = null; return; }
                if (!settings.AutoLaunchOutlook) return;

                if (outlookClosedAt == null) outlookClosedAt = DateTime.Now;
                if (firstLook || DateTime.Now - outlookClosedAt.Value >= TimeSpan.FromMinutes(settings.ReminderMinutes))
                {
                    outlookClosedAt = DateTime.Now;   // if it fails to start, try again after another interval
                    try
                    {
                        Process.Start(new ProcessStartInfo("outlook.exe") { UseShellExecute = true, WindowStyle = ProcessWindowStyle.Minimized });
                        Log.Write("Outlook started automatically");
                        // Outlook restores its own window state, so minimize it once its main window shows up.
                        ThreadPool.QueueUserWorkItem(delegate { Native.MinimizeOutlookWhenReady(TimeSpan.FromSeconds(90)); });
                    }
                    catch (Exception e) { Log.Write("Outlook start: " + e.Message); }
                }
            }

            void OnState(OutlookState state, bool arrived)
            {
                last = state;
                tray.Text = !state.Connected ? "Outlook Notifier · Outlook fermé"
                          : state.Unread == 0 ? "Outlook Notifier · aucun message non lu"
                          : state.Unread == 1 ? "Outlook Notifier · 1 message non lu"
                          : "Outlook Notifier · " + state.Unread + " messages non lus";

                if (popup.EditMode) { popup.UpdateState(state); return; }
                if (demo) return;   // keep the sample content on screen

                bool firstLook = !startupHandled && state.Connected;
                if (state.Connected) startupHandled = true;

                if (arrived && state.LatestUnread != null)
                    popup.ShowState(state);
                else if (firstLook && settings.ShowOnStartup && state.LatestUnread != null)
                    popup.ShowState(state);
                else if (popup.IsShown && state.LatestUnread == null)
                    popup.Dismiss();   // everything was read in Outlook
                else
                    popup.UpdateState(state);
            }

            void CreateTray()
            {
                var menu = new WinForms.ContextMenuStrip { ShowImageMargin = false };
                menu.Items.Add("Afficher la notification", null, delegate { popup.ShowState(last); });
                menu.Items.Add("Ouvrir la boîte de réception", null, delegate { OpenInbox(); });
                menu.Items.Add("Ouvrir le calendrier", null, delegate { OpenCalendar(); });
                menu.Items.Add(new WinForms.ToolStripSeparator());
                menu.Items.Add("Paramètres…", null, delegate { OpenSettings(); });
                menu.Items.Add("Quitter", null, delegate { app.Shutdown(); });

                tray = new WinForms.NotifyIcon { ContextMenuStrip = menu, Text = "Outlook Notifier", Icon = LoadIcon(), Visible = true };
                tray.MouseClick += (s, e) => { if (e.Button == WinForms.MouseButtons.Left) OpenSettings(); };
            }

            void OpenSettings()
            {
                if (settingsWindow != null)
                {
                    if (settingsWindow.WindowState == WindowState.Minimized) settingsWindow.WindowState = WindowState.Normal;
                    settingsWindow.Activate();
                    return;
                }

                settingsWindow = new SettingsWindow(settings);
                settingsWindow.Preview += s => popup.ApplySettings(s);
                settingsWindow.Saved += s =>
                {
                    settings = s.Clone();
                    settings.Save();
                    Startup.Apply(settings.LaunchAtStartup);
                };
                settingsWindow.Closed += delegate
                {
                    settingsWindow = null;
                    popup.ApplySettings(settings);
                    popup.ExitEditMode();
                };
                popup.ApplySettings(settingsWindow.Working);
                popup.EnterEditMode();
                settingsWindow.Show();
                settingsWindow.Activate();
            }

            static System.Drawing.Icon LoadIcon()
            {
                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OutlookNotifier.app.ico"))
                    return new System.Drawing.Icon(stream, WinForms.SystemInformation.SmallIconSize);
            }

            void WatchQuitSignal()
            {
                EventWaitHandle signal;
                try
                {
                    signal = new EventWaitHandle(false, EventResetMode.ManualReset, QuitEvent);
                    signal.Reset();   // a request still pending was meant for the previous instance
                }
                catch (Exception e) { Log.Write("Quit signal unavailable: " + e.Message); return; }
                ThreadPool.RegisterWaitForSingleObject(signal, delegate { app.Dispatcher.BeginInvoke(new Action(() => app.Shutdown())); }, null, -1, true);
            }

            void OpenInbox()
            {
                // The click just happened in this process, so it may hand foreground rights to Outlook.
                Native.AllowSetForegroundWindow(-1);
                watcher.OpenInbox();
            }

            void OpenCalendar()
            {
                Native.AllowSetForegroundWindow(-1);
                watcher.OpenCalendar();
            }
        }
    }
}
