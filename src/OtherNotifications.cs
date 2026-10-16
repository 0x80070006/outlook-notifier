using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace OutlookNotifier
{
    // Finds on-screen notifications from other sources so the card can step aside:
    //  - Windows toasts and shell flyouts (top-level CoreWindow / "notification" windows of the
    //    shell), which live in a z-band above every topmost window, so they cannot be covered;
    //  - apps drawing their own pop-ups (Element, Teams, Slack...): small, visible, topmost windows.
    static class OtherNotifications
    {
        delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hwnd, int index);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder name, int max);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr hwnd, int attr, out RECT value, int size);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr hwnd, int attr, out int value, int size);

        const int GWL_EXSTYLE = -20;
        const int WS_EX_TOPMOST = 0x00000008;
        const int WS_EX_TRANSPARENT = 0x00000020;
        const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
        const int DWMWA_CLOAKED = 14;

        static readonly HashSet<string> IgnoredClasses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW", "tooltips_class32",
            "SysShadow", "Ghost", "IME", "MSCTFIME UI", "ForegroundStaging"
        };

        static readonly HashSet<string> ShellProcesses = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ShellExperienceHost", "ShellHost", "StartMenuExperienceHost", "explorer"
        };

        // Rectangles in physical pixels.
        public static List<Rect> Find()
        {
            var found = new List<Rect>();
            int self = Process.GetCurrentProcess().Id;
            double screenArea = SystemParameters.PrimaryScreenWidth * SystemParameters.PrimaryScreenHeight * Dpi.Scale * Dpi.Scale;

            EnumWindows((hwnd, l) =>
            {
                try
                {
                    if (!IsWindowVisible(hwnd) || IsIconic(hwnd)) return true;
                    int cloaked;
                    if (DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, out cloaked, 4) == 0 && cloaked != 0) return true;

                    uint pid;
                    GetWindowThreadProcessId(hwnd, out pid);
                    if (pid == self) return true;

                    var cls = new StringBuilder(128);
                    GetClassName(hwnd, cls, cls.Capacity);
                    string className = cls.ToString();
                    if (IgnoredClasses.Contains(className)) return true;

                    int ex = GetWindowLong(hwnd, GWL_EXSTYLE);
                    if ((ex & WS_EX_TRANSPARENT) != 0) return true;   // click-through overlays

                    RECT r;
                    if (DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, out r, Marshal.SizeOf(typeof(RECT))) != 0)
                        GetWindowRect(hwnd, out r);
                    double w = r.Right - r.Left, h = r.Bottom - r.Top;
                    if (w < 60 || h < 30) return true;
                    if (w * h > screenArea * 0.35) return true;       // full-screen or main windows

                    bool topmost = (ex & WS_EX_TOPMOST) != 0;
                    bool shellSurface = className == "Windows.UI.Core.CoreWindow" || IsShellNotification(hwnd, pid);
                    bool appPopup = topmost && w <= 720 && h <= 720;
                    if (shellSurface || appPopup) found.Add(new Rect(r.Left, r.Top, w, h));
                }
                catch { }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        static bool IsShellNotification(IntPtr hwnd, uint pid)
        {
            var title = new StringBuilder(128);
            GetWindowText(hwnd, title, title.Capacity);
            if (title.ToString().IndexOf("notification", StringComparison.OrdinalIgnoreCase) < 0) return false;
            try { return ShellProcesses.Contains(Process.GetProcessById((int)pid).ProcessName); }
            catch { return false; }
        }
    }

    static class Dpi
    {
        [DllImport("user32.dll")] static extern uint GetDpiForSystem();

        public static double Scale
        {
            get
            {
                try { return GetDpiForSystem() / 96.0; }
                catch { return 1.0; }
            }
        }
    }
}
