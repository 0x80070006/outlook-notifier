using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OutlookNotifier
{
    static class Native
    {
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_TOOLWINDOW = 0x00000080;
        public const int WS_EX_NOACTIVATE = 0x08000000;

        const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
        const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        const int DWMWA_BORDER_COLOR = 34;
        const int DWMWA_SYSTEMBACKDROP_TYPE = 38;
        const int DWMWCP_DONOTROUND = 1;
        const int DWMWCP_ROUND = 2;

        const int WCA_ACCENT_POLICY = 19;
        const int ACCENT_ENABLE_TRANSPARENTGRADIENT = 2;
        const int ACCENT_ENABLE_BLURBEHIND = 3;
        const int ACCENT_ENABLE_ACRYLICBLURBEHIND = 4;

        [StructLayout(LayoutKind.Sequential)]
        struct MARGINS { public int Left, Right, Top, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        struct AccentPolicy { public int AccentState; public int AccentFlags; public uint GradientColor; public int AnimationId; }

        [StructLayout(LayoutKind.Sequential)]
        struct WindowCompositionAttributeData { public int Attribute; public IntPtr Data; public int SizeOfData; }

        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        [DllImport("dwmapi.dll")] static extern int DwmExtendFrameIntoClientArea(IntPtr hwnd, ref MARGINS margins);
        [DllImport("user32.dll")] static extern int SetWindowCompositionAttribute(IntPtr hwnd, ref WindowCompositionAttributeData data);
        [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hwnd, int index);
        [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hwnd, int index, int value);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] public static extern bool AllowSetForegroundWindow(int processId);

        public static int WindowsBuild { get { return Environment.OSVersion.Version.Build; } }
        public static bool IsWindows11 { get { return WindowsBuild >= 22000; } }

        [StructLayout(LayoutKind.Sequential)] struct POINT { public int X, Y; }
        [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT point);

        public static System.Windows.Point CursorPixels()
        {
            POINT p;
            GetCursorPos(out p);
            return new System.Windows.Point(p.X, p.Y);
        }

        // Mica behind a regular window (Windows 11 22H2+). Also themes the title bar.
        public static bool ApplyMica(IntPtr hwnd, bool dark)
        {
            int darkMode = dark ? 1 : 0;
            DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, 4);
            if (WindowsBuild < 22621) return false;
            var margins = new MARGINS { Left = -1, Right = -1, Top = -1, Bottom = -1 };
            DwmExtendFrameIntoClientArea(hwnd, ref margins);
            int mica = 2; // DWMSBT_MAINWINDOW
            return DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE, ref mica, 4) == 0;
        }

        // Acrylic (or plain transparency) behind the window. Works on Windows 10 (1803+) and 11;
        // rounded corners and themed border are Windows 11 only.
        public static void ApplyBackdrop(IntPtr hwnd, bool dark, uint tintArgb, uint borderRgb, bool blur, bool rounded)
        {
            if (IsWindows11)
            {
                int round = rounded ? DWMWCP_ROUND : DWMWCP_DONOTROUND;
                DwmSetWindowAttribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, ref round, 4);
                int border = (int)(((borderRgb & 0xFF) << 16) | (borderRgb & 0xFF00) | ((borderRgb >> 16) & 0xFF));
                DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR, ref border, 4);
                int darkMode = dark ? 1 : 0;
                DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref darkMode, 4);
            }

            var margins = new MARGINS { Left = -1, Right = -1, Top = -1, Bottom = -1 };
            DwmExtendFrameIntoClientArea(hwnd, ref margins);

            // SetWindowCompositionAttribute expects ABGR.
            uint a = tintArgb >> 24, r = (tintArgb >> 16) & 0xFF, g = (tintArgb >> 8) & 0xFF, b = tintArgb & 0xFF;
            var accent = new AccentPolicy
            {
                AccentState = !blur ? ACCENT_ENABLE_TRANSPARENTGRADIENT
                            : WindowsBuild >= 17134 ? ACCENT_ENABLE_ACRYLICBLURBEHIND : ACCENT_ENABLE_BLURBEHIND,
                AccentFlags = 2,
                GradientColor = (a << 24) | (b << 16) | (g << 8) | r
            };
            int size = Marshal.SizeOf(accent);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(accent, ptr, false);
                var data = new WindowCompositionAttributeData { Attribute = WCA_ACCENT_POLICY, Data = ptr, SizeOfData = size };
                SetWindowCompositionAttribute(hwnd, ref data);
            }
            finally { Marshal.FreeHGlobal(ptr); }
        }

        // Waits for Outlook's main window (class rctrl_renwnd32, shown after the splash screen) and
        // minimizes it without activating it. Process.MainWindowHandle is unreliable here: it can be
        // the calendar reminders dialog.
        public static void MinimizeOutlookWhenReady(TimeSpan timeout)
        {
            DateTime deadline = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < deadline)
            {
                var pids = new System.Collections.Generic.HashSet<uint>(
                    System.Linq.Enumerable.Select(Process.GetProcessesByName("OUTLOOK"), p => (uint)p.Id));
                IntPtr main = IntPtr.Zero;
                EnumWindows((h, l) =>
                {
                    uint pid;
                    GetWindowThreadProcessId(h, out pid);
                    if (!pids.Contains(pid) || !IsWindowVisible(h)) return true;
                    var cls = new System.Text.StringBuilder(64);
                    GetClassName(h, cls, cls.Capacity);
                    if (cls.ToString() != "rctrl_renwnd32") return true;
                    main = h;
                    return false;
                }, IntPtr.Zero);

                if (main != IntPtr.Zero)
                {
                    System.Threading.Thread.Sleep(1500);   // let Outlook finish restoring its layout
                    ShowWindow(main, 7 /* SW_SHOWMINNOACTIVE */);
                    return;
                }
                System.Threading.Thread.Sleep(500);
            }
        }

        delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc, IntPtr lParam);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder name, int max);

        public static void BringProcessToFront(string processName)
        {
            foreach (var p in Process.GetProcessesByName(processName))
            {
                IntPtr h = p.MainWindowHandle;
                if (h == IntPtr.Zero) continue;
                if (IsIconic(h)) ShowWindow(h, 9 /* SW_RESTORE */);
                SetForegroundWindow(h);
                return;
            }
        }
    }
}
