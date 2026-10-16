using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace OutlookNotifier
{
    // The views are plain XAML embedded as resources and loaded at runtime (no XAML compiler needed).
    static class Ui
    {
        const string PresentationNs = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";

        public static FrameworkElement Load(string name)
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("OutlookNotifier." + name))
                return (FrameworkElement)XamlReader.Load(s);
        }

        public static ControlTemplate Template(string targetType, string body)
        {
            return (ControlTemplate)XamlReader.Parse(
                "<ControlTemplate xmlns=\"" + PresentationNs + "\" TargetType=\"" + targetType + "\">" + body + "</ControlTemplate>");
        }

        public static ImageSource AppIcon()
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("OutlookNotifier.app.ico"))
                return BitmapFrame.Create(s, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
        }

        public static T Find<T>(FrameworkElement root, string name) where T : class
        {
            return root.FindName(name) as T;
        }
    }
}
