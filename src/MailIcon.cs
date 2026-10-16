using System.Windows;
using System.Windows.Media;

namespace OutlookNotifier
{
    // The app mark (blue tile, envelope, notification dot) as a resolution-independent drawing.
    // Same proportions as build/make-icon.ps1.
    static class MailIcon
    {
        public static DrawingImage Create()
        {
            var blue = new SolidColorBrush(Theme.Brand);
            var white = new SolidColorBrush(Color.FromRgb(0xFA, 0xF9, 0xF6));

            var group = new DrawingGroup();
            group.Children.Add(new GeometryDrawing(blue, null, new RectangleGeometry(new Rect(0, 0, 1, 1), 0.23, 0.23)));
            group.Children.Add(new GeometryDrawing(white, null, new RectangleGeometry(new Rect(0.21, 0.31, 0.576, 0.408), 0.075, 0.075)));

            var flap = new StreamGeometry();
            using (var ctx = flap.Open())
            {
                ctx.BeginFigure(new Point(0.19, 0.365), false, false);
                ctx.LineTo(new Point(0.5, 0.56), true, true);
                ctx.LineTo(new Point(0.81, 0.365), true, true);
            }
            var pen = new Pen(blue, 0.048) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round, LineJoin = PenLineJoin.Round };
            group.Children.Add(new GeometryDrawing(null, pen, flap));

            group.Children.Add(new GeometryDrawing(blue, null, new EllipseGeometry(new Point(0.765, 0.314), 0.1, 0.1)));
            group.Children.Add(new GeometryDrawing(white, null, new EllipseGeometry(new Point(0.765, 0.314), 0.068, 0.068)));

            var image = new DrawingImage(group);
            image.Freeze();
            return image;
        }
    }
}
