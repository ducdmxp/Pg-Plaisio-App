using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Brush = System.Windows.Media.Brush;

namespace Convert2DTo3D
{
    public static class CanvasExtension
    {
        public static System.Windows.Shapes.Line AddLine(this Canvas canvas, double X1, double Y1, double X2, double Y2, Brush brush)
        {
            System.Windows.Shapes.Line gridH = new System.Windows.Shapes.Line();
            gridH.Stroke = brush;
            gridH.X1 = X1;
            gridH.X2 = X2;  // 150 too far
            gridH.Y1 = Y1;
            gridH.Y2 = Y2;
            gridH.StrokeThickness = 1;
            canvas.Children.Add(gridH);
            return gridH;
        }

        public static System.Windows.Shapes.Line AddAxis(this Canvas canvas, double X1, double Y1, double X2, double Y2, Brush brush)
        {
            System.Windows.Shapes.Line gridH = new System.Windows.Shapes.Line();
            gridH.Stroke = brush;
            gridH.X1 = X1;
            gridH.X2 = X2;  // 150 too far
            gridH.Y1 = Y1;
            gridH.Y2 = Y2;
            gridH.StrokeThickness = 1;
            gridH.StrokeDashArray = new DoubleCollection() { 4, 2 }; // 4pixcel liền thì có 2 pixcel đứt
            canvas.Children.Add(gridH);
            return gridH;
        }

        public static System.Windows.Shapes.Line AddLineDot(this Canvas canvas, double X1, double Y1, double X2, double Y2, Brush brush)
        {
            System.Windows.Shapes.Line gridH = new System.Windows.Shapes.Line();
            gridH.Stroke = brush;
            gridH.X1 = X1;
            gridH.X2 = X2;  // 150 too far
            gridH.Y1 = Y1;
            gridH.Y2 = Y2;
            gridH.StrokeThickness = 1;
            gridH.StrokeDashArray = new DoubleCollection() { 1, 1 }; // 4pixcel liền thì có 2 pixcel đứt
            canvas.Children.Add(gridH);
            return gridH;
        }

        public static System.Windows.Shapes.Rectangle AddRectangle(this Canvas canvas, double width, double height, double Oleft, double Otop, Brush brush)
        {
            System.Windows.Shapes.Rectangle rec = new System.Windows.Shapes.Rectangle();
            rec.Width = width;
            rec.Height = height;
            Canvas.SetLeft(rec, Oleft);
            Canvas.SetTop(rec, Otop);
            rec.Stroke = brush;
            canvas.Children.Add(rec);
            return rec;
        }

        public static System.Windows.Shapes.Ellipse AddCircle(this Canvas canvas, double width, double height, double x, double y, Brush brush)
        {
            System.Windows.Shapes.Ellipse ellipse = new System.Windows.Shapes.Ellipse();
            ellipse.Stroke = brush;
            ellipse.Width = width;
            ellipse.Height = height;
            Canvas.SetLeft(ellipse, x - width / 2.0);
            Canvas.SetTop(ellipse, y - height / 2.0);
            canvas.Children.Add(ellipse);
            return ellipse;
        }

        public static System.Windows.Shapes.Rectangle AddSquare(this Canvas canvas, double CenterX, double CenterY, double width, Brush brush)
        {
            System.Windows.Shapes.Rectangle rec = new System.Windows.Shapes.Rectangle();
            rec.Width = width;
            rec.Height = width;
            Canvas.SetLeft(rec, CenterX - width / 2);
            Canvas.SetTop(rec, CenterY - width / 2);
            rec.Stroke = brush;
            canvas.Children.Add(rec);
            return rec;
        }

        public static TextBlock AddText(this Canvas canvas, string text, double Angle, double x, double y, double height, Brush brush)
        {
            TextBlock textBlock = new TextBlock();
            textBlock.Text = text;
            textBlock.FontSize = height;
            double degree = 180 * Angle / Math.PI;
            RotateTransform rote = new RotateTransform(-degree);
            if (Convert.ToInt32(degree) == 180)
            {
                rote = new RotateTransform(0);
                Angle = 0;
            }
            textBlock.RenderTransform = rote;
            textBlock.Foreground = brush;
            System.Windows.Size size = MeasureString(textBlock);
            Canvas.SetLeft(textBlock, x + size.Width / 2.0);
            Canvas.SetTop(textBlock, y + size.Height);
            canvas.Children.Add(textBlock);
            return textBlock;
        }

        public static TextBlock AddText(this Canvas canvas, string text, double x, double y, double height, Brush brush)
        {
            TextBlock textBlock = new TextBlock();
            textBlock.Text = text;
            textBlock.FontSize = height;
            textBlock.Foreground = brush;
            Canvas.SetLeft(textBlock, x);
            Canvas.SetTop(textBlock, y);
            canvas.Children.Add(textBlock);
            return textBlock;
        }

        public static TextBlock AddTextCenter(this Canvas canvas, string text, double x, double y, double height, Brush brush)
        {
            TextBlock textBlock = new TextBlock();
            textBlock.Text = text;
            textBlock.FontSize = height;
            textBlock.Foreground = brush;
            Canvas.SetLeft(textBlock, x);
            Canvas.SetTop(textBlock, y);
            canvas.Children.Add(textBlock);
            var sizeT1 = CanvasExtension.MeasureString(textBlock);
            Canvas.SetLeft(textBlock, x - sizeT1.Width / 2.0);
            Canvas.SetTop(textBlock, y - sizeT1.Height / 2.0);
            return textBlock;
        }

        public static System.Windows.Size MeasureString(TextBlock textBlock)
        {
            var formattedText = new System.Windows.Media.FormattedText(
                textBlock.Text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface(textBlock.FontFamily, textBlock.FontStyle, textBlock.FontWeight, textBlock.FontStretch),
               textBlock.FontSize,
                System.Windows.Media.Brushes.Black,
                new NumberSubstitution(),
                1);

            return new System.Windows.Size(formattedText.Width, formattedText.Height);
        }

        public static System.Windows.Point RotatePoint(this Canvas canvas, System.Windows.Point point, System.Windows.Point center, double radians)
        {
            // Translate point to origin
            double x = point.X - center.X;
            double y = point.Y - center.Y;

            // Apply rotation
            double newX = x * Math.Cos(radians) - y * Math.Sin(radians);
            double newY = x * Math.Sin(radians) + y * Math.Cos(radians);

            // Translate point back
            newX += center.X;
            newY += center.Y;

            return new System.Windows.Point(newX, newY);
        }
    }
}