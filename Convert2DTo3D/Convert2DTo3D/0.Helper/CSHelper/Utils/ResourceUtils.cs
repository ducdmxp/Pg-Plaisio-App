using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Convert2DTo3D
{
    public class ResourceUtils
    {
        public static ImageSource GetImage(string name)
        {
            return new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/" + name));
        }
    }
}