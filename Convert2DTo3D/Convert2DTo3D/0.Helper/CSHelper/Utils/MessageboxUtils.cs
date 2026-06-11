using Convert2DTo3D.Properties;

namespace Convert2DTo3D
{
    public enum ShowType
    {
        None,
        Infomation,
        Warning,
        Error
    }

    public class MessageboxUtils
    {
        public static void Show(string message, ShowType showType = ShowType.None, string title = "")
        {
            if (string.IsNullOrEmpty(title) == true)
            {
                switch (showType)
                {
                    case ShowType.Infomation:
                        title = MngLanguage.GetNameFromCurrentResource(() => Resources_ja_JP.GEN_Information);
                        break;

                    case ShowType.Warning:
                        title = MngLanguage.GetNameFromCurrentResource(() => Resources_ja_JP.GEN_Warning);
                        break;

                    case ShowType.Error:
                        title = MngLanguage.GetNameFromCurrentResource(() => Resources_ja_JP.GEN_Error);
                        break;

                    default:
                        title = "";
                        break;
                }
            }
            switch (showType)
            {
                case ShowType.Infomation:
                    System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    break;

                case ShowType.Warning:
                    System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    break;

                case ShowType.Error:
                    System.Windows.MessageBox.Show(message, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    break;

                default:
                    System.Windows.MessageBox.Show(message, title);
                    break;
            }
        }
    }
}