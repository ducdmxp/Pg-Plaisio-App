using System;
using System.Windows;

namespace Convert2DTo3D.Utils
{
    public class IO
    {
        /// <summary>
        /// prompt user with information
        /// </summary>
        public static void ShowInfo(string content, string title = "Info")
        {
            MessageBox.Show(content, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// prompt user with warning
        /// </summary>
        public static void ShowWarning(string content, string title = "Warning")
        {
            MessageBox.Show(content, title, MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        /// <summary>
        /// prompt a yes/no question to ask for user decision
        /// </summary>
        public static MessageBoxResult ShowQuestion(string content, string title = "Question")
        {
            return MessageBox.Show(content, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
        }

        /// <summary>
        /// prompt user with an exception detail
        /// </summary>
        public static void ShowError(string error, string title = "Exception")
        {
            MessageBox.Show(error, title, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        /// <summary>
        /// prompt user with an exception detail
        /// </summary>
        public static void ShowException(Exception ex, string title = "Exception")
        {
            string content = ex.Message + "\n" + ex.StackTrace.ToString();
            MessageBox.Show(content, title, MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}