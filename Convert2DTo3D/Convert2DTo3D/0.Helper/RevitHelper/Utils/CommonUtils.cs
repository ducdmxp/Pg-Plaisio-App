using Convert2DTo3D.Properties;
using System;
using System.Diagnostics;

namespace Convert2DTo3D
{
    public class CommonUtils
    {
        public static void ShowLoggerStatus()
        {
            if (Logger.Messages?.Count == 0)
            {
                MessageboxUtils.Show(MngLanguage.GetNameFromCurrentResource(() => Resources_ja_JP.GEN_Success), ShowType.Infomation);
            }
            //else if (Logger.Messages.All(k => k.MessageType == MessageType.Info))
            //{
            //    MessageboxUtils.Show(MngLanguage.GetNameFromCurrentResource(()=> Resources_ja_JP.GEN_Success), ShowType.Infomation);
            //}
            else
            {
                var window = WpfWindowControllerBase.GetResizeWindow(new StatusComponentUserControl(), 400, 320,
                    MngLanguage.GetNameFromCurrentResource(() => Resources_ja_JP.GEN_Information)
                    , WindownResize.RESIZE_BOTH);
                window.Show();
            }
        }

        public static void RunSafe(Action action, string overwriteMessage = null, MessageType messageTypeDef = MessageType.Error)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
#if DEBUG
                Logger.Log(messageTypeDef, ex.StackTrace);
                Logger.Log(messageTypeDef, overwriteMessage == null ? ex.Message : overwriteMessage);
#else
                Logger.Log(messageTypeDef, overwriteMessage == null ? ex.Message : overwriteMessage);
#endif
            }
        }

        public static void ShowThreadInfo()
        {
            int currentThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId;
            bool isUiThread = System.Windows.Application.Current?.Dispatcher?.CheckAccess() ?? false;
            if (System.Windows.Application.Current == null) isUiThread = true;
            Debug.WriteLine($"ThreadId[{currentThreadId}] - Is UI Thread[{isUiThread}]");
        }
    }
}