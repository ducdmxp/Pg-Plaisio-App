#region Namespaces

using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

#endregion

namespace Convert2DTo3D
{
    public class App : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            CurrentModule.InitialTemplate(application);
            CreateRibbon(application);
            return Result.Succeeded;
        }

        private void CreateRibbon(UIControlledApplication application)
        {
            string tabName = "Convert2DTo3D";
            string panelName = "Convert2DTo3D";

            var ribbonMng = CurrentModule.RibbonManage;

            ribbonMng.AddRibbonTab(tabName);
            var newPanel = ribbonMng.RibbonPanel(tabName, panelName);

            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData CmdConnectFCU = new PushButtonData("CmdConnectFCU", "Connect FCU", assemblyPath, typeof(Convert2DTo3D.Command.CmdConnectFCU).FullName);
            CmdConnectFCU.Image = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_16x16.ico"));
            CmdConnectFCU.LargeImage = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_32x32.ico"));
            newPanel.AddItem(CmdConnectFCU);

            PushButtonData CmdDeleteConnectFCU = new PushButtonData("CmdDeleteConnectFCU", "Delete Connect FCU", assemblyPath, typeof(Convert2DTo3D.Command.CmdDeleteConnectFCU).FullName);
            CmdDeleteConnectFCU.Image = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_16x16.ico"));
            CmdDeleteConnectFCU.LargeImage = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_32x32.ico"));
            newPanel.AddItem(CmdDeleteConnectFCU);

            PushButtonData CmdConnectWC = new PushButtonData("CmdConnectWC", "Connect WC", assemblyPath, typeof(Convert2DTo3D.Command.CmdConnectWC).FullName);
            CmdConnectWC.Image = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_16x16.ico"));
            CmdConnectWC.LargeImage = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_32x32.ico"));
            newPanel.AddItem(CmdConnectWC);

            PushButtonData CmdDeleteConnectWC = new PushButtonData("CmdDeleteConnectWC", "Delete Connect WC", assemblyPath, typeof(Convert2DTo3D.Command.CmdDeleteConnectWC).FullName);
            CmdDeleteConnectWC.Image = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_16x16.ico"));
            CmdDeleteConnectWC.LargeImage = new BitmapImage(new Uri("pack://application:,,,/Convert2DTo3D;component/Resources/Project1_32x32.ico"));
            newPanel.AddItem(CmdDeleteConnectWC);
        }

        private void AddImages(ButtonData buttonData,
                             string iconFolder,
                             string largeImage,
                             string smallImage)
        {
            if (!string.IsNullOrEmpty(iconFolder)
                && Directory.Exists(iconFolder))
            {
                string largeImagePath = Path.Combine(iconFolder, largeImage);
                if (File.Exists(largeImagePath))
                    buttonData.LargeImage = new BitmapImage(new Uri(largeImagePath));

                string smallImagePath = Path.Combine(iconFolder, smallImage);
                if (File.Exists(smallImagePath))
                    buttonData.Image = new BitmapImage(new Uri(smallImagePath));
            }
        }

        private string GetIconFolder()
        {
            string appDir = GetAppFolder();
            string imageDir = Path.Combine(appDir, "Icon");

            if (!Directory.Exists(imageDir))
                Directory.CreateDirectory(imageDir);

            return imageDir;
        }

        private string GetAppFolder()
        {
            string location = Assembly.GetExecutingAssembly().Location;
            string dir = Path.GetDirectoryName(location);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return dir;
        }
    }
}