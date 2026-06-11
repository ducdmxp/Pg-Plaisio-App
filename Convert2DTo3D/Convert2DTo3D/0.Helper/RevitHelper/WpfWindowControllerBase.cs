namespace Convert2DTo3D
{
    //    public enum WindowFit
    //    {
    //        FIX,
    //        WIDTH,
    //        HEGHT,
    //        WIDTH_AND_HEIGHT
    //    }
    //    public abstract class WpfWindowControllerBase : IExternalApplication
    //    {
    //        public System.Windows.Window Window { get; private set; }
    //        private List<RibbonItem> m_ribbonItems = null;
    //        private string m_btnName = null;
    //        public static WpfWindowControllerBase Instance;
    //        public Result OnStartup(UIControlledApplication application)
    //        {
    //            return Result.Succeeded;
    //        }
    //        public Result OnShutdown(UIControlledApplication application)
    //        {
    //            Window?.Close();
    //            return Result.Succeeded;
    //        }
    //        /// <summary>
    //        ///
    //        /// </summary>
    //        /// <typeparam name="T"></typeparam>
    //        /// <param name="m_uiapp"></param>
    //        /// <param name="ResizeMode"></param>

    //        public void ShowForm<T>(Func<ExternalEvent, UserControl> registerUserControl,
    //                                 string title = "",
    //                                WindowFit type = WindowFit.FIX,
    //                                ResizeMode ResizeMode = ResizeMode.NoResize) where T : IExternalEventHandler, new()
    //        {
    //            ShowFormRevoke(() =>
    //            {
    //                T handler = new T();
    //                if (Instance == null) Instance = this;
    //                var exEvent = ExternalEvent.Create(handler);
    //                try
    //                {
    //                    int indexMethod = 3;
    //#if DEBUG
    //#else
    //                    indexMethod = 2;
    //#endif
    //                    var methodInfo = new StackTrace().GetFrame(indexMethod).GetMethod();
    //                    m_btnName = methodInfo.ReflectedType.Name;
    //                    m_ribbonItems = CurrentModule.RibbonManage.GetButtons(m_btnName);
    //                }
    //                catch (Exception)
    //                { }
    //                var userControl = registerUserControl(exEvent);
    //                if (userControl == null) return null;
    //                var w = new System.Windows.Window();
    //                Window = w;
    //                Window.Content = userControl;
    //                if (m_ribbonItems?.Count == 1) Window.Title = m_ribbonItems[0].ItemText;
    //                SetSizeWindow(ref w, userControl, type);
    //                if (string.IsNullOrEmpty(title) == false) Window.Title = title;
    //                Window.ResizeMode = ResizeMode;
    //                return Window;
    //            });
    //        }
    //        private void ShowFormRevoke(Func<Window> funcW)
    //        {
    //            var Window = funcW();
    //            WindowHandleSearch handle = WindowHandleSearch.MainWindowHandle;
    //            handle.SetAsOwner(Window);
    //            Window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
    //            var RibbonItem = m_ribbonItems != null && m_ribbonItems.Count > 0 ? m_ribbonItems[0] : null;
    //            if (RibbonItem != null)
    //            {
    //                Window.Title = RibbonItem.ItemText;
    //                if (RibbonItem is PushButton pbt) Window.Icon = pbt.Image;
    //                if (RibbonItem is SplitButton spl)
    //                {
    //                    foreach (var item in spl.GetItems())
    //                    {
    //                        if (item is PushButton pb)
    //                        {
    //                            if (pb.ClassName.Contains(m_btnName))
    //                            {
    //                                Window.Icon = pb.Image;
    //                                break;
    //                            }
    //                        }
    //                    }
    //                }
    //            }
    //            Window.Closed += (s, e) =>
    //            {
    //                if (m_ribbonItems != null)
    //                {
    //                    foreach (var rbitem in m_ribbonItems)
    //                    {
    //                        rbitem.Enabled = true;
    //                    }
    //                }
    //            };
    //            Window.KeyDown += (s, e) => { if (e.Key == System.Windows.Input.Key.Escape) Window?.Close(); };

    //            Window.Show();
    //            if (m_ribbonItems != null)
    //            {
    //                foreach (var rbitem in m_ribbonItems)
    //                {
    //                    rbitem.Enabled = false;
    //                }
    //            }
    //        }

    //        public void ShowFormFixHeight<T>(Func<ExternalEvent, UserControl> registerUserControl,
    //                                 int width, int height,
    //                                 string title = "") where T : IExternalEventHandler, new()
    //        {
    //            ShowFormRevoke(() =>
    //            {
    //                T handler = new T();
    //                if (Instance == null) Instance = this;
    //                var exEvent = ExternalEvent.Create(handler);
    //                try
    //                {
    //                    int indexMethod = 3;
    //#if DEBUG
    //#else
    //                    indexMethod = 2;
    //#endif
    //                    var methodInfo = new StackTrace().GetFrame(indexMethod).GetMethod();
    //                    m_btnName = methodInfo.ReflectedType.Name;
    //                    m_ribbonItems = CurrentModule.RibbonManage.GetButtons(m_btnName);
    //                }
    //                catch (Exception)
    //                { }
    //                var userControl = registerUserControl(exEvent);
    //                if (userControl == null) return null;
    //                Window w = new Window
    //                {
    //                    MinHeight = height + 20,
    //                    MaxHeight = height + 20,
    //                    MinWidth = width + 20,
    //                    Width = width + 20,
    //                    Height = height + 20,
    //                    ResizeMode = ResizeMode.CanResize
    //                };
    //                if (m_ribbonItems?.Count == 1) w.Title = m_ribbonItems[0].ItemText;
    //                if (string.IsNullOrEmpty(title) == false) Window.Title = title;
    //                Window = w;
    //                Window.Content = userControl;
    //                Window.ResizeMode = ResizeMode.CanResize;
    //                return Window;
    //            });
    //        }
    //        public void ShowFormCanResize<T>(Func<ExternalEvent, UserControl> registerUserControl,
    //                         int width, int height,
    //                         string title = "") where T : IExternalEventHandler, new()
    //        {
    //            ShowFormRevoke(() =>
    //            {
    //                T handler = new T();
    //                if (Instance == null) Instance = this;
    //                var exEvent = ExternalEvent.Create(handler);
    //                try
    //                {
    //                    int indexMethod = 3;
    //#if DEBUG
    //#else
    //                    indexMethod = 2;
    //#endif
    //                    var methodInfo = new StackTrace().GetFrame(indexMethod).GetMethod();
    //                    m_btnName = methodInfo.ReflectedType.Name;
    //                    m_ribbonItems = CurrentModule.RibbonManage.GetButtons(m_btnName);
    //                }
    //                catch (Exception)
    //                { }
    //                var userControl = registerUserControl(exEvent);
    //                if (userControl == null) return null;
    //                Window w = new Window
    //                {
    //                    MinHeight = height + 20,
    //                    MinWidth = width + 20,
    //                    Width = width + 20,
    //                    Height = height + 20,
    //                    ResizeMode = ResizeMode.CanResize
    //                };
    //                if (m_ribbonItems?.Count == 1) w.Title = m_ribbonItems[0].ItemText;
    //                if (string.IsNullOrEmpty(title) == false) Window.Title = title;
    //                Window = w;
    //                Window.Content = userControl;
    //                Window.ResizeMode = ResizeMode.CanResize;
    //                return Window;
    //            });
    //        }
    //        private static void SetSizeWindow(ref System.Windows.Window w, UserControl content, WindowFit type = WindowFit.FIX)
    //        {
    //            switch (type)
    //            {
    //                case WindowFit.FIX:
    //                    w.Height = content.Height + 40;
    //                    w.Width = content.Width + 20;
    //                    break;
    //                case WindowFit.WIDTH:
    //                    w.Height = content.Height + 40;
    //                    w.SizeToContent = SizeToContent.Width;
    //                    break;
    //                case WindowFit.HEGHT:
    //                    w.Width = content.Width + 20;
    //                    w.SizeToContent = SizeToContent.Height;
    //                    break;
    //                case WindowFit.WIDTH_AND_HEIGHT:
    //                    w.SizeToContent = SizeToContent.WidthAndHeight;
    //                    break;
    //            }
    //        }
    //        public static Window GetWindowMetro(UserControl content, Window parentWindow = null, WindowFit type = WindowFit.FIX)
    //        {
    //            System.Windows.Window w = new System.Windows.Window();
    //            w.Content = content;
    //            w.Background = content.Background;
    //            SetSizeWindow(ref w, content, type);
    //            w.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
    //            if (parentWindow != null)
    //            {
    //                var wih = new WindowInteropHelper(parentWindow);
    //                IntPtr hWnd = wih.Handle;
    //                new WindowInteropHelper(w) { Owner = hWnd };
    //                w.ResizeMode = ResizeMode.NoResize;
    //            }
    //            return w;
    //        }
    //        public static Window GetWindowMetroFixHeight(System.Windows.Controls.UserControl content, int width, int height, Window parentWindow = null)
    //        {
    //            Window w = new Window
    //            {
    //                MinHeight = height + 20,
    //                MaxHeight = height + 20,
    //                MinWidth = width + 20,
    //                Width = width + 20,
    //                Height = height + 20,
    //                ResizeMode = ResizeMode.CanResize
    //            };
    //            w.Content = content;
    //            w.Background = content.Background;
    //            w.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
    //            if (parentWindow != null)
    //            {
    //                var wih = new WindowInteropHelper(parentWindow);
    //                IntPtr hWnd = wih.Handle;
    //                new WindowInteropHelper(w) { Owner = hWnd };
    //                w.ResizeMode = ResizeMode.CanResize;
    //            }
    //            return w;
    //        }
    //        public static Window GetWindowMetroCanResize(System.Windows.Controls.UserControl content, int width, int height, Window parentWindow = null)
    //        {
    //            Window w = new Window
    //            {
    //                MinHeight = height + 20,
    //                MinWidth = width + 20,
    //                Width = width + 20,
    //                Height = height + 20,
    //                ResizeMode = ResizeMode.CanResize
    //            };
    //            w.Content = content;
    //            w.Background = content.Background;
    //            w.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
    //            if (parentWindow != null)
    //            {
    //                var wih = new WindowInteropHelper(parentWindow);
    //                IntPtr hWnd = wih.Handle;
    //                new WindowInteropHelper(w) { Owner = hWnd };
    //                w.ResizeMode = ResizeMode.CanResize;
    //            }
    //            return w;
    //        }

    //    }
}