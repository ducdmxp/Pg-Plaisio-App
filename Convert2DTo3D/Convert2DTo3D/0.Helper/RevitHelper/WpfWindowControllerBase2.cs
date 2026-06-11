using Autodesk.Revit.UI;
using CreateRack.DetailCommands.Utils;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Shell;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using ColorConverter = System.Windows.Media.ColorConverter;
using Grid = System.Windows.Controls.Grid;
using Image = System.Windows.Controls.Image;
using Window = System.Windows.Window;

namespace Convert2DTo3D
{
    public enum WindowFit
    {
        FIX,
        WIDTH_CONTENT,
        HEGHT_CONTENT,
        BOTH_CONTENT
    }

    public enum WindownResize
    {
        RESIZE_WIDTH,
        RESIZE_HEIGHT,
        RESIZE_BOTH,
    }

    public class WpfWindowControllerBase : IExternalApplication
    {
        private static string _windowTitleName = "ID_IB_Title";
        private static string _IconId = "ID_Icon";
        private static string _TitleId = "ID_Title";
        private static string _MiniId = "ID_Minimum";
        private static string _MaxiId = "ID_Maximum";
        public System.Windows.Window Window { get; private set; }
        public static WpfWindowControllerBase Instance;

        public Result OnStartup(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            Window?.Close();
            return Result.Succeeded;
        }

        public void ShowFixWindow<T>(Func<ExternalEvent, System.Windows.Controls.UserControl> registerUserControl, string title = "", WindowFit type = WindowFit.FIX, Action infoAction = null) where T : IExternalEventHandler, new()
        {
            T handler = new T();
            if (Instance == null) Instance = this;
            var exEvent = ExternalEvent.Create(handler);
            var userControl = registerUserControl(exEvent);
            if (userControl == null) return;
            InvokeShowWindow(userControl, (uc) => PrivateGetFixWindow(uc, title, type, null), title);
            ConfigWindowStyle(Window, true, infoAction);
            userControl.Background = CurrentModule.BgColor;
            Window.Background = CurrentModule.BgColor;
            Window?.Show();
        }

        public static System.Windows.Window GetFixWindow(System.Windows.Controls.UserControl content, string title = "", WindowFit type = WindowFit.FIX, Action infoAction = null, Window parentWindow = null)
        {
            System.Windows.Window w = new System.Windows.Window();
            switch (type)
            {
                case WindowFit.FIX:
                    w.Height = content.Height + 40;
                    w.Width = content.Width + 20;
                    break;

                case WindowFit.WIDTH_CONTENT:
                    w.Height = content.Height + 40;
                    w.SizeToContent = SizeToContent.Width;
                    break;

                case WindowFit.HEGHT_CONTENT:
                    w.Width = content.Width + 20;
                    w.SizeToContent = SizeToContent.Height;
                    break;

                case WindowFit.BOTH_CONTENT:
                    w.SizeToContent = SizeToContent.WidthAndHeight;
                    break;
            }
            w.ResizeMode = ResizeMode.NoResize;
            SetContentForWindow(ref w, ref content, title, parentWindow);
            ConfigWindowStyle(w, true, infoAction);
            content.Background = CurrentModule.BgColor;
            w.Background = CurrentModule.BgColor;
            return w;
        }

        public void ShowResizeWindow<T>(Func<ExternalEvent, System.Windows.Controls.UserControl> registerUserControl, int width, int height, string title = "", WindownResize type = WindownResize.RESIZE_HEIGHT, Action infoAction = null) where T : IExternalEventHandler, new()
        {
            T handler = new T();
            if (Instance == null) Instance = this;
            var exEvent = ExternalEvent.Create(handler);
            var userControl = registerUserControl(exEvent);
            if (userControl == null) return;
            InvokeShowWindow(userControl, (uc) => PrivateGetResizeWindow(uc, width, height, title, type, null), title);
            ConfigWindowStyle(Window, false, infoAction);
            userControl.Background = CurrentModule.BgColor;
            Window.Background = CurrentModule.BgColor;
            Window?.Show();
        }

        public static System.Windows.Window GetResizeWindow(System.Windows.Controls.UserControl content, int width, int height, string title = "", WindownResize type = WindownResize.RESIZE_WIDTH, Action infoAction = null, Window parentWindow = null)
        {
            Window w = null;
            switch (type)
            {
                case WindownResize.RESIZE_WIDTH:
                    w = new Window
                    {
                        MinHeight = height + 20,
                        MaxHeight = height + 20,
                        MinWidth = width + 20,
                        Width = width + 20,
                        Height = height + 20,
                        ResizeMode = ResizeMode.CanResize
                    };
                    break;

                case WindownResize.RESIZE_HEIGHT:
                    w = new Window
                    {
                        MinHeight = height + 20,
                        MinWidth = width + 20,
                        MaxWidth = width + 20,
                        Width = width + 20,
                        Height = height + 20,
                        ResizeMode = ResizeMode.CanResize
                    };
                    break;

                case WindownResize.RESIZE_BOTH:
                    w = new Window
                    {
                        MinHeight = height + 20,
                        MinWidth = width + 20,
                        Width = width + 20,
                        Height = height + 20,
                        ResizeMode = ResizeMode.CanResize
                    };
                    break;
            }
            SetContentForWindow(ref w, ref content, title, parentWindow);
            ConfigWindowStyle(w, false, infoAction);
            content.Background = CurrentModule.BgColor;
            w.Background = CurrentModule.BgColor;
            return w;
        }

        #region Support

        private static System.Windows.Window PrivateGetFixWindow(System.Windows.Controls.UserControl content, string title = "", WindowFit type = WindowFit.FIX, Window parentWindow = null)
        {
            System.Windows.Window w = new System.Windows.Window();
            switch (type)
            {
                case WindowFit.FIX:
                    w.Height = content.Height + 40;
                    w.Width = content.Width + 20;
                    break;

                case WindowFit.WIDTH_CONTENT:
                    w.Height = content.Height + 40;
                    w.SizeToContent = SizeToContent.Width;
                    break;

                case WindowFit.HEGHT_CONTENT:
                    w.Width = content.Width + 20;
                    w.SizeToContent = SizeToContent.Height;
                    break;

                case WindowFit.BOTH_CONTENT:
                    w.SizeToContent = SizeToContent.WidthAndHeight;
                    break;
            }
            w.ResizeMode = ResizeMode.NoResize;
            SetContentForWindow(ref w, ref content, title, parentWindow);
            return w;
        }

        private static System.Windows.Window PrivateGetResizeWindow(System.Windows.Controls.UserControl content, int width, int height, string title = "", WindownResize type = WindownResize.RESIZE_WIDTH, Window parentWindow = null)
        {
            Window w = null;
            switch (type)
            {
                case WindownResize.RESIZE_WIDTH:
                    w = new Window
                    {
                        MinHeight = height + 20,
                        MaxHeight = height + 20,
                        MinWidth = width + 20,
                        Width = width + 20,
                        Height = height + 20,
                        ResizeMode = ResizeMode.CanResize
                    };
                    break;

                case WindownResize.RESIZE_HEIGHT:
                    w = new Window
                    {
                        MinHeight = height + 20,
                        MinWidth = width + 20,
                        MaxWidth = width + 20,
                        Width = width + 20,
                        Height = height + 20,
                        ResizeMode = ResizeMode.CanResize
                    };
                    break;

                case WindownResize.RESIZE_BOTH:
                    w = new Window
                    {
                        MinHeight = height + 20,
                        MinWidth = width + 20,
                        Width = width + 20,
                        Height = height + 20,
                        ResizeMode = ResizeMode.CanResize
                    };
                    break;
            }
            SetContentForWindow(ref w, ref content, title, parentWindow);
            return w;
        }

        private static void SetContentForWindow(ref Window w, ref System.Windows.Controls.UserControl content, string title = "", Window parentWindow = null)
        {
            w.Content = content;
            w.Background = content.Background;
            w.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            if (string.IsNullOrEmpty(title) == false) w.Title = title;
            if (string.IsNullOrEmpty(title) && parentWindow != null)
            {
                w.Title = parentWindow.Title;
                w.Icon = parentWindow.Icon;
            }
            if (parentWindow != null)
            {
                var wih = new WindowInteropHelper(parentWindow);
                IntPtr hWnd = wih.Handle;
                new WindowInteropHelper(w) { Owner = hWnd };
                w.ResizeMode = ResizeMode.NoResize;
            }
        }

        private void InvokeShowWindow(System.Windows.Controls.UserControl userControl,
                                        Func<System.Windows.Controls.UserControl, Window> functionWindow,
                                        string title = "")
        {
            string btnName = userControl.GetType().Namespace;
            var ribbonItems = CurrentModule.RibbonManage.GetButtons(btnName);
            Window = functionWindow(userControl);
            if (ribbonItems?.Count == 1) Window.Title = ribbonItems[0].ItemText.Replace("\r\n", " ");
            if (string.IsNullOrEmpty(title) == false) Window.Title = title;
            WindowHandleSearch handle = WindowHandleSearch.MainWindowHandle;
            handle.SetAsOwner(Window);
            Window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            var RibbonItem = ribbonItems != null && ribbonItems.Count > 0 ? ribbonItems[0] : null;
            if (RibbonItem != null)
            {
                Window.Title = RibbonItem.ItemText.Replace("\r\n", " ");
                if (RibbonItem is PushButton pbt) Window.Icon = pbt.Image;
                if (RibbonItem is SplitButton spl)
                {
                    foreach (var item in spl.GetItems())
                    {
                        if (item is PushButton pb)
                        {
                            if (pb.ClassName.Contains(btnName))
                            {
                                Window.Icon = pb.Image;
                                break;
                            }
                        }
                    }
                }
            }
            Window.Closed += (s, e) =>
            {
                if (ribbonItems != null)
                {
                    foreach (var rbitem in ribbonItems)
                    {
                        rbitem.Enabled = true;
                    }
                }
            };
            Window.KeyDown += (s, e) => { if (e.Key == System.Windows.Input.Key.Escape) Window?.Close(); };
            if (ribbonItems != null)
            {
                foreach (var rbitem in ribbonItems)
                {
                    rbitem.Enabled = false;
                }
            }
        }

        #endregion Support

        #region Config styleWindow

        private static void ConfigWindowStyle(System.Windows.Window w, bool isFixed, Action infoAction)
        {
            // NewStyle
            WindowChrome wc = new WindowChrome
            {
                GlassFrameThickness = new Thickness(1),
                CornerRadius = new CornerRadius(5),
                CaptionHeight = 0
            };
            //w.WindowStyle = WindowStyle.None;
            //w.AllowsTransparency = true;
            w.Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#f0f0f0"));
            WindowChrome.SetWindowChrome(w, wc);
            w.Height = w.Height - 40;
            w.Padding = new Thickness(0);
            var uc = w.Content as UserControl;
            uc.Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#f0f0f0"));
            var dockPanelRoot = uc.Content as DockPanel;
            if (dockPanelRoot == null) return;
            if (dockPanelRoot.Children[0] is Grid gr && gr.Name == _windowTitleName)
            {
                var title = (TextBlock)(ControlsUtils.GetControlByName(dockPanelRoot.Children[0], _TitleId));
                if (title != null) title.Text = w.Title;
                var icon = (Image)(ControlsUtils.GetControlByName(dockPanelRoot.Children[0], _IconId));
                if (icon != null) icon.Source = w.Icon;
                var mini = (Button)(ControlsUtils.GetControlByName(dockPanelRoot.Children[0], _MiniId));
                if (mini != null) mini.Visibility = isFixed ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
                var maxi = (Button)(ControlsUtils.GetControlByName(dockPanelRoot.Children[0], _MaxiId));
                if (maxi != null) maxi.Visibility = isFixed ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
                return;
            }
            // Tạo Grid mới
            double heightTitle = 30;
            Grid grid = new Grid
            {
                Name = _windowTitleName,
                Height = heightTitle,
                Background = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#DADADA")),
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(-10, 0, -10, 0)
            };
            DockPanel.SetDock(grid, Dock.Top);
            grid.MouseLeftButtonDown += (sender, e) => { w.DragMove(); };
            ColumnDefinition col1 = new ColumnDefinition();
            ColumnDefinition col2 = new ColumnDefinition { Width = GridLength.Auto };
            grid.ColumnDefinitions.Add(col1);
            grid.ColumnDefinitions.Add(col2);
            // Grid.Column1
            StackPanel stackPanel1 = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(5, 0, 0, 0)
            };
            Grid.SetColumn(stackPanel1, 0);
            grid.Children.Add(stackPanel1);
            Image image = new Image
            {
                Name = _IconId,
                Width = 16,
                Height = 16,
                Source = w.Icon
            };
            TextBlock textBlock = new TextBlock
            {
                Name = _TitleId,
                Text = w.Title,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.Black
            };
            stackPanel1.Children.Add(image);
            stackPanel1.Children.Add(textBlock);
            // Grid.Colum2
            StackPanel stackPanel2 = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0)
            };
            Grid.SetColumn(stackPanel2, 1);
            grid.Children.Add(stackPanel2);
            // Tạo các Button
            var iconTemp = IconTemplate();
            Thickness tn0 = new Thickness(0);
            Button minimumButton = new Button
            {
                Width = heightTitle,
                Height = heightTitle,
                Content = "M20,14H4V10H20",
                Name = _MiniId,
                Margin = tn0,
                BorderThickness = tn0,
                Template = iconTemp,
            };
            minimumButton.Click += (sender, e) =>
            {
                w.WindowState = WindowState.Minimized;
            };

            Button maxButton = new Button
            {
                Width = heightTitle,
                Height = heightTitle,
                Content = "M4,4H20V20H4V4M6,8V18H18V8H6Z",
                Name = _MaxiId,
                Margin = tn0,
                BorderThickness = tn0,
                Template = iconTemp,
            };
            maxButton.Click += (sender, e) =>
            {
                if (w.WindowState == WindowState.Maximized)
                    w.WindowState = WindowState.Normal;
                else
                    w.WindowState = WindowState.Maximized;
            };
            Button infoButton = new Button
            {
                Width = heightTitle,
                Height = heightTitle,
                //Content = "M11,18H13V16H11V18M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,20C7.59,20 4,16.41 4,12C4,7.59 7.59,4 12,4C16.41,4 20,7.59 20,12C20,16.41 16.41,20 12,20M12,6A4,4 0 0,0 8,10H10A2,2 0 0,1 12,8A2,2 0 0,1 14,10C14,12 11,11.75 11,15H13C13,12.75 16,12.5 16,10A4,4 0 0,0 12,6Z",
                Content = "M11,9H13V7H11M12,20C7.59,20 4,16.41 4,12C4,7.59 7.59,4 12,4C16.41,4 20,7.59 20,12C20,16.41 16.41,20 12,20M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M11,17H13V11H11V17Z",
                Name = "info",
                Margin = tn0,
                BorderThickness = tn0,
                Template = iconTemp,
            };
            infoButton.Click += (sender, e) =>
            {
                infoAction?.Invoke();
            };
            iconTemp = IconCloseTemplate();
            Button closeButton = new Button
            {
                Width = heightTitle,
                Height = heightTitle,
                Content = "M19,6.41L17.59,5L12,10.59L6.41,5L5,6.41L10.59,12L5,17.59L6.41,19L12,13.41L17.59,19L19,17.59L13.41,12L19,6.41Z",
                Margin = tn0,
                BorderThickness = tn0,
                Template = iconTemp,
            };
            WindowChrome.SetIsHitTestVisibleInChrome(infoButton, true);
            WindowChrome.SetIsHitTestVisibleInChrome(minimumButton, true);
            WindowChrome.SetIsHitTestVisibleInChrome(maxButton, true);
            WindowChrome.SetIsHitTestVisibleInChrome(closeButton, true);
            Border border = new Border();
            border.BorderThickness = new Thickness(1, 0, 0, 0);
            border.BorderBrush = Brushes.Gray;
            border.Margin = new Thickness(0, 7, 0, 7);
            closeButton.Click += (sender, e) => { w.Close(); };
            stackPanel2.Children.Add(infoButton);
            //stackPanel2.Children.Add(border);
            stackPanel2.Children.Add(minimumButton);
            stackPanel2.Children.Add(maxButton);
            stackPanel2.Children.Add(closeButton);

            if (isFixed)
            {
                minimumButton.Visibility = System.Windows.Visibility.Collapsed;
                maxButton.Visibility = System.Windows.Visibility.Collapsed;
            }
            //
            dockPanelRoot.Children.Insert(0, grid);
            double ml = dockPanelRoot.Margin.Left;
            double mr = dockPanelRoot.Margin.Right;
            double mt = dockPanelRoot.Margin.Top;
            if (w.ResizeMode == ResizeMode.CanResize)
            {
                var gridbar = (dockPanelRoot.Children[0] as Grid);
                (dockPanelRoot.Children[0] as Grid).Margin = new Thickness(gridbar.Margin.Left, -10, gridbar.Margin.Right, gridbar.Margin.Bottom);
            }
            else
            {
                var widthW = w.Width;
                var contentW = (w.Content as UserControl).Width;
                double marLeft = 0;
                if (!(contentW is double.NaN)) marLeft = (double)(Math.Abs(widthW - contentW) / 2.0);
                dockPanelRoot.Margin = new Thickness(0, 0, 0, 10);
                var gridbar = (dockPanelRoot.Children[0] as Grid);//
                var heightW = w.Height;
                var contentH = (w.Content as UserControl).Height;
                var rendersizeW = w.RenderSize;
                var rendersize = (dockPanelRoot.Children[0] as Grid).RenderSize;
                (dockPanelRoot.Children[0] as Grid).Margin = new Thickness(gridbar.Margin.Left, gridbar.Margin.Top, gridbar.Margin.Right, 10);
            }
        }

        private static ControlTemplate IconTemplate()
        {
            var fillNormal = Brushes.Gray;
            var fillHover = Brushes.Gray;
            var bgHover = (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#c2c2c2");//
            // Tạo ControlTemplate cho Button
            ControlTemplate buttonTemplate = new ControlTemplate(typeof(Button));

            // Tạo Border
            FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "border";
            borderFactory.SetValue(Border.BorderThicknessProperty, new Thickness(0));
            borderFactory.SetValue(Border.BorderBrushProperty, (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#FFB1B1B1"));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(0));
            borderFactory.SetValue(Border.PaddingProperty, new Thickness(0));
            borderFactory.SetValue(Border.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            borderFactory.SetValue(Border.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.SetValue(Border.BackgroundProperty, (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#01FFFFFF"));
            // Tạo Grid
            FrameworkElementFactory gridFactory = new FrameworkElementFactory(typeof(Grid));
            gridFactory.SetValue(Grid.WidthProperty, new System.Windows.Data.Binding("Width")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent),
            });
            gridFactory.SetValue(Grid.HeightProperty, new System.Windows.Data.Binding("Height")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent),
            });
            // Tạo Path
            FrameworkElementFactory pathFactory = new FrameworkElementFactory(typeof(Path));
            pathFactory.Name = "pathFill";
            pathFactory.SetValue(Path.DataProperty, new System.Windows.Data.Binding("Content")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent),
                StringFormat = "{0}" // Định dạng dữ liệu
            });
            pathFactory.SetValue(Path.FillProperty, fillNormal);
            pathFactory.SetValue(Path.RenderTransformOriginProperty, new System.Windows.Point(0.8, 0.8));

            // Tạo và áp dụng ScaleTransform cho Path
            ScaleTransform scaleTransform = new ScaleTransform();
            scaleTransform.ScaleX = 0.8;
            scaleTransform.ScaleY = 0.8;

            pathFactory.SetValue(Path.RenderTransformProperty, scaleTransform);

            // Thêm Path vào Grid
            gridFactory.AppendChild(pathFactory);

            // Thêm Grid vào Border
            borderFactory.AppendChild(gridFactory);

            // Gán Border làm Visual Tree của Button
            buttonTemplate.VisualTree = borderFactory;

            // Thay đổi Background và Fill khi hover
            Trigger hoverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty,
                Value = true
            };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgHover, "border"));
            hoverTrigger.Setters.Add(new Setter(Path.FillProperty, fillHover, "pathFill"));
            buttonTemplate.Triggers.Add(hoverTrigger);

            // Thay đổi Foreground khi bị disable
            Trigger disabledTrigger = new Trigger
            {
                Property = UIElement.IsEnabledProperty,
                Value = false
            };
            disabledTrigger.Setters.Add(new Setter(Button.ForegroundProperty, Brushes.Gray));
            buttonTemplate.Triggers.Add(disabledTrigger);

            return buttonTemplate;
        }

        private static ControlTemplate IconCloseTemplate()
        {
            var fillNormal = Brushes.Gray;// (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#C42B1C");
            var fillHover = Brushes.White;
            var bgHover = (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#C42B1C");
            // Tạo ControlTemplate cho Button
            ControlTemplate buttonTemplate = new ControlTemplate(typeof(Button));

            // Tạo Border
            FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "border";
            borderFactory.SetValue(Border.BorderThicknessProperty, new Thickness(0));
            borderFactory.SetValue(Border.BorderBrushProperty, (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#FFB1B1B1"));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(0));
            borderFactory.SetValue(Border.PaddingProperty, new Thickness(0));
            borderFactory.SetValue(Border.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            borderFactory.SetValue(Border.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.SetValue(Border.BackgroundProperty, (System.Windows.Media.Brush)new BrushConverter().ConvertFrom("#01FFFFFF"));
            // Tạo Grid
            FrameworkElementFactory gridFactory = new FrameworkElementFactory(typeof(Grid));
            gridFactory.SetValue(Grid.WidthProperty, new System.Windows.Data.Binding("Width")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent),
            });
            gridFactory.SetValue(Grid.HeightProperty, new System.Windows.Data.Binding("Height")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent),
            });
            // Tạo Path
            FrameworkElementFactory pathFactory = new FrameworkElementFactory(typeof(Path));
            pathFactory.Name = "pathFill";
            pathFactory.SetValue(Path.DataProperty, new System.Windows.Data.Binding("Content")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent),
                StringFormat = "{0}" // Định dạng dữ liệu
            });
            pathFactory.SetValue(Path.FillProperty, fillNormal);
            pathFactory.SetValue(Path.RenderTransformOriginProperty, new System.Windows.Point(0.8, 0.8));

            // Tạo và áp dụng ScaleTransform cho Path
            ScaleTransform scaleTransform = new ScaleTransform();
            scaleTransform.ScaleX = 0.8;
            scaleTransform.ScaleY = 0.8;

            pathFactory.SetValue(Path.RenderTransformProperty, scaleTransform);

            // Thêm Path vào Grid
            gridFactory.AppendChild(pathFactory);

            // Thêm Grid vào Border
            borderFactory.AppendChild(gridFactory);

            // Gán Border làm Visual Tree của Button
            buttonTemplate.VisualTree = borderFactory;

            // Thay đổi Background và Fill khi hover
            Trigger hoverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty,
                Value = true
            };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgHover, "border"));
            hoverTrigger.Setters.Add(new Setter(Path.FillProperty, fillHover, "pathFill"));
            buttonTemplate.Triggers.Add(hoverTrigger);

            // Thay đổi Foreground khi bị disable
            Trigger disabledTrigger = new Trigger
            {
                Property = UIElement.IsEnabledProperty,
                Value = false
            };
            disabledTrigger.Setters.Add(new Setter(Button.ForegroundProperty, Brushes.Gray));
            buttonTemplate.Triggers.Add(disabledTrigger);

            return buttonTemplate;
        }

        #endregion Config styleWindow
    }
}