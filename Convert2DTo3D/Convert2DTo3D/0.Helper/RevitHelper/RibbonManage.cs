using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace Convert2DTo3D
{
    public class RibbonManage
    {
        private UIControlledApplication m_uIControlledApplication;
        private Dictionary<string, Dictionary<string, RibbonPanel>> m_panels = null;
        private Assembly m_assemblyCall = null;

        public RibbonManage(UIControlledApplication uIControlledApplication)
        {
            m_uIControlledApplication = uIControlledApplication;
            m_assemblyCall = Assembly.GetCallingAssembly();
            m_panels = new Dictionary<string, Dictionary<string, RibbonPanel>>();
        }

        public void AddRibbonTab(params string[] tabs)
        {
            foreach (var item in tabs)
            {
                try
                {
                    m_uIControlledApplication.CreateRibbonTab(item);
                }
                catch
                {
                }
                m_panels.Add(item, new Dictionary<string, RibbonPanel>());
            }
        }

        public RibbonPanel RibbonPanel(string ribbonTab, string ribbonPanel)
        {
            try
            {
                RibbonPanel panel = m_uIControlledApplication.CreateRibbonPanel(ribbonTab, ribbonPanel);
            }
            catch { }
            List<RibbonPanel> panels = m_uIControlledApplication.GetRibbonPanels(ribbonTab);
            foreach (RibbonPanel p in panels)
            {
                if (p.Name == ribbonPanel)
                {
                    m_panels[ribbonTab].Add(ribbonPanel, p);
                    return p;
                }
            }
            return null;
        }

        public PushButtonData GetPushButtonFromTemplate(string namespaceName,
                                                        string text,
                                                        string image,
                                                        string imageLarge,
                                                        bool reverseImage = false,
                                                        string toolTip = "")
        {
            string executableLocation = Path.GetDirectoryName(m_assemblyCall.Location);
            string dllLocation = Path.Combine(executableLocation, m_assemblyCall.ManifestModule.Name);
            var namespaceEcute = m_assemblyCall.GetName().Name;
            PushButtonData buttondata;
            string pathResources = "pack://application:,,,/{0};component/Resources/{1}";
            Uri img = new Uri(String.Format(pathResources, namespaceEcute, image));
            Uri largeImg = new Uri(String.Format(pathResources, namespaceEcute, imageLarge));
            if (reverseImage == true)
            {
                img = new Uri(String.Format(pathResources, namespaceEcute, imageLarge));
                largeImg = new Uri(String.Format(pathResources, namespaceEcute, imageLarge));
            }
            buttondata = new PushButtonData(namespaceName, text, dllLocation, $"{namespaceName}.{namespaceName}Command");
            buttondata.ToolTip = toolTip;
            buttondata.Image = new BitmapImage(img);
            buttondata.LargeImage = new BitmapImage(largeImg);
            buttondata.AvailabilityClassName = $"{namespaceName}.{namespaceName}Available";
            return buttondata;
        }

        public List<RibbonItem> GetButtons(string value)
        {
            List<RibbonItem> result = null;
            foreach (var _tabName in m_panels.Keys)
            {
                var ribbon = GetButton(_tabName, value);
                if (ribbon != null)
                {
                    if (result == null) result = new List<RibbonItem>();
                    result.Add(ribbon);
                }
            }
            return result;
        }

        private RibbonItem GetButton(string _tabName, string value)
        {
            List<RibbonPanel> panels = m_uIControlledApplication.GetRibbonPanels(_tabName);
            return GetButtonInRibbonPanels(panels, value);
        }

        private RibbonItem GetButtonInRibbonPanels(List<RibbonPanel> panels, string value)
        {
            foreach (RibbonPanel p in panels)
            {
                var ribbonItem = GetButtonInRibbonPanel(p, value);
                if (ribbonItem != null) return ribbonItem;
            }
            return null;
        }

        private RibbonItem GetButtonInRibbonPanel(RibbonPanel p, string value)
        {
            IList<RibbonItem> lstItem = p.GetItems();
            foreach (var item in lstItem)
            {
                if (CheckPushButton(item, value) == true)
                {
                    return item;
                }
                else
                {
                    if (item is PulldownButton pullButton)
                    {
                        foreach (var push in pullButton.GetItems())
                        {
                            if (CheckPushButton(push, value) == true)
                            {
                                return push;
                            }
                        }
                    }
                    else
                    {
                        if (item.Name == value)
                        {
                            return item;
                        }
                    }
                }
            }
            return null;
        }

        private bool CheckPushButton(RibbonItem ribbonItem, string value)
        {
            if (ribbonItem is PushButton pushButton && pushButton.ClassName.Contains(value) == true)
            {
                return true;
            }
            return false;
        }
    }
}