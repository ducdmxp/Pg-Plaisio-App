using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Convert2DTo3D
{
    public enum TypeControl
    {
        Label,
        Button,
        TextBlock,
        RadioButton,
        TextBox,
        CheckBox,
        GroupBox
    }

    public class ControlsUtils
    {
        public static DependencyObject GetControlByName(DependencyObject dependencyObject, string name)
        {
            return GetAllControlsRecursiveByName(dependencyObject as DependencyObject, name);
        }

        private static DependencyObject GetAllControlsRecursiveByName(DependencyObject parent, string name)
        {
            if (parent == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);

                if (child is Control ctrl)
                {
                    if (ctrl.Name == name) return ctrl;
                }
                else if (child is Image img)
                {
                    if (img.Name == name) return img;
                }
                else if (child is TextBlock tbl)
                {
                    if (tbl.Name == name) return tbl;
                }
                var control = GetAllControlsRecursiveByName(child, name);
                if (control != null) return control;
            }
            if (parent is GroupBox gr)
            {
                if (gr.Name == name) return gr;
                var control = GetAllControlsRecursiveByName(gr.Content as DependencyObject, name);
                if (control != null) return control;
            }
            else if (parent is HeaderedContentControl headeredContentControl)
            {
                if (headeredContentControl.Name == name) return headeredContentControl;
                var control = GetAllControlsRecursiveByName(headeredContentControl.Content as DependencyObject, name);
                if (control != null) return control;
            }
            else if (parent is Image img)
            {
                if (img.Name == name) return img;
            }
            if (parent is TabControl tab)
            {
                if (tab.Name == name) return tab;
                foreach (var item in tab.Items)
                {
                    var control = GetAllControlsRecursiveByName(item as DependencyObject, name);
                    if (control != null) return control;
                }
            }
            if (parent is DataGrid dgv)
            {
                if (dgv.Name == name) return dgv;
            }
            return null;
        }

        public static List<DependencyObject> GetControls(DependencyObject dependencyObject, TypeControl typeControl)
        {
            var controls = new List<DependencyObject>();
            GetAllControlsRecursive(dependencyObject as DependencyObject, ref controls);
            return controls;
        }

        public static void Translate(UserControl userControl, IEnumerable<TypeControl> types)
        {
            if (types.Count() < 0) return;
            var controls = new List<DependencyObject>();
            GetAllControlsRecursive(userControl.Content as DependencyObject, ref controls);
            string lstTypes = string.Join("_", types.Select(k => k.ToString()).ToArray());
            foreach (var control in controls)
            {
                string nameType = control.GetType().Name;
                if (lstTypes.Contains(nameType)) TranslateControl(control);
            }
        }

        public static void Translate(UserControl userControl)
        {
            var controls = new List<DependencyObject>();
            GetAllControlsRecursive(userControl.Content as DependencyObject, ref controls);
            foreach (var control in controls)
            {
                TranslateControl(control);
            }
        }

        private static void TranslateControl(DependencyObject control)
        {
            try
            {
                var stringTranslate = "";
                if (control is GroupBox gr)
                {
                    stringTranslate = MngLanguage.GetNameFromCurrentResource(gr.Header.ToString());
                    gr.Header = string.IsNullOrEmpty(stringTranslate) == true ? gr.Header.ToString() : stringTranslate;
                }
                if (control is HeaderedContentControl headeredContentControl)
                {
                    stringTranslate = MngLanguage.GetNameFromCurrentResource(headeredContentControl.Header.ToString());
                    headeredContentControl.Header = string.IsNullOrEmpty(stringTranslate) == true ? headeredContentControl.Header.ToString() : stringTranslate;
                }
                else if (control is ContentControl ctc)
                {
                    if (ctc.Content == null) return;
                    stringTranslate = MngLanguage.GetNameFromCurrentResource(ctc.Content.ToString());
                    ctc.Content = string.IsNullOrEmpty(stringTranslate) == true ? ctc.Content : stringTranslate;
                }
                else if (control is TextBlock tb)
                {
                    stringTranslate = MngLanguage.GetNameFromCurrentResource(tb.Text);
                    tb.Text = string.IsNullOrEmpty(stringTranslate) == true ? tb.Text : stringTranslate;
                }
                else if (control is DataGridColumn dataGridColumn)
                {
                    if (dataGridColumn.Header == null) return;
                    stringTranslate = MngLanguage.GetNameFromCurrentResource(dataGridColumn.Header.ToString());
                    dataGridColumn.Header = string.IsNullOrEmpty(stringTranslate) == true ? dataGridColumn.Header.ToString() : stringTranslate;
                }
            }
            catch (Exception)
            { }
        }

        private static void GetAllControlsRecursive(DependencyObject parent, ref List<DependencyObject> controls)
        {
            if (parent == null) return;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is Control)
                {
                    controls.Add(child as Control);
                }
                else if (child is TextBlock)
                {
                    controls.Add(child as TextBlock);
                }
                GetAllControlsRecursive(child, ref controls);
            }
            if (parent is GroupBox gr)
            {
                controls.Add(gr);
                GetAllControlsRecursive(gr.Content as DependencyObject, ref controls);
            }
            else if (parent is HeaderedContentControl headeredContentControl)
            {
                controls.Add(headeredContentControl);
                GetAllControlsRecursive(headeredContentControl.Content as DependencyObject, ref controls);
            }
            if (parent is TabControl tab)
            {
                foreach (var item in tab.Items) GetAllControlsRecursive(item as DependencyObject, ref controls);
            }
            if (parent is DataGrid dgv)
            {
                foreach (var item in dgv.Columns) controls.Add(item);
            }
        }

        public static bool Valid(System.Windows.Controls.TextBox control)
        {
            return !(Validation.GetHasError(control));
        }
    }
}