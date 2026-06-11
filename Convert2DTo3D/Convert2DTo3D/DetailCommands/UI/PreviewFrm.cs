using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Convert2DTo3D.UI
{
    public partial class PreviewFrm : System.Windows.Forms.Form
    {
        private UIApplication m_uiApp;

        public PreviewFrm(UIApplication uiApp)
        {
            InitializeComponent();
            m_uiApp = uiApp;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Revit Files (*.rvt; *.rte; *.rfa; *.rft)|*.rvt; *.rte; *.rfa; *.rft";
            ofd.Multiselect = true;
            ofd.Title = "Select Revit File";
            if (ofd.ShowDialog() == DialogResult.Cancel)
                return;

            GetFamilyPreview(ofd.FileNames.FirstOrDefault(), picPreview1);
            GetFamilyPreview(ofd.FileNames.LastOrDefault(), picPreview2);
        }

        public void GetFamilyPreview(string familyFilePath, PictureBox pictureBox)
        {
            var app = m_uiApp.Application;

            // Mở file family tạm thời
            Document familyDoc = app.OpenDocumentFile(familyFilePath);

            try
            {
                if (familyDoc != null && familyDoc.IsFamilyDocument)
                {
                    ElementId viewId = familyDoc.GetDocumentPreviewSettings().PreviewViewId;

                    Autodesk.Revit.DB.View view = familyDoc.GetElement(viewId) as Autodesk.Revit.DB.View;

                    string fileName = familyDoc.Title + GetDateTimeNow();

                    string folderPath = @"C:\Users\admin\Downloads\Place Utility\Image";

                    string filePath = System.IO.Path.Combine(folderPath, fileName + ".png");

                    ImageExportOptions options = new ImageExportOptions
                    {
                        FilePath = filePath,
                        FitDirection = FitDirectionType.Horizontal,
                        HLRandWFViewsFileType = ImageFileType.PNG,
                        ShadowViewsFileType = ImageFileType.PNG,
                        ImageResolution = ImageResolution.DPI_600,
                        ExportRange = ExportRange.SetOfViews,
                        ZoomType = ZoomFitType.FitToPage
                    };

                    options.SetViewsAndSheets(new List<ElementId> { view.Id });

                    familyDoc.ExportImage(options);

                    string path = GetPngFile(folderPath, fileName);

                    pictureBox.Image = System.Drawing.Image.FromFile(path);

                    //PreviewControl preview = elemHost.Child as PreviewControl;
                    //if (preview != null)
                    //    preview.Dispose();

                    //preview = new PreviewControl(familyDoc, viewId);
                    //elemHost.Child = preview;
                    //elemHost.Child.Visibility = System.Windows.Visibility.Visible;
                    //elemHost.Enabled = true;
                }
            }
            finally
            {
                // Đóng tài liệu sau khi xử lý xong
                if (familyDoc != null)
                    familyDoc.Close(false);
            }
        }

        public static string GetPngFile(string folderPath, string nameImage)
        {
            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException($"Folder not found: {folderPath}");

            var pngFiles = Directory.GetFiles(folderPath, "*.png", SearchOption.TopDirectoryOnly);

            foreach (var item in pngFiles)
            {
                string name = Path.GetFileNameWithoutExtension(item);

                if (name.Contains(nameImage))
                    return item;
            }

            return string.Empty;
        }

        public static string GetDateTimeNow()
        {
            DateTime dateTime = DateTime.Now;

            return dateTime.Hour.ToString() + dateTime.Minute.ToString() + dateTime.Second.ToString() +
             dateTime.Millisecond.ToString();
        }

        private void PreviewFrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            PreviewControl preview1 = elemHost1.Child as PreviewControl;
            if (preview1 != null)
                preview1.Dispose();

            PreviewControl preview2 = elemHost2.Child as PreviewControl;
            if (preview2 != null)
                preview2.Dispose();
        }
    }
}