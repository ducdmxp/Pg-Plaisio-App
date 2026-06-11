using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace Convert2DTo3D
{
    public class ResourcesAssembly
    {
        public Assembly CurentAssembly { get; set; }
        public string ResourcesFolder { get; set; }
        private string ResourcesFolderComponent { get; set; }
        private string SepFileVersion = "_";

        public ResourcesAssembly(Assembly m_Assembly, string m_ResourcesFolder = "Resources")
        {
            CurentAssembly = m_Assembly;
            ResourcesFolder = (Char.IsDigit(m_ResourcesFolder[0]) ? "_" : "") + m_ResourcesFolder;
            ResourcesFolderComponent = m_ResourcesFolder;
        }

        public string WriteResourceToFile(string OutFullPathfileName)
        {
            try
            {
                FileInfo FileInfo = new FileInfo(OutFullPathfileName);
                FileInfo.Directory.Create();
                string newResourceName = GetEmbededName() + "." + ResourcesFolder + "." + FileInfo.Name;

                using (var resource = CurentAssembly.GetManifestResourceStream(newResourceName))
                {
                    FileStream file;
                    using (file = new FileStream(OutFullPathfileName, FileMode.Create, FileAccess.Write))
                    {
                        resource.CopyTo(file);
                    }
                    file.Close();
                }
                return OutFullPathfileName;
            }
            catch (Exception)
            { }
            return null;
        }

        public string WriteResourceToFileByVersion(string outputfileName)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(outputfileName);
                if (fileInfo.Directory.Exists == false) fileInfo.Directory.Create();
                var fileName = fileInfo.Name;
                var extension = fileInfo.Extension;
                var name = fileName.Replace(extension, "");
                var directory = fileInfo.DirectoryName;
                var fileNameInResources = GetFileByVersion(name);
                fileNameInResources += extension;
                string newResourceName = GetEmbededName() + "." + ResourcesFolder + "." + fileNameInResources;
                string outputPath = System.IO.Path.Combine(directory, name + extension);
                using (var resource = CurentAssembly.GetManifestResourceStream(newResourceName))
                {
                    FileStream file;
                    using (file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        resource.CopyTo(file);
                    }
                    file.Close();
                }
                return outputPath;
            }
            catch (Exception ex)
            { }
            return "";
        }

        private string GetFileByVersion(string fileNameInResources)
        {
#if REVIT2020
 fileNameInResources += SepFileVersion + 2020;
#elif REVIT2021
 fileNameInResources += SepFileVersion + 2021;
#elif REVIT2022
            fileNameInResources += SepFileVersion + 2022;
#elif REVIT2023
                fileNameInResources += SepFileVersion + 2023;
#elif REVIT2024
 fileNameInResources += SepFileVersion + 2024;
#elif REVIT2025
 fileNameInResources += SepFileVersion + 2025;
#else

#endif
            return fileNameInResources;
        }

        public string WriteResourceToFileByVersion(string resourcefileName, string outfileName)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(resourcefileName);
                if (fileInfo.Directory.Exists == false) fileInfo.Directory.Create();
                var fileName = fileInfo.Name;
                var extension = fileInfo.Extension;
                var name = fileName.Replace(extension, "");
                var directory = fileInfo.DirectoryName;
                var fileNameInResources = GetFileByVersion(name);
                fileNameInResources += extension;
                string newResourceName = GetEmbededName() + "." + ResourcesFolder + "." + fileNameInResources;
                string outputPath = System.IO.Path.Combine(directory, outfileName);
                using (var resource = CurentAssembly.GetManifestResourceStream(newResourceName))
                {
                    FileStream file;
                    using (file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        resource.CopyTo(file);
                    }
                    file.Close();
                }
                return outputPath;
            }
            catch (Exception ex)
            { }
            return "";
        }

        public string GetEmbededName()
        {
            return CurentAssembly.GetName().Name;
        }

        public Uri GetUriString(string icon)
        {
            return new Uri("pack://application:,,,/" + GetEmbededName() + ";component/" + ResourcesFolderComponent + "/" + icon);
        }

        public void LoadAssembly(string dllAssembly)
        {
            string pathName = CurentAssembly.GetName().Name + "." + ResourcesFolder + "." + dllAssembly;
            Stream stream = CurentAssembly.GetManifestResourceStream(pathName);
            byte[] result = new byte[stream.Length];
            stream.Read(result, 0, result.Length);
            Assembly.Load(result);
        }

        public string GetCurrentNameSpace()
        {
            return new StackTrace().GetFrame(1).GetMethod().ReflectedType.Namespace;
        }
    }
}