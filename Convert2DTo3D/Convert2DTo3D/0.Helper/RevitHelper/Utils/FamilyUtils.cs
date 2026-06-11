using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Convert2DTo3D
{
    public class FamilyUtils
    {
        private static string ResourceFolder = "Resources";

        public static Family LoadFamily(Document doc, string path, bool clear)
        {
            Family result = null;
            FileInfo fileInfo = new FileInfo(path);
            if (fileInfo.Exists == false) return null;
            try
            {
                doc.LoadFamily(path, out result);
            }
            catch (Exception ex)
            { }
            if (clear == false) return result;
            fileInfo.Delete();
            if (fileInfo.Directory.GetFiles()?.Count() <= 0 && fileInfo.Directory.GetDirectories()?.Count() <= 0) fileInfo.Directory.Delete();
            return result;
        }

        public static Family LoadFamilyFromResource(Document doc, string fileInResource, bool clear)
        {
            var directDll = GetDirectoryDll();
            var pathFile = System.IO.Path.Combine(directDll, ResourceFolder, fileInResource);
            var outPathFile = CurrentModule.ResourcesAssemblyMng.WriteResourceToFile(pathFile);
            if (string.IsNullOrEmpty(outPathFile) == true) return null;
            FileInfo fileInfo = new FileInfo(outPathFile);
            var result = LoadFamily(doc, outPathFile, clear);
            return null;
        }

        public static Family LoadFamilyFromResourceByVersion(Document doc, string fileInResource, bool clear)
        {
            var directDll = GetDirectoryDll();
            var pathFile = System.IO.Path.Combine(directDll, ResourceFolder, fileInResource);
            var outPathFile = CurrentModule.ResourcesAssemblyMng.WriteResourceToFileByVersion(pathFile);
            if (string.IsNullOrEmpty(outPathFile) == true) return null;
            FileInfo fileInfo = new FileInfo(outPathFile);
            var result = LoadFamily(doc, outPathFile, clear);
            return result;
        }

        public static Family LoadFamilyFromResourceByVersion(Document doc, string fileInResource, string fileOut, bool clear)
        {
            var directDll = GetDirectoryDll();
            var pathFile = System.IO.Path.Combine(directDll, ResourceFolder, fileInResource);
            var outPathFile = CurrentModule.ResourcesAssemblyMng.WriteResourceToFileByVersion(pathFile, fileOut);
            if (string.IsNullOrEmpty(outPathFile) == true) return null;
            FileInfo fileInfo = new FileInfo(outPathFile);
            var result = LoadFamily(doc, outPathFile, clear);
            return result;
        }

        private static string GetDirectoryDll()
        {
            var assembly = Assembly.GetExecutingAssembly();
            return Path.GetDirectoryName(assembly.Location);
        }

        public static bool IsExistFamily(Document doc, string name)
        {
            var family = new FilteredElementCollector(doc)
                .OfClass(typeof(Family)).Cast<Family>()
                .FirstOrDefault(x => x.Name.Equals(name));
            return family != null;
        }
    }
}