using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Convert2DTo3D.Utils
{
    public class MngShareParameter
    {
        public static bool CreateProjectParamter(Document doc)
        {
            Application app = doc.Application;

            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string sharedParamTempPath = /*@"D:\MyDocument\FlashBIM-1\SHARE PARAMETER.txt";//*/ Path.Combine(assemblyFolder, "SHARE PARAMETER.txt");

            if (!System.IO.File.Exists(sharedParamTempPath))
            {
                Console.WriteLine("Lỗi", "Không tìm thấy file Shared Parameters!");
                return false;
            }

            doc.Application.SharedParametersFilename = sharedParamTempPath;

            DefinitionFile defFile = doc.Application.OpenSharedParameterFile();

            //List<Category> lstCategory = doc.Settings.Categories.Cast<Category>().ToList();

            List<Category> lstCategory = new List<Category>()
            {
                Category.GetCategory(doc, BuiltInCategory.OST_DuctCurves) ,
                Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves) ,
                Category.GetCategory(doc, BuiltInCategory.OST_CableTray) ,
                Category.GetCategory(doc, BuiltInCategory.OST_Conduit) ,
                Category.GetCategory(doc, BuiltInCategory.OST_GenericModel) ,
            };

            DefinitionGroups defGroups = defFile.Groups;
            if (defGroups.Size == 0)
            {
                Console.WriteLine("Lỗi", "Không có Group nào trong file Shared Parameters!");
                return false;
            }

            CategorySet categories = app.Create.NewCategorySet();

            lstCategory.Where(x => IsCategoryAllowsBindingParameters(doc, x)).ToList().ForEach(c => categories.Insert(c));

            //Transaction trans = new Transaction(doc, "Tạo Project Parameter");

            try
            {
                //trans.Start();

                foreach (var item in defGroups)
                {
                    try
                    {
                        foreach (ExternalDefinition selectedParam in item.Definitions)
                        {
                            if (IsProjectParameterExists(doc, selectedParam.Name))
                            {
                                Console.WriteLine("Thông báo", $"Project parameter '{selectedParam.Name}' đã tồn tại trong project!");

                                continue;
                            }

                            Binding binding = app.Create.NewTypeBinding(categories);

                            bool success = doc.ParameterBindings.Insert(selectedParam, binding);
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                //trans.Commit();
            }
            catch (Exception)
            {
                //trans.RollBack();
            }

            return true;
        }

        public static bool IsProjectParameterExists(Document doc, string paramName)
        {
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();

            while (iterator.MoveNext())
            {
                Definition def = iterator.Key;
                if (def.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsCategoryAllowsBindingParameters(Document doc, Category cat)
        {
            try
            {
                if (cat != null && !cat.IsReadOnly && cat.AllowsBoundParameters)
                {
                    return true;
                }
            }
            catch { }
            return false;
        }
    }
}