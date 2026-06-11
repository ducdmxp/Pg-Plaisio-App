using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D
{
    public class MngProjectParameter
    {
        private Document doc { get; set; }

        public MngProjectParameter(Document m_doc)
        {
            doc = m_doc;
        }

        /// <summary>
        /// Lấy ra tất cả các parameter trong project parameter.
        /// </summary>
        /// <returns></returns>
        public List<Definition> GetAllParameters()
        {
            List<Definition> lst = new List<Definition>();
            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                if (it.Key != null)
                {
                    lst.Add(it.Key);
                }
            }
            return lst;
        }

        /// <summary>
        /// Lấy ra parameter trong shareparameter dựa vào name truyền vào.
        /// </summary>
        /// <param name="PramName"></param>
        /// <returns></returns>
        public Definition GetParameter(string PramName)
        {
            Definition def = GetAllParameters().Where(q => q.Name == PramName).FirstOrDefault();
            return def;
        }

        /// <summary>
        /// Lấy ra Binding của parameter trong project paramter theo tên của parameter.
        /// </summary>
        /// <param name="pramName"></param>
        /// <returns></returns>
        public Binding GetBindingParameter(string pramName)
        {
            Definition def = GetParameter(pramName);
            return def != null ? doc.ParameterBindings.get_Item(def) : null;
        }

#if REVIT2019 || REVIT2020 || REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024

        /// <summary>
        /// Thêm mới 1 parameter
        /// </summary>
        /// <param name="def">Parameter definition trong share parameter.</param>
        /// <param name="group">Group của project paramter</param>
        /// <param name="isInstance">Instance hoặc type của paramter</param>
        /// <param name="catset">Danh sách Categories</param>
        /// <param name="IsSetVaryByGroupInstance"></param>
        public void AddParameter(Definition def, BuiltInParameterGroup group, bool isInstance = true, CategorySet catset = null, bool IsSetVaryByGroupInstance = false)
        {
            Binding binding = doc.ParameterBindings.get_Item(def);
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
            if (binding == null)
            {
                CategorySet cats = catset;
                if (cats == null)
                {
                    cats = new CategorySet();
                    cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralColumns));
                }
                ;
                if (isInstance)
                {
                    InstanceBinding ins = app.Create.NewInstanceBinding(cats);
                    doc.ParameterBindings.Insert(def, ins, group);
                }
                else
                {
                    TypeBinding ty = app.Create.NewTypeBinding(cats);
                    doc.ParameterBindings.Insert(def, ty, group);
                }
                if (IsSetVaryByGroupInstance == true) SetVaryByGroupInstance(def);
            }
        }

        /// <summary>
        /// Chú ý: Khi sử dụng doc.ParameterBindings.ReInsert Categories thì dữ liệu SetVaryByGroupInstance sẽ trở lại là False.
        /// Nên khi sử dụng phương thức này ta cần sử dụng lại phương thức này để set SetVaryByGroupInstance
        /// </summary>
        /// <param name="defItem"></param>
        /// <param name="cat"></param>
        /// <param name="group"></param>
        public void AddParameter(Definition defItem, Category cat, BuiltInParameterGroup group)
        {
            Binding binding = doc.ParameterBindings.get_Item(defItem);
            if (binding != null)
            {
                ElementBinding elebing = binding as ElementBinding;
                CategorySet cats = elebing.Categories;
                if (cats.IsEmpty || cats.Contains(cat) == false)
                {
                    cats.Insert(cat);
                    elebing.Categories = cats;
                    doc.ParameterBindings.ReInsert(defItem, elebing, group);
                }
            }
        }

#endif

        /// <summary>
        /// Chú ý: Khi sử dụng doc.ParameterBindings.ReInsert Categories thì dữ liệu SetVaryByGroupInstance sẽ trở lại là False.
        /// Nên khi sử dụng phương thức này ta cần sử dụng lại phương thức này để set SetVaryByGroupInstance
        /// </summary>
        /// <param name="def"></param>
        public void SetVaryByGroupInstance(Definition def)
        {
            SetVaryByGroupInstance(def.Name);
        }

        /// <summary>
        /// Chú ý: Khi sử dụng doc.ParameterBindings.ReInsert Categories thì dữ liệu SetVaryByGroupInstance sẽ trở lại là False.
        /// Nên khi sử dụng phương thức này ta cần sử dụng lại phương thức này để set SetVaryByGroupInstance
        /// </summary>
        /// <param name="defName"></param>
        public void SetVaryByGroupInstance(string defName)
        {
            Definition bind = GetParameter(defName);
            if (bind == null) return;
            InternalDefinition def = bind as InternalDefinition;
            if (def == null) return;

            if (def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.Text) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.Area) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.Volume) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.Currency) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.MassDensity) ||
#if !(BUILD2019)
                // def.GetParameterType() == ConvertParameterType.Convert(MapParameterType.Speed) ||
#endif
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.URL) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.Material) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.Image) ||
                def.GetParameterType() == ConvertParameterType.Convert(ParameterTypeMap.MultilineText))
            {
                def.SetAllowVaryBetweenGroups(doc, true);
            }
        }

        /// <summary>
        /// Xoá parameter trong project parameter.
        /// </summary>
        /// <param name="def"></param>
        public void RemoveParameter(Definition def)
        {
            BindingMap map = doc.ParameterBindings;
            InternalDefinition Def = def as InternalDefinition;
            if (Def != null)
            {
                map.Remove(Def);
            }
        }

#if REVIT2020 || REVIT2021

        /// <summary>
        /// Kiểm tra có thể tạo parameter trong project parameter với các tham số này không.
        /// </summary>
        /// <param name="nameParameter">Tên của Parameter</param>
        /// <param name="groupName">Tên của GroupName trong Share Paramter</param>
        /// <param name="parameterTypeMap">kiểu dữ liệu của parameter.</param>
        /// <param name="builtInParameterGroup">Nhóm của paramter trong project ProjectParameter.</param>
        /// <param name="isSetVaryByGroupInstance">Có set VaryByGroupInstance</param>
        /// <param name="categories"></param>
        /// <param name="status">Nếu trả về là Equal tức là parameter đã tạo trong project paramter với các tham số này rồi, không cần phải tạo lại.</param>
        /// <returns></returns>
        public bool CanCreateParameterForMap(string nameParameter, ParameterTypeMap parameterTypeMap,
            BuiltInParameterGroup builtInParameterGroup, bool isInstance, bool isSetVaryByGroupInstance, List<BuiltInCategory> categories, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            InternalDefinition InDef = GetParameter(nameParameter) as InternalDefinition;
            if (InDef == null) return true;
            if (InDef.GetParameterType() != ConvertParameterType.Convert(parameterTypeMap))
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            if (InDef.ParameterGroup != builtInParameterGroup)
            {
                status = ParameterWaringStatus.Different_GroupProject;
                return false;
            }
            var bindingParameter = GetBindingParameter(nameParameter);
            if (bindingParameter == null)
            {
                status = ParameterWaringStatus.Different_InstanceOrType;
                return false;
            }
            else
            {
                if ((bindingParameter is InstanceBinding && isInstance == false) || (bindingParameter is TypeBinding && isInstance == true))
                {
                    status = ParameterWaringStatus.Different_InstanceOrType;
                    return false;
                }
            }

            if (InDef.VariesAcrossGroups != isSetVaryByGroupInstance)
            {
                status = ParameterWaringStatus.Different_VaryByGroupInstance;
                return false;
            }
            try
            {
                Binding binding = doc.ParameterBindings.get_Item(InDef);
                ElementBinding elebing = binding as ElementBinding;
                CategorySet cats = elebing.Categories;
                string categoriesNewStr = string.Join(",", categories.Select(k => (int)k).OrderBy(k => k).ToList());
                string categoriesOldStr = string.Join(",", cats.Cast<Category>().Select(k => (int)k.Id.IntegerValue).OrderBy(k => k).ToList());
                if (categoriesNewStr == categoriesOldStr)
                {
                    status = ParameterWaringStatus.Equal;
                    return false;
                }
            }
            catch (Exception)
            { }
            status = ParameterWaringStatus.Different_Categories;
            return false;
        }
        public bool CanCreateParameter(string nameParameter, ParameterType parameterTypeMap, BuiltInParameterGroup builtInParameterGroup, bool isInstance, bool isSetVaryByGroupInstance, List<BuiltInCategory> categories, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            InternalDefinition InDef = GetParameter(nameParameter) as InternalDefinition;
            if (InDef == null) return true;
            if (InDef.GetParameterType() != parameterTypeMap)
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            if (InDef.ParameterGroup != builtInParameterGroup)
            {
                status = ParameterWaringStatus.Different_GroupProject;
                return false;
            }
            var bindingParameter = GetBindingParameter(nameParameter);
            if (bindingParameter == null)
            {
                status = ParameterWaringStatus.Different_InstanceOrType;
                return false;
            }
            else
            {
                if ((bindingParameter is InstanceBinding && isInstance==false ) || (bindingParameter is TypeBinding && isInstance == true))
                {
                    status = ParameterWaringStatus.Different_InstanceOrType;
                    return false;
                }
            }

            if (InDef.VariesAcrossGroups != isSetVaryByGroupInstance)
            {
                status = ParameterWaringStatus.Different_VaryByGroupInstance;
                return false;
            }
            try
            {
                Binding binding = doc.ParameterBindings.get_Item(InDef);
                ElementBinding elebing = binding as ElementBinding;
                CategorySet cats = elebing.Categories;
                string categoriesNewStr = string.Join(",", categories.Select(k => (int)k).OrderBy(k => k).ToList());
                string categoriesOldStr = string.Join(",", cats.Cast<Category>().Select(k => (int)k.Id.IntegerValue).OrderBy(k => k).ToList());
                if (categoriesNewStr == categoriesOldStr)
                {
                    status = ParameterWaringStatus.Equal;
                    return false;
                }
            }
            catch (Exception)
            { }
            status = ParameterWaringStatus.Different_Categories;
            return false;
        }
#else

        public void AddParameter(Definition def, ForgeTypeId group, bool isInstance = true, CategorySet catset = null, bool IsSetVaryByGroupInstance = false)
        {
            Binding binding = doc.ParameterBindings.get_Item(def);
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
            if (binding == null)
            {
                CategorySet cats = catset;
                if (cats == null)
                {
                    cats = new CategorySet();
                    cats.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralColumns));
                }
                ;
                if (isInstance)
                {
                    InstanceBinding ins = app.Create.NewInstanceBinding(cats);
                    doc.ParameterBindings.Insert(def, ins, group);
                }
                else
                {
                    TypeBinding ty = app.Create.NewTypeBinding(cats);
                    doc.ParameterBindings.Insert(def, ty, group);
                }
                if (IsSetVaryByGroupInstance == true) SetVaryByGroupInstance(def);
            }
        }

        public bool CanCreateParameterForMap(string nameParameter, string groupName, ParameterTypeMap parameterTypeMap, ForgeTypeId builtInParameterGroup, bool isInstance, bool isSetVaryByGroupInstance, List<ForgeTypeId> categories, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            InternalDefinition InDef = GetParameter(nameParameter) as InternalDefinition;
            if (InDef == null) return true;
            if (InDef.GetParameterType() != ConvertParameterType.Convert(parameterTypeMap))
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            if (InDef.GetParameterGroup() != builtInParameterGroup)
            {
                status = ParameterWaringStatus.Different_GroupProject;
                return false;
            }
            var bindingParameter = GetBindingParameter(nameParameter);
            if (bindingParameter == null)
            {
                status = ParameterWaringStatus.Different_InstanceOrType;
                return false;
            }
            else
            {
                if ((bindingParameter is InstanceBinding && isInstance == false) || (bindingParameter is TypeBinding && isInstance == true))
                {
                    status = ParameterWaringStatus.Different_InstanceOrType;
                    return false;
                }
            }

            if (InDef.VariesAcrossGroups != isSetVaryByGroupInstance)
            {
                status = ParameterWaringStatus.Different_VaryByGroupInstance;
                return false;
            }
            try
            {
                Binding binding = doc.ParameterBindings.get_Item(InDef);
                ElementBinding elebing = binding as ElementBinding;
                CategorySet cats = elebing.Categories;
                string categoriesNewStr = string.Join(",", categories.Select(k => k.TypeId.ToString()).OrderBy(k => k).ToList());
                string categoriesOldStr = string.Join(",", cats.Cast<Category>().Select(k => (int)k.Id.IntegerValue).OrderBy(k => k).ToList());
                if (categoriesNewStr == categoriesOldStr)
                {
                    status = ParameterWaringStatus.Equal;
                    return false;
                }
            }
            catch (Exception)
            { }
            status = ParameterWaringStatus.Different_Categories;
            return false;
        }

        public bool CanCreateParameter(string nameParameter, ForgeTypeId parameterTypeMap, ForgeTypeId builtInParameterGroup, bool isInstance, bool isSetVaryByGroupInstance, List<ForgeTypeId> categories, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            InternalDefinition InDef = GetParameter(nameParameter) as InternalDefinition;
            if (InDef == null) return true;
            if (InDef.GetParameterType() != parameterTypeMap)
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            if (InDef.GetParameterGroup() != builtInParameterGroup)
            {
                status = ParameterWaringStatus.Different_GroupProject;
                return false;
            }
            var bindingParameter = GetBindingParameter(nameParameter);
            if (bindingParameter == null)
            {
                status = ParameterWaringStatus.Different_InstanceOrType;
                return false;
            }
            else
            {
                if ((bindingParameter is InstanceBinding && isInstance == false) || (bindingParameter is TypeBinding && isInstance == true))
                {
                    status = ParameterWaringStatus.Different_InstanceOrType;
                    return false;
                }
            }

            if (InDef.VariesAcrossGroups != isSetVaryByGroupInstance)
            {
                status = ParameterWaringStatus.Different_VaryByGroupInstance;
                return false;
            }
            try
            {
                Binding binding = doc.ParameterBindings.get_Item(InDef);
                ElementBinding elebing = binding as ElementBinding;
                CategorySet cats = elebing.Categories;
                string categoriesNewStr = string.Join(",", categories.Select(k => k.TypeId.ToString()).OrderBy(k => k).ToList());
                string categoriesOldStr = string.Join(",", cats.Cast<Category>().Select(k => (int)k.Id.IntegerValue).OrderBy(k => k).ToList());
                if (categoriesNewStr == categoriesOldStr)
                {
                    status = ParameterWaringStatus.Equal;
                    return false;
                }
            }
            catch (Exception)
            { }
            status = ParameterWaringStatus.Different_Categories;
            return false;
        }

        public bool CanCreateParameter(string nameParameter, ForgeTypeId parameterTypeMap, ForgeTypeId builtInParameterGroup, bool isInstance, bool isSetVaryByGroupInstance, List<BuiltInCategory> categories, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            InternalDefinition InDef = GetParameter(nameParameter) as InternalDefinition;
            if (InDef == null) return true;
            if (InDef.GetParameterType() != parameterTypeMap)
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            if (InDef.GetParameterGroup() != builtInParameterGroup)
            {
                status = ParameterWaringStatus.Different_GroupProject;
                return false;
            }
            var bindingParameter = GetBindingParameter(nameParameter);
            if (bindingParameter == null)
            {
                status = ParameterWaringStatus.Different_InstanceOrType;
                return false;
            }
            else
            {
                if ((bindingParameter is InstanceBinding && isInstance == false) || (bindingParameter is TypeBinding && isInstance == true))
                {
                    status = ParameterWaringStatus.Different_InstanceOrType;
                    return false;
                }
            }

            if (InDef.VariesAcrossGroups != isSetVaryByGroupInstance)
            {
                status = ParameterWaringStatus.Different_VaryByGroupInstance;
                return false;
            }
            try
            {
                Binding binding = doc.ParameterBindings.get_Item(InDef);
                ElementBinding elebing = binding as ElementBinding;
                CategorySet cats = elebing.Categories;
                string categoriesNewStr = string.Join(",", categories.Select(k => (int)k).OrderBy(k => k).ToList());
                string categoriesOldStr = string.Join(",", cats.Cast<Category>().Select(k => (int)k.Id.IntegerValue).OrderBy(k => k).ToList());
                if (categoriesNewStr == categoriesOldStr)
                {
                    status = ParameterWaringStatus.Equal;
                    return false;
                }
            }
            catch (Exception)
            { }
            status = ParameterWaringStatus.Different_Categories;
            return false;
        }

#endif
    }
}