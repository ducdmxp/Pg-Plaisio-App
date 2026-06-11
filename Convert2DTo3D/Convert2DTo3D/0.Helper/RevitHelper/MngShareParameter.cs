using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Convert2DTo3D
{
    public class MngShareParameter
    {
        private Autodesk.Revit.ApplicationServices.Application m_app { get; set; }
        public DefinitionFile ShareParameterFile { get; private set; }
        private string _PathFileOld;

        public string PathFileOld
        {
            get { return _PathFileOld; }
            private set { _PathFileOld = value; }
        }

        private string _PathFileNew;

        public string PathFileNew
        {
            get { return _PathFileNew; }
            set { _PathFileNew = value; }
        }

        private bool m_existFile = true;

        public MngShareParameter(Autodesk.Revit.ApplicationServices.Application app)
        {
            this.m_app = app;
            string pathShareParameter = app.SharedParametersFilename;
            if (string.IsNullOrEmpty(pathShareParameter) == true ||
                File.Exists(pathShareParameter) == false ||
               (File.Exists(pathShareParameter) == true && IsReadOnlyFile(pathShareParameter)))
            {
                m_existFile = false;
            }
            ShareParameterFile = app.OpenSharedParameterFile();
        }

        /// <summary>
        /// Khởi tạo đối tượng quản lý file share parameter
        /// </summary>
        /// <param name="app">Application của Revit</param>
        /// <param name="outputFileName">Đường dẫn đến file shareparameter khi tạo mới</param>
        /// <param name="isForceCreateShareFile">Nếu là true thì luôn luôn tạo mới file shareparameter ngược lại nó sẽ ghi đè lên file shareparameter cũ.</param>
        public MngShareParameter(Autodesk.Revit.ApplicationServices.Application app, string outputFileName, bool isForceCreateShareFile)
        {
            this.m_app = app;
            string pathShareParameter = app.SharedParametersFilename;
            bool isCreateFile = false;
            if (string.IsNullOrEmpty(pathShareParameter) == true ||
                File.Exists(pathShareParameter) == false ||
               (File.Exists(pathShareParameter) == true && IsReadOnlyFile(pathShareParameter)))
            {
                isCreateFile = true;
            }
            if (isForceCreateShareFile) isCreateFile = true;
            if (isCreateFile)
            {
                CreateDefaultShareParameter(outputFileName);
                app.SharedParametersFilename = outputFileName;
                PathFileNew = outputFileName;
            }
            PathFileOld = string.IsNullOrEmpty(pathShareParameter) == false || File.Exists(pathShareParameter) == true ? pathShareParameter : "";
            ShareParameterFile = app.OpenSharedParameterFile();
        }

        /// <summary>
        /// Kiểm tra file shareparameter có cho phép ghi hay không
        /// </summary>
        /// <param name="filePath">Đường dẫn để kiểm tra file shareparameter.</param>
        /// <returns></returns>
        private bool IsReadOnlyFile(string filePath)
        {
            if (!File.Exists(filePath)) return true;
            try
            {
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Write))
                {
                    return false;
                }
            }
            catch (Exception)
            { }
            return true;
        }

        /// <summary>
        /// Tạo mới một file shareparameter trống
        /// </summary>
        /// <param name="pathFileShareParameter">Đường dẫn khi tạo file share parameter</param>
        /// <returns></returns>
        public static bool CreateDefaultShareParameter(string pathFileShareParameter)
        {
            if (string.IsNullOrEmpty(pathFileShareParameter) == false)
            {
                FileInfo fileInfo = new System.IO.FileInfo(pathFileShareParameter);
                if (fileInfo.Extension == ".txt")
                {
                    string dataDefault = "# This is a Revit shared parameter file.\r\n# Do not edit manually.\r\n*META\tVERSION\tMINVERSION\r\nMETA\t2\t1\r\n*GROUP\tID\tNAME\r\n*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\r\n";
                    File.WriteAllText(pathFileShareParameter, dataDefault);
                    return true;
                }
            }
            return false;
        }

        // --------------------------------------  GROUP ------------------------------------------------------------
        /// <summary>
        /// Lấy tất cả các groups có trong file ShareParameter
        /// </summary>
        /// <returns></returns>
        public List<DefinitionGroup> GetAllGroups()
        {
            List<DefinitionGroup> res = new List<DefinitionGroup>();
            foreach (var group in ShareParameterFile.Groups)
            {
                res.Add(group);
            }
            return res;
        }

        /// <summary>
        /// Lấy ra group trong file ShareParameter dựa vào tên của group.
        /// </summary>
        /// <param name="groupName">Tên của group cần lầy ra</param>
        /// <returns></returns>
        public DefinitionGroup GetGroup(string groupName)
        {
            DefinitionGroup group = ShareParameterFile.Groups.get_Item(groupName);
            if (group == null) group = ShareParameterFile.Groups.Create(groupName);
            return group;
        }

        /// <summary>
        /// Xoá group trong file shareparameter dựa theo tên truyền vào.
        /// </summary>
        /// <param name="groupName">Tên của group</param>
        /// <returns></returns>
        public bool DeleteGroup(string groupName)
        {
            return DeleteGroup(GetGroup(groupName));
        }

        /// <summary>
        /// Xoá group trong file shareparameter dựa theo group truyền vào.
        /// </summary>
        /// <param name="group">Group trong ShareParameter</param>
        /// <returns></returns>
        public bool DeleteGroup(DefinitionGroup group)
        {
            if (group == null) return true;
            List<Definition> lstDef = GetAllPrameters(group);
            foreach (var item in lstDef)
            {
                DeleteParameter(group, item);
            }
            string path = m_app.SharedParametersFilename;
            string[] kq = File.ReadAllLines(path);
            List<string> OutfileText = new List<string>();
            foreach (var item in kq)
            {
                if (item.Contains(group.Name) == true)
                {
                    continue;
                }
                OutfileText.Add(item);
            }
            File.WriteAllLines(path, OutfileText);
            Refresh();
            return true;
        }

        // --------------------------------------  PARAMETER ------------------------------------------------------------
        /// <summary>
        /// Lấy ra tất cả của các parameters có trong file shareparameter.
        /// </summary>
        /// <returns></returns>
        public List<Tuple<DefinitionGroup, List<Definition>>> GetAllPrameters()
        {
            List<Tuple<DefinitionGroup, List<Definition>>> result = new List<Tuple<DefinitionGroup, List<Definition>>>();
            foreach (var group in GetAllGroups())
            {
                result.Add(new Tuple<DefinitionGroup, List<Definition>>(group, group.Definitions.ToList()));
            }
            return result;
        }

        /// <summary>
        /// Lấy ra tất cả các Parameter theo tên của group.
        /// </summary>
        /// <param name="groupName">Tên group trong file ShareParameter</param>
        /// <returns></returns>
        public List<Definition> GetAllPrameters(string groupName)
        {
            var group = GetGroup(groupName);
            return (group == null) ? null : group.Definitions.ToList();
        }

        /// <summary>
        /// Lấy ra tất cả các Parameter theo tên của group.
        /// </summary>
        /// <param name="group"></param>
        /// <returns></returns>
        public List<Definition> GetAllPrameters(DefinitionGroup group)
        {
            return (group == null) ? null : group.Definitions.ToList();
        }

        /// <summary>
        /// Lấy ra parameter trong file share parameter dựa theo Group và tên của parameter.
        /// </summary>
        /// <param name="Group"></param>
        /// <param name="PramName"></param>
        /// <returns></returns>
        public Definition GetPrameter(DefinitionGroup group, string pramName)
        {
            Definition def = group.Definitions.get_Item(pramName);
            return def == null ? null : def;
        }

        /// <summary>
        /// Lấy ra parameter trong file share parameter dựa theo Group và tên của parameter.
        /// </summary>
        /// <param name="groupName">Tên của group</param>
        /// <param name="pramName">Tên của parameter</param>
        /// <returns></returns>
        public Definition GetPrameter(string groupName, string pramName)
        {
            var group = GetGroup(groupName);
            if (group == null) return null;
            Definition def = group.Definitions.get_Item(pramName);
            return def == null ? null : def;
        }

        /// <summary>
        /// Thêm mới 1 parameter.
        /// </summary>
        /// <param name="Group"></param>
        /// <param name="PramName"></param>
        /// <param name="type"></param>
        /// <param name="canEdit">Nếu là true, tức là người dùng có quyền để chỉnh sửa, ngược lại là false thì người dùng không có quyền chỉnh sửa.</param>
        /// <param name="visibleInProject">Có hiển thị thông tin parameter này trong (Project doc) khi xem type không?</param>
        /// <returns></returns>
        public Definition AddParameter(DefinitionGroup group, string pramName, ParameterTypeMap type, bool canEdit = true, bool visibleInProject = true)
        {
            Definition defItem = group.Definitions.get_Item(pramName);
            if (defItem == null)
            {
                ExternalDefinitionCreationOptions externalDefOpts = new ExternalDefinitionCreationOptions(pramName, ConvertParameterType.Convert(type));
                externalDefOpts.UserModifiable = canEdit;
                externalDefOpts.Visible = visibleInProject;
                return group.Definitions.Create(externalDefOpts);
            }
            return defItem;
        }

        /// <summary>
        /// Thêm mới một parameter.
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="pramName"></param>
        /// <param name="type"></param>
        /// <param name="canEdit"></param>
        /// <param name="visibleInProject"></param>
        /// <returns></returns>
        public Definition AddParamemeter(string groupName, string pramName, ParameterTypeMap type, bool canEdit = true, bool visibleInProject = true)
        {
            var group = GetGroup(groupName);
            if (group == null) return null;
            return AddParameter(group, pramName, type, canEdit, visibleInProject);
        }

#if REVIT2020 || REVIT2021
        /// <summary>
        /// Thêm mới một parameter.
        /// </summary>
        /// <param name="group"></param>
        /// <param name="pramName"></param>
        /// <param name="type"></param>
        /// <param name="canEdit"></param>
        /// <param name="visibleInProject"></param>
        /// <returns></returns>
        public Definition AddParameter(DefinitionGroup group, string pramName, ParameterType type, bool canEdit = true, bool visibleInProject = true)
        {
            Definition defItem = group.Definitions.get_Item(pramName);
            if (defItem == null)
            {
                ExternalDefinitionCreationOptions externalDefOpts = new ExternalDefinitionCreationOptions(pramName, type);
                externalDefOpts.UserModifiable = canEdit;
                externalDefOpts.Visible = visibleInProject;
                return group.Definitions.Create(externalDefOpts);
            }
            return defItem;
        }
        /// <summary>
        /// Thêm mới một parameter.
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="pramName"></param>
        /// <param name="type"></param>
        /// <param name="canEdit"></param>
        /// <param name="visibleInProject"></param>
        /// <returns></returns>
        public Definition AddParamemeter(string groupName, string pramName, ParameterType type, bool canEdit = true, bool visibleInProject = true)
        {
            var group = GetGroup(groupName);
            if (group == null) return null;
            return AddParameter(group, pramName, type, canEdit, visibleInProject);
        }
#else

        /// <summary>
        /// Thêm mới một parameter.
        /// </summary>
        /// <param name="group"></param>
        /// <param name="pramName"></param>
        /// <param name="unitTypeId"></param>
        /// <param name="canEdit"></param>
        /// <param name="visibleInProject"></param>
        /// <returns></returns>
        public Definition AddParameter(DefinitionGroup group, string pramName, ForgeTypeId unitTypeId, bool canEdit = true, bool visibleInProject = true)
        {
            Definition defItem = group.Definitions.get_Item(pramName);
            if (defItem == null)
            {
                ExternalDefinitionCreationOptions externalDefOpts = new ExternalDefinitionCreationOptions(pramName, unitTypeId);
                externalDefOpts.UserModifiable = canEdit;
                externalDefOpts.Visible = visibleInProject;
                return group.Definitions.Create(externalDefOpts);
            }
            return defItem;
        }

        /// <summary>
        /// Thêm mới một parameter.
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="pramName"></param>
        /// <param name="unitTypeId"></param>
        /// <param name="canEdit"></param>
        /// <param name="visibleInProject"></param>
        /// <returns></returns>
        public Definition AddParamemeter(string groupName, string pramName, ForgeTypeId unitTypeId, bool canEdit = true, bool visibleInProject = true)
        {
            var group = GetGroup(groupName);
            if (group == null) return null;
            return AddParameter(group, pramName, unitTypeId, canEdit, visibleInProject);
        }

#endif

        /// <summary>
        /// Xoá một parameter.
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="parameterName"></param>
        /// <returns></returns>
        public bool DeleteParameter(string groupName, string parameterName)
        {
            var group = GetGroup(groupName);
            if (group == null) return true;
            var param = GetPrameter(group, parameterName);
            return DeleteParameter(group, param);
        }

        /// <summary>
        /// Xoá một parameter.
        /// </summary>
        /// <param name="group"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public bool DeleteParameter(DefinitionGroup group, Definition param)
        {
            if (group == null || param == null) return true;
            Definition def = group.Definitions.get_Item(param.Name);
            string path = m_app.SharedParametersFilename;
            List<string> OutfileText = new List<string>();
            if (def != null)
            {
                string[] kq = File.ReadAllLines(path);
                string IdGroup = "";
                foreach (var item in kq)
                {
                    if (item.Contains(group.Name) == true)
                    {
                        IdGroup = item.Split('\t')[1];
                    }
                    if (item.Contains(param.Name) && item.Split('\t')[5] == IdGroup)
                    {
                        continue;
                    }
                    OutfileText.Add(item);
                }
                File.WriteAllLines(path, OutfileText, Encoding.Unicode);
                Refresh();
                return true;
            }
            return false;
        }

        // --------------------------------------  OTHER ------------------------------------------------------------
        private void Refresh()
        {
            ShareParameterFile = null;
            m_app.SharedParametersFilename = m_app.SharedParametersFilename;
            ShareParameterFile = m_app.OpenSharedParameterFile();
        }

        /// <summary>
        /// Kiểm tra một parameter đã tồn tại hay chưa.
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="pramName"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public bool CanCreateParamter(string groupName, string pramName, ParameterTypeMap type, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            if (m_existFile == false) return true;
            var Group = GetGroup(groupName);
            if (Group == null) return true;
            Definition defItem = Group.Definitions.get_Item(pramName);
            if (defItem == null) return true;
            if (defItem.GetParameterType() != ConvertParameterType.Convert(type))
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            status = ParameterWaringStatus.Equal;
            return false;
        }

#if REVIT2020 || REVIT2021
        /// <summary>
        /// Kiểm tra một parameter đã tồn tại hay chưa.
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="pramName"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public bool CanCreateParamter(string groupName, string pramName, ParameterType type, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            if (m_existFile == false) return true;
            var Group = GetGroup(groupName);
            if (Group == null) return true;
            Definition defItem = Group.Definitions.get_Item(pramName);
            if (defItem == null) return true;
            if (defItem.GetParameterType() != type)
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            status = ParameterWaringStatus.Equal;
            return false;
        }
#else

        /// <summary>
        /// Kiểm tra xem có giống dữ liệu đưa vào hay không.
        /// </summary>
        /// <param name="group"></param>
        /// <param name="pramName"></param>
        /// <param name="unitTypeId"></param>
        /// <returns></returns>
        public bool CanCreateParamter(string groupName, string pramName, ForgeTypeId unitTypeId, out ParameterWaringStatus status)
        {
            status = ParameterWaringStatus.None;
            var group = GetGroup(groupName);
            if (group == null) return true;
            Definition defItem = group.Definitions.get_Item(pramName);
            if (defItem == null) return true;
            if (defItem.GetParameterType() != unitTypeId)
            {
                status = ParameterWaringStatus.Different_ParamType;
                return false;
            }
            status = ParameterWaringStatus.Equal;
            return false;
        }

#endif

        /// <summary>
        /// Set lại về đường dẫn mặc định trước khi sử dụng class này để quản lý./
        /// </summary>
        public void RoolBackPath(bool isForceClean)
        {
            if (string.IsNullOrEmpty(PathFileOld) == false && File.Exists(PathFileOld) == true)
            {
                m_app.SharedParametersFilename = PathFileOld;
                if (isForceClean)
                {
                    FileInfo fileInfo = new FileInfo(PathFileNew);
                    if (fileInfo.Exists) fileInfo.Delete();
                }
            }
        }
    }
}