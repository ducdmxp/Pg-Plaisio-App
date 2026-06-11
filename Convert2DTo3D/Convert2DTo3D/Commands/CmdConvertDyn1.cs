using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Media3D;

namespace Convert2DTo3D.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdConvertDyn1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Transaction tran = new Transaction(doc, "Test");

            try
            {
                int option = 0;

                tran.Start();

                RevitLinkInstance linkInstance = doc.GetElement(new ElementId((long)39656553)) as RevitLinkInstance;
                Level baseLevel = doc.GetElement(new ElementId((long)37374249)) as Level;
                Level aboveLevel = doc.GetElement(new ElementId((long)37374253)) as Level;

                FilterFireProtectionElements(linkInstance, baseLevel, aboveLevel,
                                            out List<Element> Felms,
                                            out List<Element> tgFElm,
                                            out List<Element> tgFlrs);

                List<List<Element>> lstListEles = CombineAndJoinLists<Element>(tgFElm, tgFlrs);

                //List<Element>> elements1 = ScopeIfLogic(false, Felms, lstListEles);

                List<Element> allDucts = GetAllElementByCategorys(doc);

                List<Element> smokeDucts = FilterSmokeDucts(doc, allDucts);

                List<Element> lstDuctValids = (option <= 1) ? allDucts : smokeDucts;

                tran.Commit();
            }
            catch (Exception)
            {
                tran.RollBack();
            }

            return Result.Succeeded;
        }

        public void FilterFireProtectionElements(RevitLinkInstance linkInstance,
                                                 Level baseLevel, Level aboveLevel,
                                                 out List<Element> Felms,
                                                 out List<Element> tgFElm,
                                                 out List<Element> tgFlrs)
        {
            Felms = new List<Element>();
            tgFElm = new List<Element>();
            tgFlrs = new List<Element>();

            // 1. Lấy Document của file Link
            Document linkDoc = linkInstance.GetLinkDocument();
            if (linkDoc == null) return;

            // 2. Danh sách tên các tham số cần kiểm tra (đã sửa lỗi dấu cách)
            string[] prmNames = {
                 "法_防火区画（面積）",
                 "法_防火区画（高層）",
                 "法_防火区画（層間）",
                 "法_防火区画（竪穴）",
                 "法_防火区画（異種）"
            };

            // 3. Lấy tất cả Tường (Walls) và Sàn (Floors) từ file Link
            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(
                new List<BuiltInCategory> { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors }
            );

            var linkedElements = new FilteredElementCollector(linkDoc)
                 .WherePasses(catFilter)
                 .WhereElementIsNotElementType()
                 .ToList();

            // 4. Lọc các phần tử có ít nhất một trong các tham số trên bằng 1 (True)
            // Tương đương dòng 10-13 trong Dynamo của bạn
            Felms = linkedElements.Where(e =>
            {
                foreach (string name in prmNames)
                {
                    Parameter p = e.LookupParameter(name);
                    // Kiểm tra nếu tham số tồn tại và có giá trị là 1 (Integer/YesNo)
                    if (p != null && p.HasValue && p.AsInteger() == 1)
                        return true;
                }
                return false;
            }).ToList();

            // 5. Lọc theo Level (Tương đương dòng 16-19)
            // Lưu ý: Trong Revit API, so sánh Level nên dùng Id để chính xác tuyệt đối
            tgFElm = Felms.Where(e =>
            {
                Parameter pLevel = e.LookupParameter("基準レベル");
                if (pLevel != null && linkDoc.GetElement(pLevel.AsElementId()) is Level level)
                    return level.Name == baseLevel.Name;

                return false;
            }).ToList();

            // 6. Lọc Sàn ở tầng trên (Tương đương dòng 22-26)
            tgFlrs = Felms
                 .Where(e => (int)e.Category.Id.Value == (int)BuiltInCategory.OST_Floors)
                 .Where(e =>
                 {
                     Parameter pLevel = e.LookupParameter("基準レベル");
                     if (pLevel != null && linkDoc.GetElement(pLevel.AsElementId()) is Level level)
                         return level.Name == aboveLevel.Name;

                     return false;
                 }).ToList();

            // Tiếp theo bạn có thể sử dụng danh sách tgFElm và tgFlrs cho các bước tiếp theo...
        }

        /// <summary>
        /// Mô phỏng List.Combine với List.Join: Gộp 2 danh sách thành các danh sách con theo cặp.
        /// </summary>
        /// <typeparam name="T">Kiểu dữ liệu (Element, string, v.v.)</typeparam>
        public List<List<T>> CombineAndJoinLists<T>(List<T> list1, List<T> list2)
        {
            List<List<T>> combinedResult = new List<List<T>>();

            // Lấy số lượng phần tử nhỏ nhất của 2 list để tránh lỗi index
            int count = Math.Min(list1.Count, list2.Count);

            for (int i = 0; i < count; i++)
            {
                // Với mỗi vị trí i, tạo một list mới chứa cả 2 phần tử (Join)
                List<T> subList = new List<T> { list1[i], list2[i] };

                // Thêm vào list tổng (Combine)
                combinedResult.Add(subList);
            }

            return combinedResult;
        }

        public List<List<Element>> ScopeIfLogic(bool test, List<List<Element>> trueResult, List<List<Element>> falseResult)
        {
            // Khai báo biến chứa kết quả cuối cùng
            List<List<Element>> finalResult;

            if (test)
            {
                // Nếu test là True, lấy dữ liệu từ nhánh True
                finalResult = trueResult;
            }
            else
            {
                // Nếu test là False, lấy dữ liệu từ nhánh False
                finalResult = falseResult;
            }

            return finalResult;
        }

        public List<Element> GetAllElementByCategorys(Document doc, BuiltInCategory builtIn = BuiltInCategory.OST_DuctCurves)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            List<Element> elements = collector
                .OfCategory(builtIn)
                .WhereElementIsNotElementType()
                .ToElements().ToList();

            return elements;
        }

        public List<Element> FilterSmokeDucts(Document doc, List<Element> allDucts)
        {
            // Danh sách để chứa các ống gió thỏa mãn điều kiện "Hút khói"
            List<Element> smokeDucts = new List<Element>();

            foreach (Element duct in allDucts)
            {
                // 1. Lấy tham số "System Type" (trong API thường dùng BuiltInParameter)
                Parameter systemTypeParam = duct.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);

                if (systemTypeParam != null && systemTypeParam.HasValue)
                {
                    // 2. Lấy ElementId của System Type và truy xuất đối tượng đó
                    ElementId systemTypeId = systemTypeParam.AsElementId();
                    Element systemTypeElement = doc.GetElement(systemTypeId);

                    if (systemTypeElement != null)
                    {
                        // 3. Lấy tên của System Type (tương đương SysTypeName trong Dynamo)
                        string sysTypeName = systemTypeElement.Name;

                        // 4. Kiểm tra xem tên có chứa "排煙" (Hút khói) không
                        // Tương đương String.Contains(..., "排煙", false)
                        bool isSmokeDuct = sysTypeName.Contains("排煙", StringComparison.OrdinalIgnoreCase);

                        if (isSmokeDuct)
                        {
                            smokeDucts.Add(duct);
                        }
                    }
                }
            }

            return smokeDucts;
        }

        public List<IntersectionPair> GetIntersects(Document doc, List<Element> setA, List<Element> setB)
        {
            List<IntersectionPair> results = new List<IntersectionPair>();

            // Tối ưu hóa: Lấy danh sách ID của Set B để giới hạn vùng quét.
            // Điều này giúp thuật toán không phải tìm kiếm trên toàn bộ mô hình.
            ICollection<ElementId> setB_Ids = setB.Select(e => e.Id).ToList();

            if (setB_Ids.Count == 0) return results;

            foreach (Element elemA in setA)
            {
                // Bỏ qua các phần tử không có hình khối (ví dụ: line, text...)
                if (elemA.get_BoundingBox(null) == null) continue;

                // Tạo bộ lọc: Tìm các phần tử giao cắt với elemA
                ElementIntersectsElementFilter clashFilter = new ElementIntersectsElementFilter(elemA);

                // Áp dụng bộ lọc NHƯNG chỉ quét trong danh sách ID của set B
                FilteredElementCollector collector = new FilteredElementCollector(doc, setB_Ids);
                IList<Element> intersectingElements = collector.WherePasses(clashFilter).ToElements();

                // Lưu kết quả vào danh sách
                foreach (Element elemB in intersectingElements)
                {
                    results.Add(new IntersectionPair
                    {
                        ElementA = elemA,
                        ElementB = elemB
                    });
                }
            }

            return results;
        }
    }

    public class IntersectionPair
    {
        public Element ElementA { get; set; }
        public Element ElementB { get; set; }
    }
}