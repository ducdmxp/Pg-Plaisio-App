using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Convert2DTo3D.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using ParameterUtils = Convert2DTo3D.Utils.ParameterUtils;

namespace Convert2DTo3D.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdCreateMEPSchedule : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = uiapp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            Global.UIApp = commandData.Application;
            Global.RVTApp = commandData.Application.Application;
            Global.UIDoc = commandData.Application.ActiveUIDocument;
            Global.Doc = commandData.Application.ActiveUIDocument.Document;
            Global.AppCreation = commandData.Application.Application.Create;

            ScheduleSleeveFrm frm = new ScheduleSleeveFrm();
            if (frm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return Result.Cancelled;

            Transaction tran = new Transaction(doc, "CreateSchedules");

            List<string> args = new List<string>() { "Share Sleeve Length", "Share Sleeve Height", "Share Sleeve Width",
                                                     "Share Sleeve Diameter","Share Sleeve Thickness", "DonVi" ,"Share Temp Param" };

            try
            {
                tran.Start();

                Convert2DTo3D.Utils.MngShareParameter.CreateProjectParamter(doc);

                ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_DuctAccessory));

                ScheduleDefinition definition = schedule.Definition;

                // Lấy tất cả available fields
                IList<SchedulableField> availableFields = definition.GetSchedulableFields();

                //Family
                var fieldFamily = availableFields.FirstOrDefault(f =>
                        f.GetName(doc).Equals("Family", StringComparison.OrdinalIgnoreCase));

                ScheduleField sfFamily = definition.AddField(fieldFamily);

                //Type

                var fieldWidth = availableFields.FirstOrDefault(f =>
                     f.GetName(doc).Equals("Share Sleeve Width", StringComparison.OrdinalIgnoreCase));

                var fieldHeight = availableFields.FirstOrDefault(f =>
                    f.GetName(doc).Equals("Share Sleeve Height", StringComparison.OrdinalIgnoreCase));

                var fieldDiameter = availableFields.FirstOrDefault(f =>
                  f.GetName(doc).Equals("Share Sleeve Diameter", StringComparison.OrdinalIgnoreCase));

                TableCellCombinedParameterData cellFieldDiameter = TableCellCombinedParameterData.Create();

                cellFieldDiameter.Prefix = "D";

                cellFieldDiameter.ParamId = fieldDiameter.ParameterId;

                TableCellCombinedParameterData cellFieldWidth = TableCellCombinedParameterData.Create();

                cellFieldWidth.ParamId = fieldWidth.ParameterId;

                TableCellCombinedParameterData cellFieldHeight = TableCellCombinedParameterData.Create();

                cellFieldHeight.Prefix = " x ";
                cellFieldHeight.ParamId = fieldHeight.ParameterId;

                definition.InsertCombinedParameterField(new List<TableCellCombinedParameterData> { cellFieldDiameter, cellFieldWidth, cellFieldHeight }, "Type", 1);

                //lenght
                var fieldLength = availableFields.FirstOrDefault(f =>
                 f.GetName(doc).Equals("Share Sleeve Length", StringComparison.OrdinalIgnoreCase));

                ScheduleField sfLength = definition.AddField(fieldLength);

                sfLength.ColumnHeading = "Length";

                //thickness
                var fieldThickness = availableFields.FirstOrDefault(f =>
                f.GetName(doc).Equals("Share Sleeve Thickness", StringComparison.OrdinalIgnoreCase));

                ScheduleField sfThickness = definition.AddField(fieldThickness);
                sfThickness.ColumnHeading = "Thickness";

                //Don vi
                var fieldDonVi = availableFields.FirstOrDefault(f =>
                f.GetName(doc).Equals("DonVi", StringComparison.OrdinalIgnoreCase));

                ScheduleField sfDonvi = definition.AddField(fieldDonVi);
                sfDonvi.ColumnHeading = "Đơn Vị";

                //Khoi luong
                var fieldCount = availableFields.FirstOrDefault(f =>
                        f.GetName(doc).Equals("Count", StringComparison.OrdinalIgnoreCase));

                ScheduleField sfCount = definition.AddField(fieldCount);

                sfCount.ColumnHeading = "Khối Lượng";

                //

                TableSectionData sectionData = schedule.GetTableData().GetSectionData(SectionType.Body);

                int numberOfColumns = sectionData.NumberOfColumns;

                for (int col = 0; col < numberOfColumns; col++)
                {
                    ScheduleField scheduleField = definition.GetField(col);

                    scheduleField.HorizontalAlignment = ScheduleHorizontalAlignment.Center;
                }

                definition.IsItemized = true;

                tran.Commit();

                uiDoc.ActiveView = schedule;
            }
            catch (Exception ex)
            {
                if (tran.HasStarted())
                    tran.RollBack();
            }

            return Result.Succeeded;
        }

        public static List<ElementType> GetElementTypesByCategory(Document doc, BuiltInCategory category)
        {
            List<ElementType> elementTypes = new List<ElementType>();

            if (doc == null)
                return elementTypes;

            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .OfClass(typeof(ElementType));

                elementTypes = collector
                    .Cast<ElementType>()
                    .OrderBy(et => et.Name)
                    .ToList();

                foreach (ElementType type in elementTypes)
                {
                    ParameterUtils.SetValueParameterByName(type, "DonVi", "cái");
                    ParameterUtils.SetValueParameterByName(type, "LoaiMEP", "");
                }

                return elementTypes;
            }
            catch (Exception ex)
            {
            }

            return elementTypes;
        }
    }
}