using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Convert2DTo3D.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Convert2DTo3D.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdTest : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            Transaction tran = new Transaction(doc, "Test");

            XYZ point = uidoc.Selection.PickPoint("Pick a point :");

            List<Room> rooms = GetAllRooms(doc);

            try
            {
                tran.Start();

                Phase phase = doc.GetElement(doc.ActiveView.get_Parameter(BuiltInParameter.VIEW_PHASE).AsElementId()) as Phase;

                Excute(doc, phase, point, rooms);

                tran.Commit();
            }
            catch (Exception ex)
            {
                tran.RollBack();

                var messsage = ex.Message;
            }

            return Result.Succeeded;
        }

        public void Excute(Document doc, Phase phase, XYZ pPicked, List<Room> rooms, List<XYZ> lstPoints = null, double distance = 0)
        {
            if (doc == null || phase == null || pPicked == null || rooms?.Count <= 0)
                return;

            Room room = GetRoomBelongTo(pPicked, rooms);

            if (room != null)
            {
                List<XYZ> pointInRooms = new List<XYZ>();

                if (lstPoints?.Count > 0)
                    pointInRooms.AddRange(lstPoints);
                pointInRooms.Add(pPicked);

                FindPointInRoom(pPicked, room, ref pointInRooms);

                DrawCurve(doc, pointInRooms, distance);

                Utils.ParameterUtils.SetValueParameterByBuiltIn(room, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, "minhduc");
            }

            List<FamilyInstance> lstDoors = GetDoorsInRoom(room);

            foreach (FamilyInstance door in lstDoors)
            {
                double distanceIten = distance;

                Room roomTo = door.get_ToRoom(phase);
                Room roomFrom = door.get_FromRoom(phase);

                Room roomValid = (room != null && roomTo != null && room.Id != roomTo?.Id) ? roomTo : roomFrom;

                if (roomValid == null) continue;

                string strValue = (string)Utils.ParameterUtils.GetValueParameterByBuilt(roomValid, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);

                if (strValue == "minhduc") continue;

                XYZ location = ((LocationPoint)door.Location).Point;

                List<XYZ> allPts = new List<XYZ>();

                XYZ point1 = location + door.FacingOrientation * 1000 / 304.8;

                if (room != null && !room.IsPointInRoom(point1)) allPts.Add(point1);

                XYZ point2 = location + door.FacingOrientation.Negate() * 1000 / 304.8;

                if (room != null && !room.IsPointInRoom(point2)) allPts.Add(point2);

                foreach (XYZ pt in allPts)
                {
                    List<XYZ> pts = new List<XYZ>() { pPicked };

                    if (room != null && room.IsPointInRoom(point1)) pts.Add(point1);

                    if (room != null && room.IsPointInRoom(point2)) pts.Add(point2);

                    XYZ pointInRoomCurrent = (room != null && room.IsPointInRoom(point1)) ? point1 : point2;
                    XYZ pointInRoomNext = (room != null && !room.IsPointInRoom(point1)) ? point1 : point2;

                    List<XYZ> pt2s = new List<XYZ>() { pPicked, pointInRoomCurrent, pointInRoomNext };

                    for (int i = 0; i < pt2s.Count; i++)
                    {
                        if (i + 1 >= pt2s.Count) continue;

                        var pt0 = pt2s[i];
                        var pt1 = pt2s[i + 1];

                        distanceIten += pt0.DistanceTo(pt1);
                    }
                    Utils.ParameterUtils.SetValueParameterByBuiltIn(roomValid, BuiltInParameter.ROOM_OCCUPANCY, (distanceIten * 304.8).ToString(".0000"));

                    pts = pts.GroupBy(x => x.ToString()).Select(x => x.FirstOrDefault()).ToList();

                    Excute(doc, phase, pt, rooms, pts, distanceIten);

                    break;
                }
            }
        }

        public List<FamilyInstance> GetDoorsInRoom(Room room)
        {
            if (room == null) return new List<FamilyInstance>();

            Document doc = room.Document;
            List<FamilyInstance> doorsInRoom = new List<FamilyInstance>();

            // 1. Thu thập tất cả các đối tượng thuộc Category Doors trong dự án
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            ICollection<Element> allDoors = collector
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilyInstance))
                .ToElements();

            // 2. Duyệt qua từng cửa để kiểm tra nó có thuộc Room này không
            foreach (Element e in allDoors)
            {
                FamilyInstance door = e as FamilyInstance;
                if (door == null) continue;

                // Kiểm tra phòng ở cả hai phía của cánh cửa (Phòng đi từ/Phòng đi đến)
                // Lưu ý: Phase của Room và Door phải khớp nhau để kết quả chính xác
                if ((door.FromRoom != null && door.FromRoom.Id == room.Id) ||
                    (door.ToRoom != null && door.ToRoom.Id == room.Id))
                {
                    doorsInRoom.Add(door);
                }
            }

            return doorsInRoom;
        }

        public void FindPointInRoom(XYZ point, Room room, ref List<XYZ> points)
        {
            if (room == null) return;

            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();

            List<Line> lines = new List<Line>();

            foreach (var lstListBoundary in room.GetBoundarySegments(options))
            {
                foreach (var item in lstListBoundary)
                {
                    Curve curve = item.GetCurve();

                    List<XYZ> pointItems = new List<XYZ>() { curve.GetEndPoint(0), curve.GetEndPoint(1) };

                    foreach (var pt in pointItems)
                    {
                        var p0 = point.To2D();
                        var p1 = pt.To2D();

                        if (Math.Abs(p0.DistanceTo(p1)) <= 100 / 304.8) continue;

                        Line lineItem = Line.CreateBound(point.To2D(), pt.To2D());

                        lines.Add(lineItem);
                    }
                }
            }

            Line lineMax = lines.OrderByDescending(x => x.ApproximateLength).FirstOrDefault();

            if (lineMax != null)
                points.Add(lineMax.GetEndPoint(1));
        }

        public void DrawCurve(Document doc, List<XYZ> points, double distance = 0)
        {
            double lenght = distance + points[points.Count - 2].DistanceTo(points[points.Count - 1]);

            for (int i = 0; i < points.Count; i++)
            {
                if (i + 1 >= points.Count) continue;

                var pt0 = points[i];
                var pt1 = points[i + 1];

                Line line = Line.CreateBound(pt0, pt1);
                Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, pt0);
                SketchPlane sketchPlane = SketchPlane.Create(doc, plane);

                var model = doc.Create.NewModelCurve(line, sketchPlane);

                //lenght += line.ApproximateLength;
            }

            lenght *= 304.8 / 1000;

            lenght = Math.Round(lenght, 4);

            CreateTextNote(doc, points.LastOrDefault(), lenght.ToString(".0000") + "m");
        }

        public Room GetRoomBelongTo(XYZ point, List<Room> rooms)
        {
            foreach (Room room in rooms)
            {
                if (room.IsPointInRoom(point))
                    return room;
            }
            return null;
        }

        public void CreateTextNote(Document doc, XYZ point, string promt)
        {
            View activeView = doc.ActiveView;

            ElementId defaultTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);

            TextNoteOptions options = new TextNoteOptions
            {
                TypeId = defaultTypeId,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                Rotation = 0 // Đơn vị là Radian
            };

            TextNote note = TextNote.Create(doc, activeView.Id, point, promt, options);
        }

        public List<Room> GetAllRooms(Document doc)
        {
            // Create a filter specifically for Room elements
            RoomFilter filter = new RoomFilter();

            // Apply the filter to the document collector
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);

            // Cast to Room to get a typed list
            List<Room> rooms = collector
                .WherePasses(filter)
                .Cast<Room>()
                .ToList();

            return rooms;
        }
    }
}