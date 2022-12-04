using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpeningsCreator
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class OpeningsCreator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Игнатов.Отверстия"))
                .FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Игнатов.Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transaction0 = new Transaction(arDoc, "Активация семейства");
            transaction0.Start();

                if (!familySymbol.IsActive)
                {
                    familySymbol.Activate();
                }

            transaction0.Commit();


            using (Transaction transaction = new Transaction(arDoc, "Расстановка отверстий"))
            {
                transaction.Start();

                foreach (Duct duct in ducts)
                {
                    Line curve = (duct.Location as LocationCurve).Curve as Line;
                    XYZ point = curve.GetEndPoint(0);
                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= curve.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext reference in intersections)
                    {
                        double proximity = reference.Proximity;
                        Reference objReference = reference.GetReference();
                        Wall wall = arDoc.GetElement(objReference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ insertionPoint = point + (direction * proximity);

                        FamilyInstance opening = arDoc.Create.NewFamilyInstance(insertionPoint, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = opening.LookupParameter("Ширина");
                        Parameter heigth = opening.LookupParameter("Высота");

                        width.Set(duct.Width);
                        heigth.Set(duct.Height);
                    }
                }


                foreach (Pipe pipe in pipes)
                {
                    Line curve = (pipe.Location as LocationCurve).Curve as Line;
                    XYZ point = curve.GetEndPoint(0);
                    XYZ direction = curve.Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= curve.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext reference in intersections)
                    {
                        double proximity = reference.Proximity;
                        Reference objReference = reference.GetReference();
                        Wall wall = arDoc.GetElement(objReference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ insertionPoint = point + (direction * proximity);

                        FamilyInstance opening = arDoc.Create.NewFamilyInstance(insertionPoint, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = opening.LookupParameter("Ширина");
                        Parameter heigth = opening.LookupParameter("Высота");

                        width.Set(pipe.Diameter);
                        heigth.Set(pipe.Diameter);
                    }
                }

                transaction.Commit();
            }

            return Result.Succeeded;
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
