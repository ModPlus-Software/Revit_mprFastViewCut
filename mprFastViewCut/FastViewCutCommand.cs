namespace mprFastViewCut
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Autodesk.Revit.Attributes;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.UI.Selection;
    using ModPlus_Revit.Utils;
    using ModPlusAPI;
    using ModPlusAPI.Windows;

    /// <inheritdoc />
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FastViewCutCommand : IExternalCommand
    {
        /// <inheritdoc />
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
#if !DEBUG
                Statistic.SendCommandStarting(new ModPlusConnector());
#endif

                var activeView = commandData.Application.ActiveUIDocument.Document.ActiveView;
                if (activeView.IsTemplate)
                {
                    // Работа в шаблоне вида невозможна
                    MessageBox.Show(Language.GetItem("h1"));
                    return Result.Cancelled;
                }

                if (activeView.ViewType == ViewType.Legend)
                {
                    // Работа в легендах невозможна
                    MessageBox.Show(Language.GetItem("h2"));
                    return Result.Cancelled;
                }

                if (activeView.ViewType == ViewType.Schedule)
                {
                    // Работа в спецификациях невозможна
                    MessageBox.Show(Language.GetItem("h3"));
                    return Result.Cancelled;
                }

                if (activeView.ViewType == ViewType.DraftingView)
                {
                    // Работа в чертежных видах невозможна
                    MessageBox.Show(Language.GetItem("h4"));
                    return Result.Cancelled;
                }

                if (activeView.ViewType == ViewType.DrawingSheet)
                {
                    // Работа на листах на данный момент недоступна. Используйте плагин на конкретных видах
                    MessageBox.Show(Language.GetItem("h7"));
                    return Result.Cancelled;
                }

                CutView(commandData.Application);

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                return Result.Failed;
            }
        }

        private void CutView(UIApplication uiApplication)
        {
            var trName = Language.GetFunctionLocalName(new ModPlusConnector());
            var uiDoc = uiApplication.ActiveUIDocument;
            var doc = uiDoc.Document;
            var selection = uiDoc.Selection;

            // Укажите прямоугольную область для создания границ подрезки
            var pickedBox = selection.PickBox(PickBoxStyle.Crossing, Language.GetItem("h5"));

            if (pickedBox.Min.DistanceTo(pickedBox.Max) < 1.0.MmToFt())
                return;

            var view = doc.ActiveView;

            if (view is View3D view3D)
            {
                // from: https://thebuildingcoder.typepad.com/blog/2009/12/crop-3d-view-to-room.html

                var bb = view3D.CropBox;
                var transform = bb.Transform;
                var transformInverse = transform.Inverse;

                var pt1 = transformInverse.OfPoint(pickedBox.Min);
                var pt2 = transformInverse.OfPoint(pickedBox.Max);

                var xMin = GetSmaller(pt1.X, pt2.X);
                var xMax = GetBigger(pt1.X, pt2.X);
                var yMin = GetSmaller(pt1.Y, pt2.Y);
                var yMax = GetBigger(pt1.Y, pt2.Y);

                bb.Max = new XYZ(xMax, yMax, bb.Max.Z);
                bb.Min = new XYZ(xMin, yMin, bb.Min.Z);

                using (var tr = new Transaction(doc, trName))
                {
                    tr.Start();
                    if (!view.CropBoxActive)
                    {
                        view.CropBoxActive = true;
                        view.CropBoxVisible = false;
                    }

                    view3D.CropBox = bb;
                    tr.Commit();
                }
            }
            else if (view is ViewSheet viewSheet)
            {
                // на листах нужно определить какой Viewport попадает в выделенную область и подрезать его
                Viewport intersectedViewport = null;
                var intersectArea = double.NaN;
                foreach (var viewport in viewSheet.GetAllViewports()
                    .Select(id => doc.GetElement(id) as Viewport)
                    .Where(e => e != null))
                {
                    if (IsIntersect(viewport, pickedBox, out var a) &&
                        (double.IsNaN(intersectArea) || a > intersectArea))
                    {
                        intersectedViewport = viewport;
                        intersectArea = a;
                    }
                }

                if (intersectedViewport != null)
                {
                    view = doc.GetElement(intersectedViewport.ViewId) as View;

                    // TODO Нужно выполнить трансформацию точек из пространства листа в пространство вида
                    CropView(view, pickedBox, doc, trName);
                }
            }
            else
            {
                CropView(view, pickedBox, doc, trName);
            }
        }

        private static void CropView(View view, PickedBox pickedBox, Document doc, string trName)
        {
            var cropRegionShapeManager = view.GetCropRegionShapeManager();

            var pt1 = pickedBox.Min;
            var pt3 = pickedBox.Max;
            var plane = CreatePlane(view.UpDirection, pt3);
            var pt2 = plane.ProjectOnto(pt1);
            plane = CreatePlane(view.UpDirection, pt1);
            var pt4 = plane.ProjectOnto(pt3);

            var line1 = TryCreateLine(pt1, pt2);
            var line2 = TryCreateLine(pt2, pt3);
            var line3 = TryCreateLine(pt3, pt4);
            var line4 = TryCreateLine(pt4, pt1);

            if (line1 == null || line2 == null || line3 == null || line4 == null)
            {
                // Не удалось получить валидную прямоугольную область. Попробуйте еще раз
                MessageBox.Show(Language.GetItem("h6"));
                return;
            }

            var curveLoop = CurveLoop.Create(new List<Curve>
            {
                line1, line2, line3, line4
            });

            if (curveLoop.IsRectangular(CreatePlane(view.ViewDirection, view.Origin)))
            {
                using (var tr = new Transaction(doc, trName))
                {
                    tr.Start();
                    if (!view.CropBoxActive)
                    {
                        view.CropBoxActive = true;
                        view.CropBoxVisible = false;
                    }

                    cropRegionShapeManager.SetCropShape(curveLoop);

                    tr.Commit();
                }
            }
            else
            {
                // Не удалось получить валидную прямоугольную область. Попробуйте еще раз
                MessageBox.Show(Language.GetItem("h6"));
            }
        }

        private static Plane CreatePlane(XYZ vectorNormal, XYZ origin)
        {
            return Plane.CreateByNormalAndOrigin(vectorNormal, origin);
        }

        private static double GetBigger(double d1, double d2)
        {
            return d1 > d2 ? d1 : d2;
        }

        private static double GetSmaller(double d1, double d2)
        {
            return d1 < d2 ? d1 : d2;
        }
        
        private static Line TryCreateLine(XYZ pt1, XYZ pt2)
        {
            try
            {
                if (pt1.DistanceTo(pt2) < 1.0.MmToFt())
                    return null;

                return Line.CreateBound(pt1, pt2);
            }
            catch
            {
                return null;
            }
        }

        private bool IsIntersect(Viewport viewport, PickedBox pickedBox, out double intersectArea)
        {
            var plane = CreatePlane(XYZ.BasisZ, XYZ.Zero);
            Outline outline = viewport.GetBoxOutline();
            var viewportRectangle = GetRectangle(outline, plane);

            Outline pickedOutline = new Outline(pickedBox.Min, pickedBox.Max);
            var pickedRectangle = GetRectangle(pickedOutline, plane);

            var intersectRectangle = System.Drawing.Rectangle.Intersect(viewportRectangle, pickedRectangle);
            if (intersectRectangle.IsEmpty)
            {
                intersectArea = double.NaN;
                return false;
            }

            intersectArea = intersectRectangle.Width * intersectRectangle.Height;
            return true;
        }

        private System.Drawing.Rectangle GetRectangle(Outline outline, Plane plane)
        {
            var min = plane.ProjectOnto(outline.MinimumPoint);
            var minPointX = (int)Math.Round(min.X.FtToMm());
            var minPointY = (int)Math.Round(min.Y.FtToMm());
            var max = plane.ProjectOnto(outline.MaximumPoint);
            var maxPointX = (int)Math.Round(max.X.FtToMm());
            var maxPointY = (int)Math.Round(max.Y.FtToMm());
            var minX = Math.Min(minPointX, maxPointX);
            var maxX = Math.Max(minPointX, maxPointX);
            var minY = Math.Min(minPointY, maxPointY);
            var maxY = Math.Max(minPointY, maxPointY);
            return new System.Drawing.Rectangle(minX, minY, maxX - minX, maxY - minY);
        }
    }
}
