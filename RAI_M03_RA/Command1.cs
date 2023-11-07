#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace RAI_M03_RA
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. Get grid lines
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            collector.OfClass(typeof(Grid));

            // 2. Create reference array and point list
            ReferenceArray refArrayGridVert = new ReferenceArray();
            ReferenceArray refArrayGridHorz = new ReferenceArray();
            List<XYZ> pointListGridVert = new List<XYZ>();
            List<XYZ> pointListGridHorz = new List<XYZ>();

            // 3a. Loop through vertical lines
            foreach (Grid curLine in collector)
            {
                // 3a. Get start point of grid line
                Curve curve = curLine.Curve;
                XYZ startPoint = curve.Evaluate(1, true);

                // 7. Check if line is vertical
                if (IsLineVertical(curve) == true)
                    continue;

                // 3b. Add lines to ref array
                refArrayGridVert.Append(new Reference(curLine));

                // 3c. Add midpoint to list
                pointListGridVert.Add(startPoint);

            }

            // 3b. Loop through horizontal lines
            foreach (Grid curLine in collector)
            {
                // 3a. Get start point of grid line
                Curve curve = curLine.Curve;
                XYZ startPoint = curve.Evaluate(1, true);

                // 7. Check if line is vertical
                if (IsLineVertical(curve) == false)
                    continue;

                // 3b. Add lines to ref array
                refArrayGridHorz.Append(new Reference(curLine));

                // 3c. Add midpoint to list
                pointListGridHorz.Add(startPoint);

            }

            // 4a. Order vertical grid point list left to right
            List<XYZ> sortedListVert = pointListGridVert.OrderBy(p => p.X).ThenBy(p => p.Y).Distinct().ToList();
            XYZ point1Vert = sortedListVert.First();
            XYZ point2Vert = sortedListVert.Last();

            // 4b. Order horizontal grid point list left to right
            List<XYZ> sortedListHorz = pointListGridHorz.OrderBy(p => p.Y).ThenBy(p => p.X).Distinct().ToList();
            XYZ point1Horz = sortedListHorz.First();
            XYZ point2Horz = sortedListHorz.Last();

            // 5a. Create line for vertical grid dimension
            Line dimLineVert = Line.CreateBound(point1Vert, new XYZ(point1Vert.X, point2Vert.Y, 0));

            // 5b. Create line for horizontal grid dimension
            Line dimLineHorz = Line.CreateBound(point1Horz, new XYZ(point2Horz.X, point1Horz.Y, 0));

            // 6. Create grid dimensions
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create grid dimensions");
                Dimension newDimVert = doc.Create.NewDimension(doc.ActiveView, dimLineVert, refArrayGridVert);
                Dimension newDimHorz = doc.Create.NewDimension(doc.ActiveView, dimLineHorz, refArrayGridHorz);
                XYZ moveVectorX = new XYZ(3, 0, 0);
                XYZ moveVectorY = new XYZ(0, -3, 0);
                ElementTransformUtils.MoveElement(doc, newDimVert.Id, moveVectorX);
                ElementTransformUtils.MoveElement(doc, newDimHorz.Id, moveVectorY);
                t.Commit();
            }

            return Result.Succeeded;
        }

        private bool IsLineVertical(Curve curLine)
        {
            XYZ p1 = curLine.GetEndPoint(0);
            XYZ p2 = curLine.GetEndPoint(1);

            if (Math.Abs(p1.X - p2.X) < Math.Abs(p1.Y - p2.Y))
                return true;
            return false;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
