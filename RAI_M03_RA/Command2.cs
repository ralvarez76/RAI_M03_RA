#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. Get rooms
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            collector.OfCategory(BuiltInCategory.OST_Rooms);
     

            // 3. Set options and get room boundaries
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish;
            
            int dim_counter = 0;

            using (Transaction t = new Transaction(doc))
            {

                t.Start("Create room dimensions");

                // 3a. Loop through rooms
                foreach (Room curRoom in collector)
                {
                    // 2. Create reference array and point list
                    ReferenceArray refArrayVert = new ReferenceArray();
                    ReferenceArray refArrayHorz = new ReferenceArray();
                    List<XYZ> pointListVert = new List<XYZ>();
                    List<XYZ> pointListHorz = new List<XYZ>();

                    List<BoundarySegment> boundSegList = curRoom.GetBoundarySegments(options).First().ToList();


                    // 4. Loop through room boundaries
                    foreach (BoundarySegment curSeg in boundSegList)
                    {

                        // 4a. Get boundary geometry
                        Curve boundCurve = curSeg.GetCurve();
                        XYZ dimPoint = boundCurve.Evaluate(0.2, true);

                        // 4b. Check if line is vertical
                        if (IsLineVertical(boundCurve) == true)
                        {
                            // 4ba. Get boundary wall
                            Element curWallVert = doc.GetElement(curSeg.ElementId);

                            // 4bb. Add to ref and point array
                            if (curWallVert != null)
                            {
                                refArrayVert.Append(new Reference(curWallVert));
                                pointListVert.Add(dimPoint);
                            }
                        }
                        else
                        {
                            // 4bc. Get boundary wall
                            Element curWallHorz = doc.GetElement(curSeg.ElementId);

                            // 4bd. Add to ref and point array
                            if (curWallHorz != null)
                            {
                                refArrayHorz.Append(new Reference(curWallHorz));
                                pointListHorz.Add(dimPoint);
                            }
                        }
                    }

                    // 5a. Order vertical grid point list left to right
                    List<XYZ> sortedListVert = pointListVert.OrderBy(p => p.X).ThenBy(p => p.Y).Distinct().ToList();
                    XYZ point1Vert = sortedListVert.First();
                    XYZ point2Vert = sortedListVert.Last();

                    // 5b. Order horizontal grid point list left to right
                    List<XYZ> sortedListHorz = pointListHorz.OrderBy(p => p.Y).ThenBy(p => p.X).Distinct().ToList();
                    XYZ point1Horz = sortedListHorz.First();
                    XYZ point2Horz = sortedListHorz.Last();

                    // 6a. Create line for vertical dimension
                    Line dimLineVert = Line.CreateBound(point1Vert, new XYZ(point2Vert.X, point1Vert.Y, 0));

                    // 6b. Create line for horizontal dimension
                    Line dimLineHorz = Line.CreateBound(point1Horz, new XYZ(point1Horz.X, point2Horz.Y, 0));
                                
                    // 7. Create dimensions
                    Dimension newDimVert = doc.Create.NewDimension(doc.ActiveView, dimLineVert, refArrayVert);
                    dim_counter++;

                    Dimension newDimHorz = doc.Create.NewDimension(doc.ActiveView, dimLineHorz, refArrayHorz);
                    dim_counter++;
                                    
                    }
                t.Commit();
            }

            TaskDialog.Show("Complete", $"Inserted {dim_counter} dimensions");

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
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}
