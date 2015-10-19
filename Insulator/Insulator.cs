/*
    The Insulator creates a proper, dynamic Insulation in Autodesk's (R) Revit (R)
    Copyright (C) 2014  Maximilian Thumfart

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

*/

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;
using System.Net.Sockets;
using System.Net;
using System.Xml.Serialization;
using System.Runtime.Serialization;
using System.IO;
using System.Threading;
using System.Windows.Media.Imaging;

namespace Insulator
{
    /// <summary>
    /// Create Insulator Button
    /// </summary>
    public class CreateButton : IExternalApplication
    {
        /// <summary>
        /// Path to Assembly
        /// </summary>
        static string path = typeof(CreateButton).Assembly.Location;

        /// <summary>
        /// Create Button on Startup
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
        public Result OnStartup(UIControlledApplication a)
        {
            // Create Ribbon Panel
            RibbonPanel ribbonPanel = a.CreateRibbonPanel("Insulator");

            // Create a new Button
            PushButton pushButton = ribbonPanel.AddItem(new PushButtonData("Insulator", "Insulator", path, "Insulator.DrawInsulation")) as PushButton;
            pushButton.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://grevit.net/?p=2509"));

            // Apply the Icon
            pushButton.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                Properties.Resources.Insulator.GetHbitmap(),
                IntPtr.Zero,
                System.Windows.Int32Rect.Empty,
                BitmapSizeOptions.FromWidthAndHeight(32, 32));

            return Result.Succeeded;
        }


        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }


    /// <summary>
    /// The actual draw Insulation Command
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DrawInsulation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get environment variables
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            UIDocument uidoc = uiApp.ActiveUIDocument;

            // Create and open the material selection dialog
            InsulatorMaterial insmat = new InsulatorMaterial();
            insmat.ShowDialog();

            try
            {
                Transaction trans = new Transaction(doc, "DrawInsulation");
                trans.Start();

                foreach (Element em in SelectObjects(uidoc))
                {
                    // Get DetailLines or DetailArcs
                    if (em.GetType() == typeof(DetailLine) || em.GetType() == typeof(DetailArc))
                    {
                        // Define the Curve as the lower curve
                        DetailCurve lower = (DetailCurve)em;
                        Curve lowercurve = (Curve)lower.GeometryCurve;

                        // Prompt for the upper curve
                        Reference refe = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select upper");
                        DetailCurve upper = (DetailCurve)doc.GetElement(refe.ElementId);
                        Curve uppercurve = (Curve)upper.GeometryCurve;

                        // Collector for ElementIds for grouping
                        List<ElementId> Ids = new List<ElementId>();

                        // Initial location along the curve
                        double location = 0;

                        // Start with a left loop
                        bool left = true;

                        // Walk along the curve and draw loops
                        while ((location / lowercurve.Length) < 1)
                        {
                            location += drawInsulation(Ids, doc, lowercurve, location, uppercurve, ref left, new List<Interruption>(), insmat.zigzag);
                        }

                        // Create a group
                        doc.Create.NewGroup(Ids);

                    }

                    // Handle Walls
                    if (em.GetType() == typeof(Wall))
                    {
                        // Get the LocationCurve and cast the Wall
                        Wall wall = (Wall)em;
                        LocationCurve c = (LocationCurve)wall.Location;

                        // List for Interruptions like Doors and Windows
                        List<Interruption> Breaks = new List<Interruption>();

                        // Get the Compound Structure
                        CompoundStructure cs = wall.WallType.GetCompoundStructure();
                        
                        // Id Collection for Grouping
                        List<ElementId> Ids = new List<ElementId>();

                        // Get the walls Startpoint
                        XYZ p1 = c.Curve.GetEndPoint(0);
                        double x2 = c.Curve.GetEndPoint(1).X;

                        // Retrieve Boundaries of the Insulation layer
                        Tuple<Curve, Curve> curves = getInsulationLayer(doc, wall);

                        // Collect Openings from the wall
                        removeOpenings(ref Breaks, wall, doc);

                        // initial Insulation location
                        double location = 0;

                        // Start with a left loop
                        bool left = true;

                        // Set this wall as banned from intersection checks
                        List<int> BannedWalls = new List<int>() { wall.Id.IntegerValue };

                        // Get extensions to left hand joined walls
                        Curve lowerExtendedLeft = handleJoinsLeft(doc, c, curves.Item2, ref BannedWalls);

                        // Get extensions to other layers
                        // adjust the interruptions if the insulation layer is extended
                        double offset = lowerExtendedLeft.Length - curves.Item2.Length;
                        if (offset > 0) foreach (Interruption ir in Breaks) ir.Extend(offset);

                        // Get Extensions to right hand joined walls
                        Tuple<Curve, Curve> extended = handleJoinsRight(doc, c, lowerExtendedLeft, curves.Item1, ref BannedWalls);

                        // draw insulation loops anlong the curve
                        while (location / extended.Item2.Length < 1)
                            location += drawInsulation(Ids, doc, extended.Item2, location, extended.Item1, ref left, Breaks, insmat.zigzag);

                        // Group loops
                        doc.Create.NewGroup(Ids);
                    }

                }

                
                trans.Commit();
                trans.Dispose();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return Result.Failed;
            }
            return Result.Succeeded;
        }





        /// <summary>
        /// Select Elements form Model
        /// </summary>
        /// <param name="uidoc">UiDoc</param>
        /// <returns>Selection</returns>
        public static List<Element> SelectObjects(Autodesk.Revit.UI.UIDocument uidoc)
        {
            List<Element> selectedList = new List<Element>();

            ICollection<ElementId> selected = uidoc.Selection.GetElementIds();

            if (selected.Count > 0) foreach (ElementId elementid in selected) selectedList.Add(uidoc.Document.GetElement(elementid));
            else
            {
                IList<Reference> select = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select objects");
                foreach (Reference refer in select) selectedList.Add(uidoc.Document.GetElement(refer.ElementId));
            }

            return selectedList;
        }

        /// <summary>
        /// Remove Openings from wall baseline
        /// </summary>
        /// <param name="breaks">List of Interruptions</param>
        /// <param name="wall">Wall</param>
        /// <param name="doc">Document</param>
        public void removeOpenings(ref List<Interruption> breaks, Wall wall, Document doc)
        {
            LocationCurve c = (LocationCurve)wall.Location;
            foreach (ElementId elemId in wall.FindInserts(true, false, false, false))
            {
                Element insertedElement = doc.GetElement(elemId);
                if (insertedElement.GetType() == typeof(FamilyInstance))
                {
                    FamilyInstance o = (FamilyInstance)insertedElement;
                    if (o.Location.GetType() == typeof(LocationPoint))
                    {
                        LocationPoint lp = (LocationPoint)o.Location;
                        c.Curve.GetEndPoint(0).DistanceTo(lp.Point);
                        double width = o.Symbol.LookupParameter("Width").AsDouble();
                        breaks.Add(new Interruption(c.Curve, lp.Point, width));
                    }

                }
            }
        }

        /// <summary>
        /// Get the Insulation Layer of Left hand joined Walls
        /// </summary>
        public Curve handleJoinsLeft(Document doc, LocationCurve wallLocCurve, Curve lower, ref List<int> bannedWalls)
        {

            if (lower.GetType() == typeof(Arc) || lower.GetType() == typeof(NurbSpline)) return lower;

            XYZ IntersectionPoint = null;
            ElementArray elems = wallLocCurve.get_ElementsAtJoin(0);

            if (elems == null || elems.Size == 0) return lower;

            foreach (Element elem in elems)
            {
                if (elem.GetType() == typeof(Wall))
                {
                    Wall wnext = (Wall)elem;

                    if (!bannedWalls.Contains(wnext.Id.IntegerValue))
                    {

                        Tuple<Curve, Curve> curves = getInsulationLayer(doc, wnext);
                        Curve lowerunbound = lower.Clone();
                        lowerunbound.MakeUnbound();
                        Curve upper = curves.Item1.Clone();
                        upper.MakeUnbound();

                        IntersectionResultArray result = new IntersectionResultArray();

                        lowerunbound.Intersect(upper, out result);
                        if (result != null && result.Size == 1)
                        {
                            IntersectionPoint = result.get_Item(0).XYZPoint;
                        }

                    }

                }
            }

            if (IntersectionPoint == null) return lower;

            Line l = Line.CreateBound(IntersectionPoint, lower.GetEndPoint(1));

            return l;

        }


        /// <summary>
        /// Get the Insulation Layer of Right hand joined Walls
        /// </summary>
        public Tuple<Curve, Curve> handleJoinsRight(Document doc, LocationCurve wallLocCurve, Curve lower, Curve upper, ref List<int> bannedWalls)
        {
            Tuple<Curve, Curve> existing = new Tuple<Curve, Curve>(upper, lower);

            if (lower.GetType() == typeof(Arc) || lower.GetType() == typeof(NurbSpline)) return existing;
            if (upper.GetType() == typeof(Arc) || upper.GetType() == typeof(NurbSpline)) return existing;

            XYZ IntersectionPointLower = null;
            XYZ IntersectionPointUpper = null;

            ElementArray elems = wallLocCurve.get_ElementsAtJoin(1);

            if (elems == null || elems.Size == 0) return existing;

            foreach (Element elem in elems)
            {
                if (elem.GetType() == typeof(Wall))
                {
                    Wall wnext = (Wall)elem;

                    if (!bannedWalls.Contains(wnext.Id.IntegerValue))
                    {

                        Tuple<Curve, Curve> curves = getInsulationLayer(doc, wnext);
                        Curve lowerunbound = lower.Clone();
                        lowerunbound.MakeUnbound();
                        Curve upperunbound = upper.Clone();
                        upperunbound.MakeUnbound();
                        Curve lowernext = curves.Item2.Clone();
                        lowernext.MakeUnbound();

                        IntersectionResultArray result = new IntersectionResultArray();

                        lowerunbound.Intersect(lowernext, out result);
                        if (result != null && result.Size == 1)
                        {
                            IntersectionPointLower = result.get_Item(0).XYZPoint;
                        }

                        upperunbound.Intersect(lowernext, out result);
                        if (result != null && result.Size == 1)
                        {
                            IntersectionPointUpper = result.get_Item(0).XYZPoint;
                        }

                    }

                }
            }

            if (IntersectionPointLower == null || IntersectionPointUpper == null) return existing;



            return new Tuple<Curve, Curve>(Line.CreateBound(upper.GetEndPoint(0), IntersectionPointUpper), Line.CreateBound(lower.GetEndPoint(0), IntersectionPointLower));

        }



        /// <summary>
        /// Get the Insulation Layer from a wall
        /// </summary>
        public Tuple<Curve, Curve> getInsulationLayer(Document doc, Wall wall)
        {
            double offset = 0;
            LocationCurve c = (LocationCurve)wall.Location;

            CompoundStructure cs = wall.WallType.GetCompoundStructure();
            foreach (CompoundStructureLayer layer in cs.GetLayers())
            {
                if (layer.MaterialId != null && layer.MaterialId != ElementId.InvalidElementId)
                {
                    Material mat = doc.GetElement(layer.MaterialId) as Material;
                    if (mat.Name.ToLower().Contains("soft insulation"))
                    {

                        double ypos = (wall.Width / 2) - offset;

                        double rotationangle = -1.57079633;
                        double layeroffset = ypos - layer.Width;
                        if (ypos - layer.Width < 0) { rotationangle *= -1; ypos = (wall.Width / 2) * -1 + offset; layeroffset = ypos + layer.Width; }

                        if (wall.Flipped) rotationangle *= -1;


                        Transform tra = Transform.CreateRotationAtPoint(XYZ.BasisZ, rotationangle, c.Curve.GetEndPoint(0));
                        Curve rotated = c.Curve.CreateTransformed(tra);


                        XYZ LowerStartPoint = rotated.Evaluate(((layeroffset) / rotated.Length), true);
                        XYZ UpperStartPoint = null;

                        if (ypos < 0)
                        { 
                            Curve rotatedHelper = c.Curve.CreateTransformed(Transform.CreateRotationAtPoint(XYZ.BasisZ, rotationangle * -1, c.Curve.GetEndPoint(0)));
                            UpperStartPoint = rotatedHelper.Evaluate((ypos * -1 / rotatedHelper.Length), true);
                        }      
                        else
                            UpperStartPoint = rotated.Evaluate((ypos / rotated.Length), true);

                        XYZ v = c.Curve.GetEndPoint(0) - LowerStartPoint;
                        XYZ u = c.Curve.GetEndPoint(0) - UpperStartPoint;

                        Curve locationCurvelower = c.Curve.CreateTransformed(Transform.CreateTranslation(v));
                        Curve locationCurveupper = c.Curve.CreateTransformed(Transform.CreateTranslation(u));

                        return new Tuple<Curve, Curve>(locationCurveupper, locationCurvelower);
                    }
                }
                offset += layer.Width;
            }

            return null;
        }





        /// <summary>
        /// DrawInsulation
        /// </summary>
        public double drawInsulation(List<ElementId> grp, Document doc, Curve lower, double distance, Curve upper, ref bool leftwing, List<Interruption> Breaks, bool zigzag)
        {

            double dist = distance / lower.Length;
            if (dist > 1) return 100;

            XYZ P1 = lower.Evaluate(dist, true);
            IntersectionResultArray result = new IntersectionResultArray();
            XYZ normal = lower.GetCurveTangentToEndPoint(P1);

            SetComparisonResult scr = Line.CreateUnbound(P1, normal).Intersect(upper, out result);
            if (result == null || result.Size == 0)
            {
                if (dist > 0.5) return 100;
                else
                {
                    upper = upper.Clone();
                    upper.MakeUnbound();
                    scr = Line.CreateUnbound(P1, normal).Intersect(upper, out result);
                }

            }

            XYZ P3 = result.get_Item(0).XYZPoint;
            double height = P1.DistanceTo(P3);
            double r = height / 4;
            double distr = (distance + r) / lower.Length;
            if (distr > 1) return 100;

            foreach (Interruption interrupt in Breaks)
            {
                if (distr > lower.ComputeNormalizedParameter(interrupt.from) && distr < lower.ComputeNormalizedParameter(interrupt.to)) return r;
            }

            XYZ P2 = (distr < 1) ? lower.Evaluate(distr, true) : P1;


            double disth = (distance + height) / lower.Length;

            SetComparisonResult scr2 = Line.CreateUnbound(P2, lower.GetCurveTangentToEndPoint(P2)).Intersect(upper, out result);
            if (result == null || result.Size == 0) return 100;
            XYZ P4 = (P1 != P2) ? result.get_Item(0).XYZPoint : upper.GetEndPoint(1);

            if (zigzag)
                drawZigZag(grp, doc, P1, P2, P3, P4, leftwing);
            else
                drawSoftLoop(grp, doc, P1, P2, P3, P4, leftwing);

            if (leftwing) leftwing = false; else leftwing = true;
            
            return r;

        }

        /// <summary>
        /// Draw Zig Zag Loop
        /// </summary>
        public void drawZigZag(List<ElementId> elementIdsToGroup, Document doc, XYZ P1, XYZ P2, XYZ P3, XYZ P4, bool leftwing)
        {
            Line zig = (leftwing) ? Line.CreateBound(P1, P4) : Line.CreateBound(P3, P2);
            elementIdsToGroup.Add(doc.Create.NewDetailCurve(doc.ActiveView, zig).Id);
        }

        /// <summary>
        /// Draw Soft Insulation Loop
        /// </summary>
        public void drawSoftLoop(List<ElementId> elementIdsToGroup, Document doc, XYZ P1, XYZ P2, XYZ P3, XYZ P4, bool leftwing)
        {

            Curve bottom = (P1 != P2) ? Line.CreateBound(P1, P2) : null;
            Curve top = Line.CreateBound(P3, P4);
            Curve left = Line.CreateBound(P1, P3);
            Curve right = Line.CreateBound(P2, P4);

            XYZ bottom02 = (P1 != P2) ? bottom.Evaluate(0.2, true) : P1;
            XYZ top02 = top.Evaluate(0.2, true);

            XYZ bottom08 = (P1 != P2) ? bottom.Evaluate(0.8, true) : P1;
            XYZ top08 = top.Evaluate(0.8, true);


            Line verticalAxisQ1 = Line.CreateBound(bottom08, top08);
            Line verticalAxisQ2 = Line.CreateBound(bottom02, top02);


            XYZ A = (leftwing) ? P1 : P3;
            XYZ B = (leftwing) ? right.Evaluate(0.25, true) : right.Evaluate(0.75, true);
            XYZ C = (leftwing) ? verticalAxisQ1.Evaluate(0.4, true) : verticalAxisQ1.Evaluate(0.6, true);
            XYZ D = (leftwing) ? verticalAxisQ2.Evaluate(0.6, true) : verticalAxisQ2.Evaluate(0.4, true);
            XYZ E = (leftwing) ? left.Evaluate(0.75, true) : left.Evaluate(0.25, true);
            XYZ F = (leftwing) ? P4 : P2;

            Arc a = Arc.Create(A, C, B);
            Line b = Line.CreateBound(C, D);
            Arc c = Arc.Create(D, F, E);

            elementIdsToGroup.Add(doc.Create.NewDetailCurve(doc.ActiveView, a).Id);
            elementIdsToGroup.Add(doc.Create.NewDetailCurve(doc.ActiveView, b).Id);
            elementIdsToGroup.Add(doc.Create.NewDetailCurve(doc.ActiveView, c).Id);

        }


    }




    public class Interruption
    {
        public double from;
        public double to;

        public Interruption(Curve a, XYZ point, double width)
        {
            double parameter = a.GetEndPoint(0).DistanceTo(point);
            this.from = parameter - width / 2;
            this.to = parameter + width / 2;
        }

        public void Extend(double left)
        {
            this.from += left;
            this.to += left;
        }
    }


    public static class Extensions
    {
        public static Line CreateLineSafe(this Document doc, XYZ p1, XYZ p2)
        {
            
            if (p1.DistanceTo(p2) > doc.Application.ShortCurveTolerance)
            {
                return Line.CreateBound(p1, p2);
            }

            return null;
        }
    }
}
