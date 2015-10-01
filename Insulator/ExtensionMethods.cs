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
    public static class ExtensionMethods
    {
        /// <summary>
        /// Get Normal at a point of a curve
        /// </summary>
        /// <param name="Point">Point to evaluate</param>
        /// <returns>Normal</returns>
        public static XYZ GetCurveTangentToEndPoint(this Curve curve, XYZ Point)
        {
            // Tessellate the curve
            IList<XYZ> pts = curve.Tessellate();

            // Get the endpoint
            XYZ closestPoint = curve.GetEndPoint(1);

            // Walk through tessellation points and get the closest point to the input parameter
            foreach (XYZ pt in pts)
            {
                if (pt.DistanceTo(Point) < closestPoint.DistanceTo(Point) && pt.DistanceTo(Point) > 0.0001) closestPoint = pt;
            }
            
            // If the actual endpoint is almost the same than the input parameter
            // us the whole curve as a tangent definition
            if (closestPoint.DistanceTo(Point) < 0.0001)
            {
                Point = curve.GetEndPoint(0);
                closestPoint = curve.GetEndPoint(1);
            }

            // Draw the tangent and return its normal
            return GetCurveNormal(Line.CreateBound(Point, closestPoint));
        }

        /// <summary>
        /// Get Curves Normal
        /// From Jeremy Tammik
        /// </summary>
        /// <param name="curve">Curve</param>
        /// <returns>Normal Vector</returns>
        public static XYZ GetCurveNormal(Curve curve)
        {
            IList<XYZ> pts = curve.Tessellate();
            int n = pts.Count;

            XYZ p = pts[0];
            XYZ q = pts[n - 1];
            XYZ v = q - p;
            XYZ w, normal = null;

            if (2 == n)
            {

                // for non-vertical lines, use Z axis to 
                // span the plane, otherwise Y axis:

                double dxy = Math.Abs(v.X) + Math.Abs(v.Y);

                w = (dxy > 0.001)
                  ? XYZ.BasisZ
                  : XYZ.BasisY;

                normal = v.CrossProduct(w).Normalize();
            }
            else
            {
                int i = 0;
                while (++i < n - 1)
                {
                    w = pts[i] - p;
                    normal = v.CrossProduct(w);
                    if (!normal.IsZeroLength())
                    {
                        normal = normal.Normalize();
                        break;
                    }
                }

            }
            return normal;
        }
    }
}
