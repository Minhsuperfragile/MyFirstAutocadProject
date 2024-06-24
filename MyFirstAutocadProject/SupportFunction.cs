using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Windows.Forms;

namespace MyFirstAutocadProject {
    internal class SupportFunction {

        #region Beginner Functions
        public bool isLayerExisted(Transaction trans, LayerTable lt, string layerName) {
            LayerTableRecord ltr;
            foreach (ObjectId ob in lt) {
                ltr = trans.GetObject(ob, OpenMode.ForRead) as LayerTableRecord;
                if (ltr.Name == layerName) { return true; }
            }
            return false;
        }

        public void newLayer(string layerName) {
            ObjectId layerID = Commands.db.LayerTableId; // Get all Layer
            Transaction trans = Commands.db.TransactionManager.StartTransaction(); // start transaction
            LayerTable layerTable = trans.GetObject(layerID, OpenMode.ForWrite) as LayerTable;

            int dup = 1;

            while (isLayerExisted(trans, layerTable, layerName)) {
                layerName = "Minh(" + dup + ")";
                dup++;
            }

            LayerTableRecord newLayer = new LayerTableRecord() { Name = layerName };
            layerTable.Add(newLayer);
            trans.AddNewlyCreatedDBObject(newLayer, true);

            trans.Commit();
        }
        #endregion

        #region Find elevation support functions
        public Polyline[] GetAllPolylineInLayer() { // this function try to collect all pline that represent elevation
            Polyline[] resultSet = { };
            Transaction tran = Commands.db.TransactionManager.StartTransaction();

            PromptEntityOptions prtOpt = new PromptEntityOptions("\nSelect a polyline in your desired layer");
            prtOpt.SetRejectMessage("\nInvalid");
            prtOpt.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult selectedPolyLine = Commands.ed.GetEntity(prtOpt);
            if (selectedPolyLine.Status == PromptStatus.OK) {
                Polyline resultPolyline = tran.GetObject(selectedPolyLine.ObjectId, OpenMode.ForRead) as Polyline;
                ObjectId layer = resultPolyline.LayerId;

                BlockTable bt = tran.GetObject(Commands.db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tran.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId entityId in modelSpace) {
                    Polyline pl = tran.GetObject(entityId, OpenMode.ForRead) as Polyline; // this is null sometimes
                    if (pl != null) {
                        if (pl.LayerId == layer) {
                            Array.Resize(ref resultSet, resultSet.Length + 1);
                            resultSet[resultSet.Length - 1] = pl;
                            //Commands.ed.WriteMessage("\nElevation of Polyline in layer: " + pl.Elevation);
                        }
                    }
                }
            }
            Commands.ed.WriteMessage("\nThere are " + resultSet.Length + " polylines in that layer");
            tran.Commit();
            return resultSet;
        }

        public double GetEnhancedPrecisionElevation(double[,] elements) {
            if (elements[1, 1] == 0) {
                if (elements[0, 1] == 0) {
                    throw new System.Exception("Failed to find distance to point");
                }
                return elements[0, 0];
            }
            else if (elements[0, 1] == 0) {
                return elements[1, 0];
            }

            double distanceRatio;
            double elevationDiff = elements[0, 0] - elements[1, 0];
            double lower;
            //double upper;
            //Commands.ed.WriteMessage("\nDistance 1st: " + elements[0, 0] + " " + elements[0, 1] + "\nDistance 2d: " + elements[1, 0] + " " + elements[1, 1]);

            if (elevationDiff < 0) {
                lower = elements[0, 0];
                //upper = elements[1, 0];
                distanceRatio = elements[0, 1] / (elements[1, 1] + elements[0, 1]);
            }
            else if (elevationDiff > 0) {
                lower = elements[1, 0];
                //upper = elements[0, 0];
                distanceRatio = elements[1, 1] / (elements[1, 1] + elements[0, 1]);
            }
            else {
                return elements[0, 0];
            }
            double result = lower + Math.Abs(elevationDiff) * distanceRatio;

            if (result == 0) { Commands.ed.WriteMessage("\nDistance 1st: " + elements[0, 0] + " " + elements[0, 1] + "\nDistance 2d: " + elements[1, 0] + " " + elements[1, 1]); }

            return result;
        }

        public double[,] TwoClosestPolyLineToPoint(Point3d point, Polyline[] polylineList) {
            double minDistance = double.MaxValue;
            //Polyline nearestLine;

            double[,] result = { { 0, 0 }, { 0, 0 } }; // [


            try {
                foreach (Polyline polyline in polylineList) {
                    double distance = GetDistanceToPolyline(point, polyline);
                    if (distance == -1) {
                        //Commands.ed.WriteMessage("\n -1");
                        //throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NotHandled, "Can't find points");
                        continue;
                    }
                    if (distance < minDistance) {
                        result[1, 0] = result[0, 0];
                        result[1, 1] = result[0, 1];
                        //nearestLine = polyline;
                        minDistance = distance;
                        result[0, 0] = polyline.Elevation;
                        result[0, 1] = minDistance;
                    }
                    else if (distance < result[1, 1] && distance > minDistance) {
                        result[1, 0] = polyline.Elevation;
                        result[1, 1] = distance;
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex) {
                Commands.ed.WriteMessage(ex.ToString());
            }

            //Commands.ed.WriteMessage("\nElevation is: " + nearestLine.Elevation);
            return result;
        }

        public double GetDistanceToPolyline(Point3d point, Polyline polyline) {
            double currentDistance;
            double minD = double.MaxValue;
            // Iterate through each segment of the polyline
            for (int i = 0; i < polyline.NumberOfVertices - 1; i++) {
                Point3d startPoint = polyline.GetPoint3dAt(i);
                Point3d endPoint = polyline.GetPoint3dAt(i + 1);

                try {
                    // Calculate the distance between the given point and the closest point on the segment
                    currentDistance = GetClosestPointOnSegment(new Point2d(point.X, point.Y),
                        new Point2d(startPoint.X, startPoint.Y),
                        new Point2d(endPoint.X, endPoint.Y));

                    if (currentDistance == -1) { continue; }
                    // Update the minimum distance
                    if (currentDistance < minD) {
                        return currentDistance;
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex) {
                    Commands.ed.WriteMessage(ex.ToString());
                    Console.WriteLine(ex.ToString());
                }

            }

            return -1;
        }

        public double GetClosestPointOnSegment(Point2d target, Point2d startPoint, Point2d endPoint) {
            if (!IsNearToPoint(target, startPoint, endPoint, 10)) {
                return -1;
            }

            if (IsObtuseTriangle(target, startPoint, endPoint)) {
                return -1;
            }

            double h2 = distanceToLine(target, startPoint, endPoint);
            return h2;

        }

        public bool IsObtuseTriangle(Point2d A, Point2d B, Point2d C) { // this function allows A be obtuse angle, but not B and C
            double AB = A.GetDistanceTo(B);
            double BC = B.GetDistanceTo(C);
            double AC = A.GetDistanceTo(C);
            if (AC * AC > AB * AB + BC * BC) { return true; } // angle B is obtuse
            else if (AB * AB > AC * AC + BC * BC) { return true; } // angle C is obtuse
            else { return false; }
        }

        public double DistanceBetween(Vector3d v1, Vector3d v2) {
            double result = Math.Sqrt(Math.Pow(v1.X - v2.X, 2) + Math.Pow(v1.Y - v2.Y, 2) + Math.Pow(v1.Z - v2.Z, 2));
            return result;
        }

        public double heightOfTriangular(Point2d target, Point2d B, Point2d A) {

            double area = Math.Abs(target.X * (A.Y - B.Y) + B.X * (target.Y - A.Y) + A.X * (target.Y - B.Y)) / 2;
            //Commands.ed.WriteMessage(" " + area);
            return area / B.GetDistanceTo(A);
        }

        public double distanceToLine(Point2d target, Point2d start, Point2d B) {
            Vector2d direction = B - start;
            double t = direction.DotProduct(target - start) / Math.Pow(direction.Length, 2);
            Vector2d b = new Vector2d(
                start.X - target.X + t * direction.X,
                start.Y - target.Y + t * direction.Y
                    );

            return b.Length;
        }

        public bool IsNearToPoint(Point2d target, Point2d A, Point2d B, double range) {
            if (A.X < target.X - range || A.X > target.X + range) { return false; }
            if (B.X < target.X - range || B.X > target.X + range) { return false; }
            if (A.Y < target.Y - range || A.Y > target.Y + range) { return false; }
            if (B.Y < target.Y - range || B.Y > target.Y + range) { return false; }
            return true;
        }

        #endregion

        #region CalculateIntersectPoint support functions
        public DBObjectCollection GetAllLinesInLayer() {
            Transaction tran = Commands.db.TransactionManager.StartTransaction();
            PromptEntityOptions prtOpt = new PromptEntityOptions("\nSelect a line in your desired layer");
            prtOpt.SetRejectMessage("\nInvalid");
            prtOpt.AddAllowedClass(typeof(Line), true);

            PromptEntityResult selectedLines = Commands.ed.GetEntity(prtOpt);
            if (selectedLines.Status == PromptStatus.OK) {
                Line inputLine = tran.GetObject(selectedLines.ObjectId, OpenMode.ForRead) as Line;
                ObjectId layer = inputLine.LayerId;

                LayerTableRecord ltr = tran.GetObject(layer, OpenMode.ForRead) as LayerTableRecord;

                // Create a selection filter for lines
                TypedValue[] filterList = new TypedValue[]
                {
                        new TypedValue((int)DxfCode.Start, "LINE"),
                        new TypedValue((int)DxfCode.LayerName, ltr.Name)
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                // Select all lines on the given layer
                PromptSelectionResult selectionResult = Commands.ed.SelectAll(filter);
                if (selectionResult.Status == PromptStatus.OK) {
                    SelectionSet selectionSet = selectionResult.Value;
                    ObjectId[] objectIds = selectionSet.GetObjectIds();

                    // Collect the line objects
                    DBObjectCollection lines = new DBObjectCollection();
                    foreach (ObjectId objectId in objectIds) {
                        if (objectId.ObjectClass == RXClass.GetClass(typeof(Line))) {
                            Line line = tran.GetObject(objectId, OpenMode.ForRead) as Line;
                            lines.Add(line);
                            Commands.ed.WriteMessage("\nAdded a line to list");
                        }
                    }
                    return lines;
                }

            }
            return null;
        }
        public Point3d GetLineIntersectionWithPLine(Polyline polyLine, Line line) {
            Point3dCollection points = new Point3dCollection();

            for (int i = 0; i < polyLine.NumberOfVertices - 1; i++) {
                Point3d point1 = polyLine.GetPoint3dAt(i);
                Point3d point2 = polyLine.GetPoint3dAt(i + 1);

                Line segment = new Line(point1, point2);
                segment.IntersectWith(line, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero);

            }
            return points[0];
        }

        public double GetLengthAt(int index, Polyline pl) {
            double length = 0;
            if (index > pl.NumberOfVertices) {
                return -1;
            }
            for (int i = 0; i < index; i++) {
                if (pl.GetSegmentType(i) == SegmentType.Line) {
                    length += pl.GetLineSegmentAt(i).Length;
                }
                else if (pl.GetSegmentType(i) == SegmentType.Arc) {
                    CircularArc3d arc = pl.GetArcSegmentAt(i);
                    length += arc.Radius * (arc.EndAngle - arc.StartAngle);
                }
            }
            return length;
        }
        #endregion

        #region Read All Text

        #endregion

        #region General Utils
        public String OpenFile() {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            fileDialog.InitialDirectory = @"C:\";
            fileDialog.RestoreDirectory = true;

            if (fileDialog.ShowDialog() == DialogResult.OK) {
                String fileStream = fileDialog.FileName;
                try {
                    using (StreamReader reader = new StreamReader(fileStream)) {
                        //fileContent = reader.ReadToEnd();
                    }
                    Commands.ed.WriteMessage("\nSelected :" + fileStream);
                    return fileStream;
                }
                catch {
                    Commands.ed.WriteMessage("Can't open");
                    return string.Empty;
                }
            }
            else {
                Commands.ed.WriteMessage("\nFailed to get file");
                return string.Empty;
            }
        }
        #endregion
    }
}
