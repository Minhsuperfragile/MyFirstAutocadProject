// adding autocad api
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using System;
using System.IO;
using System.Reflection;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyFirstAutocadProject {
    public class Commands {
        public static DocumentCollection dc = AcAp.DocumentManager;
        public static Database db = AcAp.DocumentManager.MdiActiveDocument.Database;
        public static Editor ed = AcAp.DocumentManager.MdiActiveDocument.Editor;

        private SupportFunction func = new SupportFunction();
        public static string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "log.txt");

        #region Command Function
        [CommandMethod("NTMTransaction")]
        public void cmdTransaction() { // transaction must be enabled to read or write file 
            ObjectId layerID = db.LayerTableId;

            Transaction trans = db.TransactionManager.StartTransaction(); // start transaction
            LayerTable layerTable = trans.GetObject(layerID, OpenMode.ForRead) as LayerTable;

            int count = 0;
            foreach (ObjectId ob in layerTable) {
                count++;
            }

            trans.Commit(); // close transaction

            ed.WriteMessage("\nCo " + count + " layer trong ban ve");
        }

        [CommandMethod("NTMCreateLayer")]
        public void cmdCreateLayer() {
            func.newLayer("Minhdeptrai");
        }

        [CommandMethod("NTMCreateLine")]
        public void cmdCreateLine() {
            Transaction trans = db.TransactionManager.StartTransaction();

            ObjectId blockID = db.CurrentSpaceId;
            BlockTableRecord currentSpace = trans.GetObject(blockID, OpenMode.ForWrite) as BlockTableRecord;

            PromptPointOptions promptPointOptions = new PromptPointOptions("\nSelect a point") {
                AllowArbitraryInput = false,
                AllowNone = true
            };

            PromptPointResult inputPoint = ed.GetPoint(promptPointOptions);
            if (inputPoint.Status != PromptStatus.OK) return;

            int lengthX = 100, lengthY = 200;

            Point3d horizontalEndPoint = new Point3d(inputPoint.Value.X, inputPoint.Value.Y + lengthY, 0);
            Point3d verticalEndPoint = new Point3d(inputPoint.Value.X + lengthX, inputPoint.Value.Y, 0);

            Line verticalLine = new Line(inputPoint.Value, verticalEndPoint);
            Line horizontalLine = new Line(inputPoint.Value, horizontalEndPoint);

            currentSpace.AppendEntity(verticalLine);
            trans.AddNewlyCreatedDBObject(verticalLine, true);

            currentSpace.AppendEntity(horizontalLine);
            trans.AddNewlyCreatedDBObject(horizontalLine, true);

            trans.Commit();
        }

        [CommandMethod("NTMAddText")]
        public void cmdAddText() {
            Transaction trans = db.TransactionManager.StartTransaction();

            ObjectId blockID = db.CurrentSpaceId;
            BlockTableRecord currentSpace = trans.GetObject(blockID, OpenMode.ForWrite) as BlockTableRecord;

            try {
                Point3d startTextPoint = new Point3d(10, 10, 0);

                DBText newText = new DBText() {
                    Position = startTextPoint,
                    Height = 10,
                    TextString = "Hello World"
                };

                currentSpace.AppendEntity(newText);
                trans.AddNewlyCreatedDBObject(newText, true);
            }
            catch {
                ed.WriteMessage("\nFail to create text");
            }
            trans.Commit();
        }

        [CommandMethod("NTMSelect")]
        public void cmdSelectEntity() {
            PromptEntityOptions prtSelOpt = new PromptEntityOptions("\nSelect a polyline:") {
                AllowNone = true,
                AllowObjectOnLockedLayer = true
            };
            prtSelOpt.SetRejectMessage("\nNot a line!");
            prtSelOpt.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult prtResult = ed.GetEntity(prtSelOpt);
            if (prtResult.Status != PromptStatus.OK) { return; }

            Transaction trans = db.TransactionManager.StartTransaction();
            Polyline polylineObj = trans.GetObject(prtResult.ObjectId, OpenMode.ForRead) as Polyline;

            ed.WriteMessage("Polyline have " + polylineObj.NumberOfVertices + " vertices");

            trans.Commit();
        }

        [CommandMethod("NTMSeletion")]
        public void cmdSelection() {
            PromptSelectionOptions prptOpts = new PromptSelectionOptions() {
                AllowDuplicates = false,
                RejectObjectsFromNonCurrentSpace = false,
                RejectPaperspaceViewport = true,
                RejectObjectsOnLockedLayers = true,
                MessageForAdding = "\nAdd Line",
                MessageForRemoval = "\nRemove Line"
            };

            TypedValue[] typedValue = new TypedValue[1];
            typedValue[0] = new TypedValue((int)DxfCode.Start, "LINE");
            SelectionFilter setFilter = new SelectionFilter(typedValue);

            PromptSelectionResult prptResult = ed.GetSelection(prptOpts, setFilter);
            if (prptResult.Status != PromptStatus.OK) { return; }

            Transaction trans = db.TransactionManager.StartTransaction();

            try {
                SelectionSet selSet = prptResult.Value;
                Line lineObj;
                double length = 0.0;
                foreach (SelectedObject selected in selSet) {
                    lineObj = trans.GetObject(selected.ObjectId, OpenMode.ForRead) as Line;
                    length += lineObj.Length;
                }
                ed.WriteMessage("\nTotal length is: " + length);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex) {
                AcAp.ShowAlertDialog("Failed");
                ed.WriteMessage(ex.Message);
            }

            trans.Commit();
        }

        [CommandMethod("NTMGetElevationOfPolyLineLinear")]
        public void cmdFindNearestPolyLineLinear() {
            string saveFilePath = func.OpenFile();

            PromptEntityOptions prptOpts = new PromptEntityOptions("\nSelect a PolyLine:") {
                AllowNone = true,
                AllowObjectOnLockedLayer = true,
            };
            prptOpts.SetRejectMessage("\nNot a line!");
            prptOpts.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult prptResult = ed.GetEntity(prptOpts);
            if (prptResult.Status != PromptStatus.OK) { return; }

            Transaction trans = db.TransactionManager.StartTransaction();
            Polyline inputPolyLine = trans.GetObject(prptResult.ObjectId, OpenMode.ForRead) as Polyline;
            trans.Commit();

            Polyline[] polylines = func.GetAllPolylineInLayer();

            for (int i = 0; i < inputPolyLine.NumberOfVertices - 1; i++) {
                double length = inputPolyLine.GetPoint3dAt(i).DistanceTo(inputPolyLine.GetPoint3dAt(i + 1));
                //double kLength = length / 100;
                double ratio = length / 100;

                Vector3d startPoint = inputPolyLine.GetPoint3dAt(i).GetAsVector();
                Vector3d endPoint = inputPolyLine.GetPoint3dAt(i + 1).GetAsVector();

                Vector3d direction = (endPoint - startPoint).GetNormal();

                Vector3d currentPoint = startPoint;

                for (int j = 0; j <= 100; j++) {

                    String pointInfo = currentPoint.X.ToString() + "," + currentPoint.Y.ToString() + "," +
                    func.GetEnhancedPrecisionElevation(
                        func.TwoClosestPolyLineToPoint(
                        new Point3d(
                            currentPoint.X,
                            currentPoint.Y,
                            0),
                            polylines
                        ));

                    currentPoint += direction.MultiplyBy(ratio);
                    try {
                        using (StreamWriter wr = File.AppendText(saveFilePath)) {
                            wr.WriteLine(pointInfo);
                            wr.Close();
                        }
                    }
                    catch {
                        ed.WriteMessage("\nFailed to open file");
                        return;
                    }
                    ed.WriteMessage("\n" + pointInfo);
                }

            }

            AcAp.ShowAlertDialog("Done!");
        }

        [CommandMethod("NTMGetElevationOfPolylineBinary")]
        public void cmdFindNearestPolylineBinary() {
            #region get user input
            string saveFilePath = func.OpenFile();

            PromptEntityOptions prptOpts = new PromptEntityOptions("\nSelect a PolyLine:") {
                AllowNone = true,
                AllowObjectOnLockedLayer = true
            };
            prptOpts.SetRejectMessage("\nNot a line!");
            prptOpts.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult prptResult = ed.GetEntity(prptOpts);
            if (prptResult.Status != PromptStatus.OK) { return; }

            Transaction trans = db.TransactionManager.StartTransaction();
            Polyline inputPolyLine = trans.GetObject(prptResult.ObjectId, OpenMode.ForRead) as Polyline;
            trans.Commit();

            Polyline[] polylines = func.GetAllPolylineInLayer();
            #endregion

            for (int i = 0; i < inputPolyLine.NumberOfVertices - 1; i++) {
                Vector3d startPointVector = inputPolyLine.GetPoint3dAt(i).GetAsVector();
                Vector3d endPointVector = inputPolyLine.GetPoint3dAt(i + 1).GetAsVector();

                Vector3d direction = (endPointVector - startPointVector).GetNormal();

                Vector3d currentPointVector = startPointVector;
                Vector3d nextPointVector;

                double stackRealDistance = 0, stackFlatDistance = 0;
                try {
                    using (StreamWriter wr = File.AppendText(saveFilePath)) {
                        wr.WriteLine(currentPointVector.X + "," + currentPointVector.Y + "," + func.GetEnhancedPrecisionElevation(func.TwoClosestPolyLineToPoint(new Point3d(currentPointVector.X, currentPointVector.Y, 0), polylines)) + ",0,0,0,0");
                        wr.Close();
                    }
                }
                catch {
                    ed.WriteMessage("\nFailed to open file");
                    return;
                }

                while (direction.DotProduct(endPointVector - currentPointVector) > 0) {
                    double realDistance, flatDistance;
                    double lambda = 4; // determine the start value of distance to move 

                    double currentElevation = func.GetEnhancedPrecisionElevation(func.TwoClosestPolyLineToPoint(new Point3d(currentPointVector.X, currentPointVector.Y, 0), polylines));
                    Point3d currentPoint = new Point3d(currentPointVector.X, currentPointVector.Y, currentElevation);

                    nextPointVector = currentPointVector + direction.MultiplyBy(lambda);

                    Vector3d auxPointVector = currentPointVector;

                    int flipper;
                    int count = 0;
                    while (true) {
                        count++;

                        double nextElevation = func.GetEnhancedPrecisionElevation(func.TwoClosestPolyLineToPoint(new Point3d(nextPointVector.X, nextPointVector.Y, 0), polylines));
                        Point3d nextPoint = new Point3d(nextPointVector.X, nextPointVector.Y, nextElevation);

                        realDistance = currentPoint.DistanceTo(nextPoint);
                        flatDistance = func.DistanceBetween(currentPointVector, nextPointVector);

                        if (count > 100) {
                            stackFlatDistance += flatDistance;
                            stackRealDistance += realDistance;
                            String info = nextPoint.X + "," + nextPoint.Y + "," + nextElevation + "," + flatDistance + "," + stackFlatDistance + "," + realDistance + "," + stackRealDistance;
                            using (StreamWriter wr = File.AppendText(saveFilePath)) {
                                wr.WriteLine(info + ",Overloaded");
                                wr.Close();
                            }
                            break;
                        }

                        if (realDistance < 6.9) {
                            lambda = func.DistanceBetween(auxPointVector, nextPointVector);
                            auxPointVector = nextPointVector;

                            //if (flipper != 1 || count >= 10) {  }
                            flipper = 1;
                            nextPointVector += direction.MultiplyBy(flipper * lambda / 2);

                        }
                        else if (realDistance > 7.1) {
                            lambda = func.DistanceBetween(auxPointVector, nextPointVector);
                            auxPointVector = nextPointVector;

                            //if (flipper != -1 || count >= 10) {  }
                            flipper = -1;

                            nextPointVector += direction.MultiplyBy(flipper * lambda / 2);
                        }
                        else {
                            stackFlatDistance += flatDistance;
                            stackRealDistance += realDistance;

                            String info = nextPoint.X + "," + nextPoint.Y + "," + nextElevation + "," + flatDistance + "," + stackFlatDistance + "," + realDistance + "," + stackRealDistance;

                            using (StreamWriter wr = File.AppendText(saveFilePath)) {
                                wr.WriteLine(info);
                                wr.Close();
                            }
                            break;
                        }

                    }
                    currentPointVector = nextPointVector;
                }
            }
            AcAp.ShowAlertDialog("Done!");
        }

        [CommandMethod("NTMGetElevationOfLineBinary")]
        public void cmdFindNearestLineBinary() {
            #region get user input
            string saveFilePath = func.OpenFile();

            PromptEntityOptions prptOpts = new PromptEntityOptions("\nSelect a Line:") {
                AllowNone = true,
                AllowObjectOnLockedLayer = true
            };
            prptOpts.SetRejectMessage("\nNot a line!");
            prptOpts.AddAllowedClass(typeof(Line), true);

            PromptEntityResult prptResult = ed.GetEntity(prptOpts);
            if (prptResult.Status != PromptStatus.OK) { return; }

            Transaction trans = db.TransactionManager.StartTransaction();
            Line inputPolyLine = trans.GetObject(prptResult.ObjectId, OpenMode.ForRead) as Line;
            trans.Commit();

            Polyline[] polylines = func.GetAllPolylineInLayer();
            #endregion

            Vector3d startPointVector = inputPolyLine.StartPoint.GetAsVector();
            Vector3d endPointVector = inputPolyLine.EndPoint.GetAsVector();

            Vector3d direction = (endPointVector - startPointVector).GetNormal();

            Vector3d currentPointVector = startPointVector;
            Vector3d nextPointVector;

            double stackRealDistance = 0, stackFlatDistance = 0;
            try {
                using (StreamWriter wr = File.AppendText(saveFilePath)) {
                    wr.WriteLine(currentPointVector.X + "," + currentPointVector.Y + "," + func.GetEnhancedPrecisionElevation(func.TwoClosestPolyLineToPoint(new Point3d(currentPointVector.X, currentPointVector.Y, 0), polylines)) + ",0,0,0,0");
                    wr.Close();
                }
            }
            catch {
                ed.WriteMessage("\nFailed to open file");
                return;
            }
            double lambda;
            while (direction.DotProduct(endPointVector - currentPointVector) > 0) {
                double realDistance, flatDistance;
                lambda = 4;// determine the start value of distance to move 

                double currentElevation = func.GetEnhancedPrecisionElevation(func.TwoClosestPolyLineToPoint(new Point3d(currentPointVector.X, currentPointVector.Y, 0), polylines));
                Point3d currentPoint = new Point3d(currentPointVector.X, currentPointVector.Y, currentElevation);

                nextPointVector = currentPointVector + direction.MultiplyBy(lambda);

                Vector3d auxPointVector = currentPointVector;

                int flipper;
                int count = 0;
                while (true) { //start binary search
                    count++;

                    double nextElevation = func.GetEnhancedPrecisionElevation(func.TwoClosestPolyLineToPoint(new Point3d(nextPointVector.X, nextPointVector.Y, 0), polylines));
                    Point3d nextPoint = new Point3d(nextPointVector.X, nextPointVector.Y, nextElevation);

                    realDistance = currentPoint.DistanceTo(nextPoint);
                    flatDistance = func.DistanceBetween(currentPointVector, nextPointVector);


                    if (count > 100) {
                        stackFlatDistance += flatDistance;
                        stackRealDistance += realDistance;
                        String info = nextPoint.X + "," + nextPoint.Y + "," + nextElevation + "," + flatDistance + "," + stackFlatDistance + "," + realDistance + "," + stackRealDistance;
                        using (StreamWriter wr = File.AppendText(saveFilePath)) {
                            wr.WriteLine(info + ",Overloaded");
                            wr.Close();
                        }
                        break;
                    }

                    if (realDistance < 6.9) {
                        lambda = func.DistanceBetween(auxPointVector, nextPointVector);
                        auxPointVector = nextPointVector;

                        //if (flipper != 1 || count >= 10) {  }
                        flipper = 1;

                        nextPointVector += direction.MultiplyBy(flipper * lambda / 2);
                    }
                    else if (realDistance > 7.1) {
                        lambda = func.DistanceBetween(auxPointVector, nextPointVector);
                        auxPointVector = nextPointVector;

                        //if (flipper != -1 || count >= 10) {  }
                        flipper = -1;

                        nextPointVector += direction.MultiplyBy(flipper * lambda / 2);
                    }
                    else {
                        stackFlatDistance += flatDistance;
                        stackRealDistance += realDistance;

                        String info = nextPoint.X + "," + nextPoint.Y + "," + nextElevation + "," + flatDistance + "," + stackFlatDistance + "," + realDistance + "," + stackRealDistance;

                        using (StreamWriter wr = File.AppendText(saveFilePath)) {
                            wr.WriteLine(info);
                            wr.Close();
                        }
                        break;
                    }

                }

                currentPointVector = nextPointVector;
            }



            AcAp.ShowAlertDialog("Done!");
        }

        [CommandMethod("NTMCalculateIntersectionPoint")]
        public void cmdCalculateIntersectionPoint() {
            PromptEntityOptions prptOpts = new PromptEntityOptions("\nSelect a PolyLine:") {
                AllowNone = true,
                AllowObjectOnLockedLayer = true,
            };
            prptOpts.SetRejectMessage("\nNot a line!");
            prptOpts.AddAllowedClass(typeof(Polyline), true);

            ed.WriteMessage("Ark for input PolyLine");

            PromptEntityResult prptResult = ed.GetEntity(prptOpts);
            if (prptResult.Status != PromptStatus.OK) { return; }

            Transaction trans = db.TransactionManager.StartTransaction();
            Polyline inputPolyLine = trans.GetObject(prptResult.ObjectId, OpenMode.ForWrite) as Polyline;
            trans.Commit();

            ed.WriteMessage("Got the PolyLine");

            DBObjectCollection crossLines = func.GetAllLinesInLayer();
            ed.WriteMessage("Got the list<Line> to CalculateIntersectionPoint function " + crossLines.Count);

            foreach (Line line in crossLines) {
                Point3d intersectionPoint;
                LineSegment3d crossLineSegment = new LineSegment3d(line.StartPoint, line.EndPoint);
                for (int i = 0; i < crossLines.Count - 1; i++) {
                    if (inputPolyLine.GetSegmentType(i) == SegmentType.Line) {
                        LineSegment3d segment = inputPolyLine.GetLineSegmentAt(i);
                        Point3d[] ips = segment.IntersectWith(crossLineSegment);
                        if (ips == null) { continue; }
                        else {
                            intersectionPoint = ips[0];
                            ed.WriteMessage(intersectionPoint.ToString());
                            inputPolyLine.AddVertexAt(i, new Point2d(intersectionPoint.X, intersectionPoint.Y), 0, 0, 0);
                            break;
                        }
                    }
                    else if (inputPolyLine.GetSegmentType(i) == SegmentType.Arc) {
                        CircularArc3d segment = inputPolyLine.GetArcSegmentAt(i);
                        Point3d[] ips = segment.IntersectWith(crossLineSegment);
                        if (ips == null) { continue; }
                        else {
                            intersectionPoint = ips[0];
                            ed.WriteMessage(intersectionPoint.ToString());
                            inputPolyLine.AddVertexAt(i, new Point2d(intersectionPoint.X, intersectionPoint.Y), 0, 0, 0);
                            break;
                        }
                    }
                }
            }
            ed.WriteMessage("Complete calculate!");
        }

        [CommandMethod("NTMWriteCoordinate")]
        public void cmdWriteCoordinate() {
            string saveFilePath = func.OpenFile();

            // read file:
            try {
                StreamReader reader = new StreamReader(saveFilePath);

                String line;
                int count = 1;

                int TSIZE = 3;


                double kccdX = 100;
                double kccdY = 100;
                double kclX;
                double kclY = 112.5;

                for (line = reader.ReadLine(); line != null; line = reader.ReadLine()) {
                    String[] linesList = line.Split(',');
                    if (count != 1) {
                        kccdX += float.Parse(linesList[1]);
                    }

                    // kccd text
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        DBText kccd = new DBText();
                        kccd.TextString = linesList[0];
                        kccd.Height = TSIZE;
                        kccd.Rotation = 1.571;
                        kccd.WidthFactor = 1;
                        kccd.Position = new Point3d(kccdX + 1.5, kccdY - 9, 0);
                        modelSpace.AppendEntity(kccd);
                        tr.AddNewlyCreatedDBObject(kccd, true);

                        ed.WriteMessage("created kccd \n");
                        tr.Commit();
                    }
                    // cao do text
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        DBText cd = new DBText();

                        cd.TextString = linesList[2];
                        cd.Height = TSIZE;
                        cd.Rotation = 1.571;
                        cd.WidthFactor = 1;
                        cd.Position = new Point3d(kccdX + 1.5, kccdY + 23, 0);
                        modelSpace.AppendEntity(cd);
                        tr.AddNewlyCreatedDBObject(cd, true);

                        tr.Commit();
                        ed.WriteMessage("created cd = " + cd.TextString + " \n");
                    }
                    // phan chia kcl
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        Line verticalLine = new Line(new Point3d(kccdX, kccdY + 7.5, 0), new Point3d(kccdX, kccdY + 17.5, 0));
                        modelSpace.AppendEntity(verticalLine);
                        tr.AddNewlyCreatedDBObject(verticalLine, true);
                        tr.Commit();
                    }
                    // stt line
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        double sttY;
                        if (count % 2 == 0) {
                            sttY = kccdY - 27;
                        }
                        else {
                            sttY = kccdY - 17;
                        }
                        Line markerLine = new Line(new Point3d(kccdX, sttY, 0), new Point3d(kccdX, sttY + 5, 0));
                        modelSpace.AppendEntity(markerLine);
                        tr.AddNewlyCreatedDBObject(markerLine, true);

                        tr.Commit();
                    }
                    //cao do line
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        Line cdLine = new Line(new Point3d(kccdX, kccdY + 40, 0), new Point3d(kccdX, kccdY + double.Parse(linesList[2]) + 40, 0));
                        modelSpace.AppendEntity(cdLine);
                        tr.AddNewlyCreatedDBObject(cdLine, true);

                        tr.Commit();
                    }
                    // stt text
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        DBText stt = new DBText();
                        stt.TextString = "G" + count.ToString();
                        stt.Height = TSIZE;
                        stt.Rotation = 0;
                        stt.WidthFactor = 1;
                        if (count % 2 == 0) {
                            stt.Position = new Point3d(kccdX - 3.5, kccdY - 30, 0);
                        }
                        else {
                            stt.Position = new Point3d(kccdX - 3.5, kccdY - 20, 0);
                        }
                        modelSpace.AppendEntity(stt);
                        tr.AddNewlyCreatedDBObject(stt, true);

                        ed.WriteMessage("created stt \n");
                        tr.Commit();
                    }
                    if (count == 1) {
                        count++;
                        continue;
                    }
                    // kcl text
                    using (Transaction tr = db.TransactionManager.StartTransaction()) {
                        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        DBText kcl = new DBText();
                        kclX = kccdX - float.Parse(linesList[1]) / 2;

                        kcl.TextString = linesList[1];
                        kcl.Height = TSIZE;
                        kcl.Rotation = 1.571;
                        kcl.WidthFactor = 1;
                        kcl.Position = new Point3d(kclX + 2, kclY - 4, 0);
                        modelSpace.AppendEntity(kcl);
                        tr.AddNewlyCreatedDBObject(kcl, true);
                        ed.WriteMessage("created kcl \n");
                        tr.Commit();
                    }

                    count++;
                }

                using (Transaction tr = db.TransactionManager.StartTransaction()) {
                    BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Line upperLine = new Line(new Point3d(100, kccdY + 7.5, 0), new Point3d(kccdX, kccdY + 7.5, 0));
                    modelSpace.AppendEntity(upperLine);
                    tr.AddNewlyCreatedDBObject(upperLine, true);
                    tr.Commit();
                }

                using (Transaction tr = db.TransactionManager.StartTransaction()) {
                    BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Line lowerLine = new Line(new Point3d(100, kccdY + 17.5, 0), new Point3d(kccdX, kccdY + 17.5, 0));
                    modelSpace.AppendEntity(lowerLine);
                    tr.AddNewlyCreatedDBObject(lowerLine, true);
                    tr.Commit();
                }
            }
            catch {
                ed.WriteMessage("\nFailed to open file");
                return;
            }
        }

        [CommandMethod("NTMWritePolyline")]
        public void cmdWritePolyline() {
            string saveFilePath = func.OpenFile();

            try {
                StreamReader reader = new StreamReader(saveFilePath);
                using (Transaction trans = db.TransactionManager.StartTransaction()) {
                    BlockTable blockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = trans.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Polyline pll = new Polyline();

                    int count = 0;
                    for (String line = reader.ReadLine(); line != null; line = reader.ReadLine()) {
                        String[] XYZ = line.Split(',');
                        pll.AddVertexAt(count, new Point2d(double.Parse(XYZ[0]), double.Parse(XYZ[1])), 0, 0, 0);
                        count++;
                    }
                    modelSpace.AppendEntity(pll);
                    trans.AddNewlyCreatedDBObject(pll, true);

                    trans.Commit();

                    ed.WriteMessage("\nPolyline created successfully.");
                }
            }
            catch {
                ed.WriteMessage("\nFailed to perform");
                return;
            }
        }

        #endregion
    }
}
