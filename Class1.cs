using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace AutoCADCommands
{
  public class MyCommands
  {
    [CommandMethod("KEYEDPLAN")]
    public static void KEYEDPLAN()
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      Entity imageEntity = null;  // This will store RasterImage, Image, or OLE2Frame

      PromptSelectionResult selectionResult = ed.GetSelection();
      if (selectionResult.Status == PromptStatus.OK)
      {
        SelectionSet selectionSet = selectionResult.Value;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
          // First, let's find the RasterImage, Image, or OLE object and get its extents
          foreach (SelectedObject selObj in selectionSet)
          {
            Entity ent = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
            if (ent is RasterImage || ent is Image || ent is Ole2Frame)
            {
              imageEntity = ent;
              break;  // Once found, break out of the loop
            }
          }

          // If imageEntity is null, it means no appropriate image entity was found in the selection
          if (imageEntity == null)
          {
            ed.WriteMessage("\nNote: No appropriate image entity was found in the selection. Leaders for 'AREA OF WORK' will not be created.");
          }

          // Now, let's handle the other entities
          foreach (SelectedObject selObj in selectionSet)
          {
            Entity ent = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
            if (ent != null)
            {
              string objectType = ent.GetType().Name;
              string handle = ent.Handle.ToString();
              ed.WriteMessage($"\nSelected Object: Handle = {handle}, Type = {objectType}");

              if (ent is DBText || ent is MText)
              {
                WipeoutAroundText(selObj.ObjectId);
                if (imageEntity != null)  // Only proceed if imageEntity is not null
                {
                  if (ent is DBText dbTextEnt && dbTextEnt.TextString == "AREA OF WORK")
                  {
                    CreateLeaderFromTextToPoint(dbTextEnt, trans, imageEntity.GeometricExtents);
                  }
                  else if (ent is MText mTextEnt && mTextEnt.Contents == "AREA OF WORK")
                  {
                    CreateLeaderFromTextToPoint(mTextEnt, trans, imageEntity.GeometricExtents);
                  }
                }
              }
              else if (ent is Autodesk.AutoCAD.DatabaseServices.Polyline)
              {
                // Your existing code to handle polylines...
                HatchSelectedPolyline(selObj.ObjectId);
              }
              else if (ent is RasterImage || ent is Image || ent is Ole2Frame)  // Check image entity type again to handle other operations on it
              {
                Extents3d extents;
                Point3d endPoint;

                if (ent is RasterImage rasterImg)
                {
                  extents = rasterImg.GeometricExtents;
                }
                else if (ent is Image image)
                {
                  extents = image.GeometricExtents;
                }
                else if (ent is Ole2Frame oleFrame)
                {
                  extents = oleFrame.GeometricExtents;
                }
                else
                {
                  // If none match, continue to the next iteration
                  continue;
                }

                endPoint = new Point3d(extents.MinPoint.X, extents.MinPoint.Y, 0);  // Bottom left corner

                CreateEntitiesAtEndPoint(trans, extents, endPoint, "KEYED PLAN", "SCALE: NONE");
              }
            }
          }
          trans.Commit();
        }
      }
      else
      {
        ed.WriteMessage("\nNo objects were selected.");
      }
    }

    [CommandMethod("CREATEBLOCK")]
    public void CREATEBLOCK()
    {
      var (doc, db, _) = MyCommands.GetGlobals();

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        BlockTableRecord existingBtr = null;
        ObjectId existingBtrId = ObjectId.Null;

        // Check if block already exists
        if (bt.Has("CIRCLEI"))
        {
          existingBtrId = bt["CIRCLEI"];

          if (existingBtrId != ObjectId.Null)
          {
            existingBtr = (BlockTableRecord)tr.GetObject(existingBtrId, OpenMode.ForWrite);

            if (existingBtr != null && existingBtr.Name == "CIRCLEI")
            {
              doc.Editor.WriteMessage("\nBlock 'CIRCLEI' already exists and matches the new block. Exiting the function.");
              return; // Exit the function if existing block matches the new block
            }
          }
        }

        // Delete existing block and its contents
        if (existingBtr != null)
        {
          foreach (ObjectId id in existingBtr.GetBlockReferenceIds(true, true))
          {
            DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
            obj.Erase(true);
          }

          existingBtr.Erase(true);

          doc.Editor.WriteMessage("\nExisting block 'CIRCLEI' and its contents have been deleted.");
        }

        BlockTableRecord btr = new BlockTableRecord();
        btr.Name = "CIRCLEI";

        bt.UpgradeOpen();
        ObjectId btrId = bt.Add(btr);
        tr.AddNewlyCreatedDBObject(btr, true);

        // Create a circle centered at 0,0 with radius 2.0
        Circle circle = new Circle(new Point3d(0, 0, 0), new Vector3d(0, 0, 1), 0.09);
        circle.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2); // Set circle color to yellow

        btr.AppendEntity(circle);
        tr.AddNewlyCreatedDBObject(circle, true);

        // Create a text entity
        DBText text = new DBText();
        text.Position = new Point3d(-0.042, -0.045, 0); // centered at the origin
        text.Height = 0.09; // Set the text height
        text.TextString = "1";
        text.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2); // Set text color to yellow

        // Check if the text style "ROMANS" exists
        TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
        if (textStyleTable.Has("ROMANS"))
        {
          text.TextStyleId = textStyleTable["ROMANS"]; // apply the "ROMANS" text style to the text entity
        }

        // Check if the layer "E-TEXT" exists
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
        if (lt.Has("E-TEXT"))
        {
          circle.Layer = "E-TEXT"; // Set the layer of the circle to "E-TEXT"
          text.Layer = "E-TEXT"; // Set the layer of the text to "E-TEXT"
        }

        btr.AppendEntity(text);
        tr.AddNewlyCreatedDBObject(text, true);

        tr.Commit();
      }
    }

    [CommandMethod("KEEPBREAKERS")]
    public void KEEPBREAKERS()
    {
      var (doc, db, ed) = MyCommands.GetGlobals();

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        if (!bt.Has("CIRCLEI"))
        {
          PromptKeywordOptions pko = new PromptKeywordOptions("\nThe block 'CIRCLEI' does not exist. Do you want to create it? [Yes/No] ", "Yes No");
          pko.AllowNone = true;
          PromptResult pr = ed.GetKeywords(pko);
          String prompt = pr.StringResult.ToLower();
          if (prompt == "no" || prompt == "n")
            return;
          else if (prompt == "yes" || prompt == "y")
            CREATEBLOCK(); // Assuming this function exists and will create the required block
        }

        var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        var point1 = ed.GetPoint("\nSelect first point: ").Value;
        var point2 = ed.GetPoint("\nSelect second point: ").Value;

        // Swap the points if the y-coordinate of the first point is lower than that of the second point
        if (point1.Y < point2.Y)
        {
          Point3d tempPoint = point1;
          point1 = point2;
          point2 = tempPoint;
        }

        var point3 = ed.GetPoint("\nSelect third point: ").Value;

        var direction = point3.X > point1.X ? 1 : -1;
        var dist = (point1 - point2).Length;

        var line1Start = new Point3d(point1.X + direction * 0.05, point1.Y, 0);
        var line1End = new Point3d(line1Start.X + direction * 0.2, line1Start.Y, 0);
        var line2Start = new Point3d(line1Start.X, point2.Y, 0);
        var line2End = new Point3d(line1End.X, line2Start.Y, 0);

        string layerName = CreateOrGetLayer("E-TEXT", db, tr);

        var line1 = new Line(line1Start, line1End) { Layer = layerName };
        var line2 = new Line(line2Start, line2End) { Layer = layerName };

        var mid1 = new Point3d((line1Start.X + line1End.X) / 2, (line1Start.Y + line1End.Y) / 2, 0);
        var mid2 = new Point3d((line2Start.X + line2End.X) / 2, (line2Start.Y + line2End.Y) / 2, 0);
        var mid3 = new Point3d((mid1.X + mid2.X) / 2, (mid1.Y + mid2.Y) / 2, 0);

        var circleTop = new Point3d(mid3.X, mid3.Y + 0.09, 0);
        var circleBottom = new Point3d(mid3.X, mid3.Y - 0.09, 0);

        var line3 = new Line(mid1, circleTop) { Layer = layerName };
        var line4 = new Line(mid2, circleBottom) { Layer = layerName };

        if (dist > 0.3)
        {
          btr.AppendEntity(line1);
          btr.AppendEntity(line2);
          tr.AddNewlyCreatedDBObject(line1, true);
          tr.AddNewlyCreatedDBObject(line2, true);
        }

        var blkRef = new BlockReference(mid3, bt["CIRCLEI"]) { Layer = layerName };
        btr.AppendEntity(blkRef);
        tr.AddNewlyCreatedDBObject(blkRef, true);

        if (dist > 0.3)
        {
          btr.AppendEntity(line3);
          btr.AppendEntity(line4);
          tr.AddNewlyCreatedDBObject(line3, true);
          tr.AddNewlyCreatedDBObject(line4, true);
        }

        tr.Commit();
      }
    }

    [CommandMethod("T24")]
    public void T24()
    {
      // Get the current database and start a transaction
      Database acCurDb;
      acCurDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
      Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        // Get the user to select a PNG file
        OpenFileDialog ofd = new OpenFileDialog();
        ofd.Filter = "PNG Files (*.png)|*.png";
        if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK)
          return;

        string strFileName = ofd.FileName;

        // Determine the parent folder of the selected file
        string parentFolder = Path.GetDirectoryName(strFileName);

        // Fetch all relevant files in the folder
        string[] allFiles = Directory.GetFiles(parentFolder, "*.png")
            .OrderBy(f =>
            {
              var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(f);
              if (fileNameWithoutExtension.Contains("Page"))
              {
                // If the filename contains "Page", we extract the number after it.
                var lastPart = fileNameWithoutExtension.Split(' ').Last();
                return int.Parse(lastPart);
              }
              else
              {
                // If the filename does not contain "Page", we extract the last number after the hyphen.
                var lastPart = fileNameWithoutExtension.Split('-').Last();
                return int.Parse(lastPart);
              }
            }).ToArray();

        // Variable to track current image position
        int currentRow = 0;
        int currentColumn = 0;
        Point3d selectedPoint = Point3d.Origin;
        Vector3d width = new Vector3d(0, 0, 0);
        Vector3d height = new Vector3d(0, 0, 0);

        foreach (string file in allFiles)
        {
          string imageName = Path.GetFileNameWithoutExtension(file);

          RasterImageDef acRasterDef;
          bool bRasterDefCreated = false;
          ObjectId acImgDefId;

          // Get the image dictionary
          ObjectId acImgDctID = RasterImageDef.GetImageDictionary(acCurDb);

          // Check to see if the dictionary does not exist, it not then create it
          if (acImgDctID.IsNull)
          {
            acImgDctID = RasterImageDef.CreateImageDictionary(acCurDb);
          }

          // Open the image dictionary
          DBDictionary acImgDict = acTrans.GetObject(acImgDctID, OpenMode.ForRead) as DBDictionary;

          // Check to see if the image definition already exists
          if (acImgDict.Contains(imageName))
          {
            acImgDefId = acImgDict.GetAt(imageName);

            acRasterDef = acTrans.GetObject(acImgDefId, OpenMode.ForWrite) as RasterImageDef;
          }
          else
          {
            // Create a raster image definition
            RasterImageDef acRasterDefNew = new RasterImageDef();

            // Set the source for the image file
            acRasterDefNew.SourceFileName = file;

            // Load the image into memory
            acRasterDefNew.Load();

            // Add the image definition to the dictionary
            acTrans.GetObject(acImgDctID, OpenMode.ForWrite);
            acImgDefId = acImgDict.SetAt(imageName, acRasterDefNew);

            acTrans.AddNewlyCreatedDBObject(acRasterDefNew, true);

            acRasterDef = acRasterDefNew;

            bRasterDefCreated = true;
          }

          // Open the Block table for read
          BlockTable acBlkTbl;
          acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

          // Open the Block table record Paper space for write
          BlockTableRecord acBlkTblRec;
          acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

          // Create the new image and assign it the image definition
          using (RasterImage acRaster = new RasterImage())
          {
            acRaster.ImageDefId = acImgDefId;

            // Define the width and height of the image
            if (selectedPoint == Point3d.Origin)
            {
              // Check to see if the measurement is set to English (Imperial) or Metric units
              if (acCurDb.Measurement == MeasurementValue.English)
              {
                width = new Vector3d((acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth * 0.8) / 25.4, 0, 0);
                height = new Vector3d(0, (acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight * 0.8) / 25.4, 0);
              }
              else
              {
                width = new Vector3d(acRasterDef.ResolutionMMPerPixel.X * acRaster.ImageWidth * 0.8, 0, 0);
                height = new Vector3d(0, acRasterDef.ResolutionMMPerPixel.Y * acRaster.ImageHeight * 0.8, 0);
              }
            }

            // Prompt the user to select a point
            // Only for the first image
            if (selectedPoint == Point3d.Origin)
            {
              PromptPointResult ppr = ed.GetPoint("\nSelect a point to insert images:");
              if (ppr.Status != PromptStatus.OK)
                return;
              selectedPoint = ppr.Value;
            }

            // Calculate the new position based on the row and column
            Point3d currentPos = new Point3d(
                selectedPoint.X - (currentColumn * width.X) - width.X, // Subtract width.X to shift the starting point to the top right corner of the first image
                selectedPoint.Y - (currentRow * height.Y) - height.Y, // Subtract height.Y to shift the starting point to the top right corner of the first image
                0
            );

            // Define and assign a coordinate system for the image's orientation
            CoordinateSystem3d coordinateSystem = new CoordinateSystem3d(currentPos, width, height);
            acRaster.Orientation = coordinateSystem;

            // Set the rotation angle for the image
            acRaster.Rotation = 0;

            // Add the new object to the block table record and the transaction
            acBlkTblRec.AppendEntity(acRaster);
            acTrans.AddNewlyCreatedDBObject(acRaster, true);

            // Connect the raster definition and image together so the definition
            // does not appear as "unreferenced" in the External References palette.
            RasterImage.EnableReactors(true);
            acRaster.AssociateRasterDef(acRasterDef);

            if (bRasterDefCreated)
            {
              acRasterDef.Dispose();
            }
          }

          // Move to the next column
          currentColumn++;

          // Start a new row every 3 images
          if (currentColumn % 3 == 0)
          {
            currentRow++;
            currentColumn = 0;
          }
        }

        // Save the new object to the database
        acTrans.Commit();
      }
    }

    [CommandMethod("SUMTEXT")]
    public void SUMTEXT()
    {
      var (doc, db, ed) = MyCommands.GetGlobals();

      PromptSelectionResult selection = ed.SelectImplied();

      if (selection.Status != PromptStatus.OK)
      {
        PromptSelectionOptions opts = new PromptSelectionOptions();
        opts.MessageForAdding = "Select text objects to sum: ";
        opts.AllowDuplicates = false;
        opts.RejectObjectsOnLockedLayers = true;

        selection = ed.GetSelection(opts);
        if (selection.Status != PromptStatus.OK) return;
      }

      double sum = 0.0;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        foreach (SelectedObject so in selection.Value)
        {
          DBText text = tr.GetObject(so.ObjectId, OpenMode.ForRead) as DBText;
          MText mtext = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;

          if (text != null)
          {
            double value;
            string textString = text.TextString.Replace("sq ft", "").Trim();
            if (Double.TryParse(textString, out value)) sum += value;
          }
          else if (mtext != null)
          {
            double value;
            string mTextContents = mtext.Contents.Replace("sq ft", "").Trim();
            if (Double.TryParse(mTextContents, out value)) sum += value;
          }
        }
        ed.WriteMessage($"\nThe sum of selected text objects is: {sum} sq ft");
        tr.Commit();
      }
    }

    [CommandMethod("AREACALCULATOR")]
    public void AREACALCULATOR()
    {
      var (doc, _, ed) = MyCommands.GetGlobals();

      PromptSelectionOptions opts = new PromptSelectionOptions();
      opts.MessageForAdding = "Select polylines or rectangles: ";
      opts.AllowDuplicates = false;
      opts.RejectObjectsOnLockedLayers = true;

      PromptSelectionResult selection = ed.GetSelection(opts);
      if (selection.Status != PromptStatus.OK) return;

      using (Transaction tr = doc.TransactionManager.StartTransaction())
      {
        foreach (ObjectId objId in selection.Value.GetObjectIds())
        {
          var obj = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
          if (obj == null)
          {
            ed.WriteMessage("\nSelected object is not a valid entity.");
            continue;
          }

          Autodesk.AutoCAD.DatabaseServices.Polyline polyline = obj as Autodesk.AutoCAD.DatabaseServices.Polyline;
          if (polyline != null)
          {
            double area = polyline.Area;
            area /= 144; // Converting from square inches to square feet
            ed.WriteMessage("\nThe area of the selected polyline is: " + area + " sq ft");

            // Get the bounding box of the polyline
            Extents3d bounds = (Extents3d)polyline.Bounds;

            // Calculate the center of the bounding box
            Point3d center = new Point3d((bounds.MinPoint.X + bounds.MaxPoint.X) / 2, (bounds.MinPoint.Y + bounds.MaxPoint.Y) / 2, 0);

            // Check if the center of the bounding box lies within the polyline. If not, use the first vertex.
            if (!IsPointInside(polyline, center))
            {
              center = polyline.GetPoint3dAt(0);
            }

            DBText text = new DBText
            {
              Height = 9,
              TextString = Math.Ceiling(area) + " sq ft",
              Rotation = 0,
              HorizontalMode = TextHorizontalMode.TextCenter,
              VerticalMode = TextVerticalMode.TextVerticalMid,
              Layer = "0"
            };

            text.Position = center;
            text.AlignmentPoint = center;

            var currentSpace = (BlockTableRecord)tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
            currentSpace.AppendEntity(text);
            tr.AddNewlyCreatedDBObject(text, true);
          }
          else
          {
            ed.WriteMessage("\nSelected object is not a polyline.");
            continue;
          }
        }

        tr.Commit();
      }
    }

    [CommandMethod("GETTEXTATTRIBUTES")]
    public void GETTEXTATTRIBUTES()
    {
      var (doc, db, ed) = MyCommands.GetGlobals();

      var textId = SelectTextObject();
      if (textId.IsNull)
      {
        ed.WriteMessage("\nNo text object selected.");
        return;
      }

      var textObject = GetTextObject(textId);
      if (textObject == null)
      {
        ed.WriteMessage("\nFailed to get text object.");
        return;
      }

      var coordinate = GetCoordinate();
      if (coordinate == null)
      {
        ed.WriteMessage("\nInvalid coordinate selected.");
        return;
      }

      var startPoint = new Point3d(textObject.Position.X - coordinate.X, textObject.Position.Y - coordinate.Y, 0);

      string startXStr = startPoint.X == 0 ? "" : (startPoint.X > 0 ? $" + {startPoint.X}" : $" - {-startPoint.X}");
      string startYStr = startPoint.Y == 0 ? "" : (startPoint.Y > 0 ? $" + {startPoint.Y}" : $" - {-startPoint.Y}");

      var formattedText = $"CreateAndPositionText(tr, \"{textObject.TextString}\", \"{textObject.TextStyleName}\", {textObject.Height}, {textObject.WidthFactor}, {textObject.Color.ColorIndex}, \"{textObject.Layer}\", new Point3d(endPoint.X{startXStr}, endPoint.Y{startYStr}, 0));";

      var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
      var filePath = Path.Combine(desktopPath, "TextAttributes.txt");

      SaveTextToFile(formattedText, filePath);
      ed.WriteMessage($"\nText attributes saved to file: {filePath}");
    }

    [CommandMethod("GETLINEATTRIBUTES")]
    public void GETLINEATTRIBUTES()
    {
      var (doc, db, ed) = MyCommands.GetGlobals();

      PromptEntityOptions linePromptOptions = new PromptEntityOptions("\nSelect a line: ");
      linePromptOptions.SetRejectMessage("\nSelected object is not a line.");
      linePromptOptions.AddAllowedClass(typeof(Line), true);

      PromptEntityResult lineResult = ed.GetEntity(linePromptOptions);
      if (lineResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nNo line selected.");
        return;
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        Line line = tr.GetObject(lineResult.ObjectId, OpenMode.ForRead) as Line;
        if (line == null)
        {
          ed.WriteMessage("\nSelected object is not a line.");
          return;
        }

        PromptPointOptions startPointOptions = new PromptPointOptions("\nSelect the reference point: ");
        PromptPointResult startPointResult = ed.GetPoint(startPointOptions);
        if (startPointResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo reference point selected.");
          return;
        }

        Point3d startPoint = startPointResult.Value;
        Vector3d vector = line.EndPoint - line.StartPoint;

        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var filePath = Path.Combine(desktopPath, "LineAttributes.txt");

        SaveLineAttributesToFile(line, startPoint, vector, filePath);

        ed.WriteMessage($"\nLine attributes saved to file: {filePath}");
      }
    }

    [CommandMethod("VP")]
    public void CREATEVIEWPORTFROMREGION()
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // Prompt for sheet name
      PromptResult sheetNameResult = ed.GetString("\nPlease enter the sheet name: ");
      if (sheetNameResult.Status != PromptStatus.OK)
        return;

      string inputSheetName = sheetNameResult.StringResult;
      string matchedLayoutName = null;

      // Check if the input directly matches a layout or needs to be prefixed by "E-"
      string expectedSheetName = inputSheetName.StartsWith("E-") ? inputSheetName : "E-" + inputSheetName;

      // Prompt for the first corner of the rectangle
      PromptPointOptions pointOpts1 = new PromptPointOptions("Please select the first corner of the region in modelspace:");
      PromptPointResult pointResult1 = ed.GetPoint(pointOpts1);
      if (pointResult1.Status != PromptStatus.OK)
        return;

      // Prompt for the opposite corner of the rectangle
      PromptPointOptions pointOpts2 = new PromptPointOptions("Please select the opposite corner of the region in modelspace:");
      pointOpts2.BasePoint = pointResult1.Value;
      pointOpts2.UseBasePoint = true;
      PromptPointResult pointResult2 = ed.GetPoint(pointOpts2);
      if (pointResult2.Status != PromptStatus.OK)
        return;

      var correctedPoints = GetCorrectedPoints(pointResult1.Value, pointResult2.Value);
      Extents3d rectExtents = new Extents3d(correctedPoints.Min, correctedPoints.Max);
      double rectWidth = rectExtents.MaxPoint.X - rectExtents.MinPoint.X;
      double rectHeight = rectExtents.MaxPoint.Y - rectExtents.MinPoint.Y;

      ed.WriteMessage($"Checking width {rectWidth}, and height {rectHeight}");

      Dictionary<double, double> scales = new Dictionary<double, double>
    {
        { 0.25, 48.0 },
        { 3.0 / 16.0, 64.0 },
        { 1.0 / 8.0, 96.0 },
        { 3.0 / 32.0, 128.0 },
        { 1.0 / 16.0, 192.0 }
    };

      double scaleToFit = 0.0;
      double viewportWidth = 0.0;
      double viewportHeight = 0.0;

      foreach (var scaleEntry in scales.OrderByDescending(e => e.Key))
      {
        viewportWidth = rectWidth / scaleEntry.Value;
        viewportHeight = rectHeight / scaleEntry.Value;

        if (viewportWidth <= 30 && viewportHeight <= 22)
        {
          scaleToFit = scaleEntry.Key;
          break;
        }

        ed.WriteMessage($"\nChecking scale {scaleEntry.Key}: viewportWidth = {viewportWidth}, viewportHeight = {viewportHeight}");
      }

      if (scaleToFit == 0.0)
      {
        ed.WriteMessage("Couldn't fit the rectangle in the specified scales");
        return;
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // Get the layout dictionary
        DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

        foreach (var layoutEntry in layoutDict)
        {
          string layoutName = layoutEntry.Key;
          if (layoutName.StartsWith("E-") && layoutName.Contains(inputSheetName))
          {
            matchedLayoutName = layoutName;
            break;
          }
        }

        if (string.IsNullOrEmpty(matchedLayoutName))
        {
          ed.WriteMessage($"No matching layout found for '{inputSheetName}'.");
          return;
        }

        if (!layoutDict.Contains(matchedLayoutName))
        {
          ed.WriteMessage($"Sheet (Layout) named '{matchedLayoutName}' not found in the drawing.");
          return;
        }
        ObjectId layoutId = layoutDict.GetAt(matchedLayoutName);
        Layout layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;

        BlockTable blockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord paperSpace = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;

        LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        if (!layerTable.Has("DEFPOINTS"))
        {
          LayerTableRecord layerRecord = new LayerTableRecord
          {
            Name = "DEFPOINTS",
            Color = Color.FromColorIndex(ColorMethod.ByAci, 7) // White color
          };

          layerTable.UpgradeOpen(); // Switch to write mode
          ObjectId layerId = layerTable.Add(layerRecord);
          tr.AddNewlyCreatedDBObject(layerRecord, true);
        }

        Point2d modelSpaceCenter = new Point2d((rectExtents.MinPoint.X + rectWidth / 2), (rectExtents.MinPoint.Y + rectHeight / 2));

        Viewport viewport = new Viewport();

        // This is the placement of the viewport on the PAPER SPACE (typically the center of your paper space or wherever you want the viewport to appear)
        viewport.CenterPoint = new Point3d(32.7086 - viewportWidth / 2.0, 23.3844 - viewportHeight / 2.0, 0.0);
        viewport.Width = viewportWidth;
        viewport.Height = viewportHeight;
        viewport.CustomScale = scaleToFit / 12;
        viewport.Layer = "DEFPOINTS";

        // This is the center of the view in MODEL SPACE (the actual content you want to show inside the viewport)
        ed.WriteMessage($"\nModelSpaceCenterX: {modelSpaceCenter.X}, ModelSpaceCenterY: {modelSpaceCenter.Y}");
        viewport.ViewTarget = new Point3d(modelSpaceCenter.X, modelSpaceCenter.Y, 0.0);
        viewport.ViewDirection = new Vector3d(0, 0, 1);

        ed.WriteMessage($"\nSet viewport scale to {viewport.CustomScale}");

        paperSpace.AppendEntity(viewport);
        tr.AddNewlyCreatedDBObject(viewport, true);

        db.TileMode = false; // Set to Paper Space

        // Set the current layout to the one you are working on
        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CTAB", layout.LayoutName);

        viewport.On = true; // Now turn the viewport on

        // Prompt user for the type of viewport
        PromptResult viewportTypeResult = ed.GetString("\nPlease enter the type of viewport (e.g., lighting, power, roof): ");
        if (viewportTypeResult.Status == PromptStatus.OK)
        {
          string viewportTypeUpperCase = viewportTypeResult.StringResult.ToUpper();
          string finalViewportText = "ELECTRICAL " + viewportTypeUpperCase + " PLAN";

          // Getting scale in string format
          string scaleStr = ScaleToFraction(12 * viewport.CustomScale);
          string text2 = $"SCALE: {scaleStr}\" = 1'";

          // Create extents using the viewport properties
          Point3d minPoint = new Point3d(viewport.CenterPoint.X - viewport.Width / 2.0, viewport.CenterPoint.Y - viewport.Height / 2.0, 0.0);
          Point3d maxPoint = new Point3d(viewport.CenterPoint.X + viewport.Width / 2.0, viewport.CenterPoint.Y + viewport.Height / 2.0, 0.0);
          Extents3d viewportExtents = new Extents3d(minPoint, maxPoint);

          // Use the function to create the title
          CreateEntitiesAtEndPoint(tr, viewportExtents, minPoint, finalViewportText, text2);
        }

        tr.Commit();
      }

      ed.Regen();
    }

    public static void HatchSelectedPolyline(ObjectId? polyId = null)
    {
      Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database acCurDb = acDoc.Database;

      if (!polyId.HasValue)
      {
        // Prompt the user to select a polyline
        PromptEntityOptions opts = new PromptEntityOptions("\nSelect a polyline: ");
        opts.SetRejectMessage("\nThat is not a polyline. Please select a polyline.");
        opts.AddAllowedClass(typeof(Polyline), true);
        PromptEntityResult per = acDoc.Editor.GetEntity(opts);

        if (per.Status != PromptStatus.OK) return;

        polyId = per.ObjectId;
      }

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        BlockTable acBt = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

        // Check if the user is in paper space or model space and then use that space for the operations
        BlockTableRecord acBtr;
        if (acCurDb.TileMode)
        {
          if (acCurDb.PaperSpaceVportId == acDoc.Editor.CurrentViewportObjectId)
          {
            acBtr = acTrans.GetObject(acBt[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
          }
          else
          {
            acBtr = acTrans.GetObject(acBt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
          }
        }
        else
        {
          acBtr = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        }

        Polyline acPoly = acTrans.GetObject(polyId.Value, OpenMode.ForWrite) as Polyline;

        // If the polyline is not closed, close it
        if (!acPoly.Closed)
        {
          acPoly.Closed = true;
        }

        ObjectIdCollection oidCol = new ObjectIdCollection();
        oidCol.Add(polyId.Value);

        using (Hatch acHatch = new Hatch())
        {
          acBtr.AppendEntity(acHatch);
          acTrans.AddNewlyCreatedDBObject(acHatch, true);

          acHatch.SetHatchPattern(HatchPatternType.PreDefined, "ANSI31");
          acHatch.Associative = true;
          acHatch.AppendLoop(HatchLoopTypes.External, oidCol);
          acHatch.EvaluateHatch(true);
        }

        acTrans.Commit();
      }
    }

    public static (Document doc, Database db, Editor ed) GetGlobals()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var db = doc.Database;
      var ed = doc.Editor;

      return (doc, db, ed);
    }

    public static void WipeoutAroundText(ObjectId? textObjectId = null)
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      if (!textObjectId.HasValue)
      {
        // Prompt the user to select a text object
        PromptEntityOptions opts = new PromptEntityOptions("\nSelect a text object: ");
        opts.SetRejectMessage("\nOnly text objects are allowed.");
        opts.AddAllowedClass(typeof(DBText), true);
        opts.AddAllowedClass(typeof(MText), true);
        PromptEntityResult per = ed.GetEntity(opts);

        if (per.Status != PromptStatus.OK) return;

        textObjectId = per.ObjectId;
      }

      double margin = 0.05;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        Entity ent = (Entity)tr.GetObject(textObjectId.Value, OpenMode.ForRead);
        double rotation = 0;
        Point3d basePoint = Point3d.Origin;  // default to origin

        if (ent is DBText text)
        {
          rotation = text.Rotation;
          basePoint = text.Position;
          text.UpgradeOpen();
          text.Rotation = 0;
        }
        else if (ent is MText mtext)
        {
          rotation = mtext.Rotation;
          basePoint = mtext.Location;
          mtext.UpgradeOpen();
          mtext.Rotation = 0;
        }

        Extents3d extents = ent.GeometricExtents; // Recalculate extents after rotation

        Point3d minPoint = extents.MinPoint;
        Point3d maxPoint = extents.MaxPoint;

        // Add margin
        Point3d pt1 = new Point3d(minPoint.X - margin, minPoint.Y - margin, 0);
        Point3d pt2 = new Point3d(maxPoint.X + margin, maxPoint.Y + margin, 0);
        Point3d pt3 = new Point3d(pt1.X, pt2.Y, 0);
        Point3d pt4 = new Point3d(pt2.X, pt1.Y, 0);

        Point2dCollection pts = new Point2dCollection
        {
            new Point2d(pt1.X, pt1.Y),
            new Point2d(pt4.X, pt4.Y),
            new Point2d(pt2.X, pt2.Y),
            new Point2d(pt3.X, pt3.Y),
            new Point2d(pt1.X, pt1.Y)  // Close the loop
        };

        Wipeout wo = new Wipeout();
        wo.SetDatabaseDefaults(db);
        wo.SetFrom(pts, new Vector3d(0.0, 0.0, 1.0));

        BlockTableRecord currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        // Rotate wipeout and text back to their original rotation using common base point
        wo.TransformBy(Matrix3d.Rotation(rotation, new Vector3d(0, 0, 1), basePoint));
        if (ent is DBText)
        {
          ((DBText)ent).Rotation = rotation;
        }
        else if (ent is MText)
        {
          ((MText)ent).Rotation = rotation;
        }

        currentSpace.AppendEntity(wo);
        tr.AddNewlyCreatedDBObject(wo, true);

        // Send the text object to the front
        DrawOrderTable dot = (DrawOrderTable)tr.GetObject(currentSpace.DrawOrderTableId, OpenMode.ForWrite);
        ObjectIdCollection ids = new ObjectIdCollection { textObjectId.Value };
        dot.MoveToTop(ids);

        tr.Commit();
      }
    }

    public static bool IsPointInside(Polyline polyline, Point3d point)
    {
      int numIntersections = 0;
      for (int i = 0; i < polyline.NumberOfVertices; i++)
      {
        Point3d point1 = polyline.GetPoint3dAt(i);
        Point3d point2 = polyline.GetPoint3dAt((i + 1) % polyline.NumberOfVertices); // Get next point, or first point if we're at the end

        // Check if point is on an horizontal segment
        if (point1.Y == point2.Y && point1.Y == point.Y && point.X > Math.Min(point1.X, point2.X) && point.X < Math.Max(point1.X, point2.X))
        {
          return true;
        }

        if (point.Y > Math.Min(point1.Y, point2.Y) && point.Y <= Math.Max(point1.Y, point2.Y) && point.X <= Math.Max(point1.X, point2.X) && point1.Y != point2.Y)
        {
          double xinters = (point.Y - point1.Y) * (point2.X - point1.X) / (point2.Y - point1.Y) + point1.X;

          // Check if point is on the polygon boundary (other than horizontal)
          if (Math.Abs(point.X - xinters) < Double.Epsilon)
          {
            return true;
          }

          // Count intersections
          if (point.X < xinters)
          {
            numIntersections++;
          }
        }
      }
      // If the number of intersections is odd, the point is inside.
      return numIntersections % 2 != 0;
    }

    public static void CreateEntitiesAtEndPoint(Transaction trans, Extents3d extents, Point3d endPoint, string text1, string text2)
    {
      // First Text - "KEYED PLAN"
      CreateAndPositionText(trans, text1, "section title", 0.25, 0.85, 2, "E-TXT1",
          new Point3d(endPoint.X - 0.0217553592831337, endPoint.Y - 0.295573529244971, 0));

      // Polyline
      CreatePolyline(Color.FromColorIndex(ColorMethod.ByAci, 2), "E-TXT1",
          new Point2d[] { new Point2d(endPoint.X, endPoint.Y - 0.38), new Point2d(extents.MaxPoint.X, endPoint.Y - 0.38) }, 0.0625, 0.0625);

      // Second Text - "SCALE: NONE"
      CreateAndPositionText(trans, text2, "gmep", 0.1, 1.0, 2, "E-TXT1",
          new Point3d(endPoint.X, endPoint.Y - 0.57, 0));
    }

    private static string ScaleToFraction(double scale)
    {
      var knownScales = new Dictionary<double, string>
    {
        { 0.25, "1/4" },
        { 3.0 / 16.0, "3/16" },
        { 1.0 / 8.0, "1/8" },
        { 3.0 / 32.0, "3/32" },
        { 0.0625, "1/16" }
    };

      return knownScales.ContainsKey(scale) ? knownScales[scale] : scale.ToString();
    }

    private (Point3d Min, Point3d Max) GetCorrectedPoints(Point3d p1, Point3d p2)
    {
      Point3d minPoint = new Point3d(
          Math.Min(p1.X, p2.X),
          Math.Min(p1.Y, p2.Y),
          0
      );

      Point3d maxPoint = new Point3d(
          Math.Max(p1.X, p2.X),
          Math.Max(p1.Y, p2.Y),
          0
      );

      return (minPoint, maxPoint);
    }

    public string CreateOrGetLayer(string layerName, Database db, Transaction tr)
    {
      var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

      if (!lt.Has(layerName)) // check if layer exists
      {
        lt.UpgradeOpen(); // switch to write mode
        LayerTableRecord ltr = new LayerTableRecord();
        ltr.Name = layerName;
        lt.Add(ltr);
        tr.AddNewlyCreatedDBObject(ltr, true);
      }

      return layerName;
    }

    private void SaveTextToFile(string text, string filePath)
    {
      using (StreamWriter writer = new StreamWriter(filePath, true))
      {
        writer.WriteLine(text);
      }
    }

    private void SaveLineAttributesToFile(Line line, Point3d startPoint, Vector3d vector, string filePath)
    {
      using (StreamWriter writer = new StreamWriter(filePath, true))
      {
        double startX = line.StartPoint.X - startPoint.X;
        double startY = line.StartPoint.Y - startPoint.Y;
        double endX = line.EndPoint.X - startPoint.X;
        double endY = line.EndPoint.Y - startPoint.Y;

        string startXStr = startX == 0 ? "" : (startX > 0 ? $" + {startX}" : $" - {-startX}");
        string startYStr = startY == 0 ? "" : (startY > 0 ? $" + {startY}" : $" - {-startY}");
        string endXStr = endX == 0 ? "" : (endX > 0 ? $" + {endX}" : $" - {-endX}");
        string endYStr = endY == 0 ? "" : (endY > 0 ? $" + {endY}" : $" - {-endY}");

        writer.WriteLine($"CreateLine(tr, btr, endPoint.X{startXStr}, endPoint.Y{startYStr}, endPoint.X{endXStr}, endPoint.Y{endYStr}, \"{line.Layer}\");");
      }
    }

    private static ObjectId CreateText(string content, string style, TextHorizontalMode horizontalMode, TextVerticalMode verticalMode, double height, double widthFactor, Autodesk.AutoCAD.Colors.Color color, string layer)
    {
      var (doc, db, _) = MyCommands.GetGlobals();

      // Check if the layer exists
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

        if (!layerTable.Has(layer))
        {
          // Layer doesn't exist, create it
          var newLayer = new LayerTableRecord();
          newLayer.Name = layer;

          layerTable.UpgradeOpen();
          layerTable.Add(newLayer);
          tr.AddNewlyCreatedDBObject(newLayer, true);
        }

        tr.Commit();
      }

      using (var tr = doc.TransactionManager.StartTransaction())
      {
        var textStyleId = GetTextStyleId(style);
        var textStyle = (TextStyleTableRecord)tr.GetObject(textStyleId, OpenMode.ForRead);

        var text = new DBText
        {
          TextString = content,
          Height = height,
          WidthFactor = widthFactor,
          Color = color,
          Layer = layer,
          TextStyleId = textStyleId,
          HorizontalMode = horizontalMode,
          VerticalMode = verticalMode,
          Justify = AttachmentPoint.BaseLeft
        };

        var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        currentSpace.AppendEntity(text);
        tr.AddNewlyCreatedDBObject(text, true);

        tr.Commit();

        return text.ObjectId;
      }
    }

    private static void CreateAndPositionText(Transaction tr, string content, string style, double height, double widthFactor, int colorIndex, string layerName, Point3d position, TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft, TextVerticalMode verticalMode = TextVerticalMode.TextBase)
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, (short)colorIndex);
      var textId = CreateText(content, style, horizontalMode, verticalMode, height, widthFactor, color, layerName);
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
      text.Position = position;
    }

    private static ObjectId GetTextStyleId(string styleName)
    {
      var (doc, db, _) = MyCommands.GetGlobals();
      var textStyleTable = (TextStyleTable)db.TextStyleTableId.GetObject(OpenMode.ForRead);

      if (textStyleTable.Has(styleName))
      {
        return textStyleTable[styleName];
      }
      else
      {
        // Return the ObjectId of the "Standard" style
        return textStyleTable["Standard"];
      }
    }

    private ObjectId SelectTextObject()
    {
      var (doc, _, ed) = MyCommands.GetGlobals();

      var promptOptions = new PromptEntityOptions("\nSelect a text object: ");
      promptOptions.SetRejectMessage("Selected object is not a text object.");
      promptOptions.AddAllowedClass(typeof(DBText), exactMatch: true);

      var promptResult = ed.GetEntity(promptOptions);
      if (promptResult.Status == PromptStatus.OK)
        return promptResult.ObjectId;

      return ObjectId.Null;
    }

    private DBText GetTextObject(ObjectId objectId)
    {
      using (var tr = objectId.Database.TransactionManager.StartTransaction())
      {
        var textObject = tr.GetObject(objectId, OpenMode.ForRead) as DBText;
        if (textObject != null)
          return textObject;

        return null;
      }
    }

    private Point3d GetCoordinate()
    {
      var (doc, _, ed) = MyCommands.GetGlobals();

      var promptOptions = new PromptPointOptions("\nSelect a coordinate: ");
      var promptResult = ed.GetPoint(promptOptions);

      if (promptResult.Status == PromptStatus.OK)
        return promptResult.Value;

      return new Point3d(0, 0, 0);
    }

    public static void CreatePolyline(
    Color color,
    string layer,
    Point2d[] vertices,
    double startWidth,
    double endWidth)
    {
      Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database acCurDb = acDoc.Database;
      Editor ed = acDoc.Editor;

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        BlockTable acBlkTbl;
        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

        // Get the current space (either ModelSpace or the current layout)
        BlockTableRecord acBlkTblRec;
        acBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

        if (acBlkTblRec == null)
        {
          ed.WriteMessage("\nFailed to retrieve the current space block record.");
          return;
        }

        Polyline acPoly = new Polyline();
        for (int i = 0; i < vertices.Length; i++)
        {
          acPoly.AddVertexAt(i, vertices[i], 0, startWidth, endWidth);
        }

        acPoly.Color = color;
        acPoly.Layer = layer;

        acBlkTblRec.AppendEntity(acPoly);
        acTrans.AddNewlyCreatedDBObject(acPoly, true);

        acTrans.Commit();

        // Outputting details for debugging
        ed.WriteMessage($"\nPolyline created in layer: {layer} with color: {color.ColorName}. StartPoint: {vertices[0].ToString()} EndPoint: {vertices[vertices.Length - 1].ToString()}");
      }
    }

    private static void CreateLeaderFromTextToPoint(Entity textEnt, Transaction trans, Extents3d imageExtents)
    {
      Extents3d textExtents = textEnt.GeometricExtents;
      Point3d leftMid = new Point3d(textExtents.MinPoint.X, (textExtents.MinPoint.Y + textExtents.MaxPoint.Y) / 2, 0);
      Point3d rightMid = new Point3d(textExtents.MaxPoint.X, (textExtents.MinPoint.Y + textExtents.MaxPoint.Y) / 2, 0);

      // Find the nearest Polyline in the document
      Polyline closestPoly = FindClosestPolyline(textEnt.Database, imageExtents);
      if (closestPoly == null) return;

      // Determine the closest side to the polyline and create the leader accordingly
      Point3d closestPointOnPoly = closestPoly.GetClosestPointTo(leftMid, false);
      Point3d secondPoint;
      if (leftMid.DistanceTo(closestPointOnPoly) <= rightMid.DistanceTo(closestPointOnPoly))
      {
        // Left side is closer
        secondPoint = new Point3d(leftMid.X - 0.25, leftMid.Y, 0);
      }
      else
      {
        // Right side is closer
        secondPoint = new Point3d(rightMid.X + 0.25, rightMid.Y, 0);
        closestPointOnPoly = closestPoly.GetClosestPointTo(rightMid, false);
      }

      // Create the leader
      Leader acLdr = new Leader();
      acLdr.SetDatabaseDefaults();
      acLdr.AppendVertex(closestPointOnPoly);
      acLdr.AppendVertex(secondPoint);
      acLdr.AppendVertex((leftMid.DistanceTo(closestPointOnPoly) <= rightMid.DistanceTo(closestPointOnPoly)) ? leftMid : rightMid);

      // Set the leader's color to yellow
      acLdr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2); // Yellow

      BlockTableRecord acBlkTblRec = trans.GetObject(textEnt.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
      acBlkTblRec.AppendEntity(acLdr);
      trans.AddNewlyCreatedDBObject(acLdr, true);
    }

    private static Polyline FindClosestPolyline(Database db, Extents3d imageExtents)
    {
      Polyline closestPoly = null;
      double closestDist = double.MaxValue;

      using (Transaction trans = db.TransactionManager.StartTransaction())
      {
        BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
        foreach (ObjectId entId in btr)
        {
          Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;
          if (ent is Polyline)
          {
            Polyline poly = ent as Polyline;
            Point3d closestPoint = poly.GetClosestPointTo(imageExtents.MinPoint, false);
            double currentDist = closestPoint.DistanceTo(imageExtents.MinPoint);

            if (currentDist < closestDist && IsPointInsideExtents(closestPoint, imageExtents))
            {
              closestDist = currentDist;
              closestPoly = poly;
            }
          }
        }
      }

      return closestPoly;
    }

    private static bool IsPointInsideExtents(Point3d pt, Extents3d extents)
    {
      return pt.X >= extents.MinPoint.X && pt.X <= extents.MaxPoint.X &&
             pt.Y >= extents.MinPoint.Y && pt.Y <= extents.MaxPoint.Y &&
             pt.Z >= extents.MinPoint.Z && pt.Z <= extents.MaxPoint.Z;
    }
  }
}