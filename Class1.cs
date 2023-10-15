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
using Autodesk.AutoCAD.PlottingServices;
using System.Windows.Media.Media3D;

[assembly: CommandClass(typeof(AutoCADCommands.MyCommands))]
[assembly: CommandClass(typeof(AutoCADCommands.HotkeyManager))]

namespace AutoCADCommands
{
  public static class PolylineExtensions
  {
    public static bool IsPointInside(this Polyline polyline, Point3d point)
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
  }

  public class MyCommands
  {
    [CommandMethod("WO")]
    public void WO()
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // Prompt for first point
      PromptPointResult pr = ed.GetPoint("\nSpecify the first corner of the rectangle: ");
      if (pr.Status != PromptStatus.OK) return;
      Point3d pt1 = pr.Value;

      // Prompt for second point using the first point as a base point
      pr = ed.GetCorner("\nSpecify the opposite corner of the rectangle: ", pt1);
      if (pr.Status != PromptStatus.OK) return;
      Point3d pt2 = pr.Value;

      // Calculate the other two corners
      Point3d pt3 = new Point3d(pt1.X, pt2.Y, 0);
      Point3d pt4 = new Point3d(pt2.X, pt1.Y, 0);

      // Convert them to Point2dCollection
      Point2dCollection pts = new Point2dCollection
    {
        new Point2d(pt1.X, pt1.Y),
        new Point2d(pt4.X, pt4.Y),
        new Point2d(pt2.X, pt2.Y),
        new Point2d(pt3.X, pt3.Y),
        new Point2d(pt1.X, pt1.Y)  // Close the loop
    };

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);

        Wipeout wo = new Wipeout();
        wo.SetDatabaseDefaults(db);
        wo.SetFrom(pts, new Vector3d(0.0, 0.0, 1.0)); // Normal vector pointing up

        btr.AppendEntity(wo);
        tr.AddNewlyCreatedDBObject(wo, true);
        tr.Commit();
      }
    }

    [CommandMethod("WOT")]
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

    [CommandMethod("HatchPoly")]
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

    [CommandMethod("TYPEOFOBJECT")]
    public static void PrintSelectedObjectTypes()
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      PromptSelectionResult selectionResult = ed.GetSelection();
      if (selectionResult.Status == PromptStatus.OK)
      {
        SelectionSet selectionSet = selectionResult.Value;
        using (Transaction trans = db.TransactionManager.StartTransaction())
        {
          foreach (SelectedObject selObj in selectionSet)
          {
            if (selObj != null && selObj.ObjectId.IsValid)
            {
              Entity ent = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
              if (ent != null)
              {
                string objectType = ent.GetType().Name;
                string handle = ent.Handle.ToString();
                ed.WriteMessage($"\nSelected Object: Handle = {handle}, Type = {objectType}");
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
            if (!polyline.IsPointInside(center))
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

    [CommandMethod("IMPORTPANEL")]
    public void IMPORTPANEL()
    {
      var (doc, db, ed) = MyCommands.GetGlobals();

      List<Dictionary<string, object>> panelDataList = ImportExcelData();

      var spaceId = (db.TileMode == true) ? SymbolUtilityServices.GetBlockModelSpaceId(db) : SymbolUtilityServices.GetBlockPaperSpaceId(db);

      // Get the insertion point from the user
      var promptOptions = new PromptPointOptions("\nSelect top right corner point: ");
      var promptResult = ed.GetPoint(promptOptions);
      if (promptResult.Status != PromptStatus.OK)
        return;

      // Initial point
      var topRightCorner = promptResult.Value;
      var originalTopRightCorner = promptResult.Value;

      // Lowest Y point
      double lowestY = topRightCorner.Y;

      int counter = 0;

      foreach (var panelData in panelDataList)
      {
        bool is2Pole = !panelData.ContainsKey("phase_c_left");
        var endPoint = new Point3d(0, 0, 0);

        using (var tr = db.TransactionManager.StartTransaction())
        {
          var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForWrite);

          CREATEBLOCK();

          // Create initial values
          var startPoint = new Point3d(topRightCorner.X - 8.9856, topRightCorner.Y, 0);
          var layerName = "0";

          // Create the independent header text objects
          CreateTextsWithoutPanelData(tr, layerName, startPoint, is2Pole);

          // Create the dependent header text objects
          CreateTextsWithPanelData(tr, layerName, startPoint, panelData);

          // Create breaker text objects
          ProcessTextData(tr, btr, startPoint, panelData, is2Pole);

          // Get end of data
          var endOfDataY = GetEndOfDataY((List<string>)panelData["description_left"], startPoint);
          endPoint = new Point3d(topRightCorner.X, endOfDataY - 0.2533, 0);

          // Create all the data lines
          ProcessLineData(tr, btr, startPoint, endPoint, endOfDataY, is2Pole);

          // Create footer text objects
          CreateFooterText(tr, endPoint, panelData, is2Pole);

          // Create the middle lines
          CreateCenterLines(btr, tr, startPoint, endPoint, is2Pole);

          // Create the notes section
          CreateNotes(btr, tr, startPoint, endPoint, panelData["existing"] as string);

          // Create the calculations section
          CreateCalculations(btr, tr, startPoint, endPoint, panelData);

          // Create the border of the panel
          CreateRectangle(btr, tr, topRightCorner, startPoint, endPoint, layerName);

          tr.Commit();
        }

        // Check if the endPoint.Y is the lowest point
        if (endPoint.Y < lowestY)
        {
          lowestY = endPoint.Y;
        }

        counter++;

        // After printing 3 panels, reset X and decrease Y by 5
        if (counter % 3 == 0)
        {
          topRightCorner = new Point3d(originalTopRightCorner.X, lowestY - 1.5, 0);
          // Reset lowestY
          lowestY = topRightCorner.Y;
        }
        else
        {
          // Increase x-coordinate by 10 for the next panel
          topRightCorner = new Point3d(topRightCorner.X - 9.6, topRightCorner.Y, 0);
        }
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

    public static (Document doc, Database db, Editor ed) GetGlobals()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var db = doc.Database;
      var ed = doc.Editor;

      return (doc, db, ed);
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

    private static FileInfo GetFileInfo(string filePath)
    {
      var fileInfo = new FileInfo(filePath);
      if (!fileInfo.Exists)
      {
        throw new FileNotFoundException($"The file {filePath} does not exist.");
      }
      return fileInfo;
    }

    private static ExcelPackage GetExcelPackage(FileInfo fileInfo)
    {
      return new ExcelPackage(fileInfo);
    }

    private static bool ValidateWorkbook(ExcelWorkbook workbook)
    {
      if (workbook == null)
      {
        Console.WriteLine("Workbook not found.");
        return false;
      }
      return true;
    }

    private static Dictionary<string, object> ProcessThreePolePanel(ExcelWorksheet selectedWorksheet, int row, int col)
    {
      List<string> descriptionLeft = new List<string>();
      List<string> phaseALeft = new List<string>();
      List<string> phaseBLeft = new List<string>();
      List<string> phaseCLeft = new List<string>();
      List<string> breakerLeft = new List<string>();
      List<string> circuitLeft = new List<string>();
      List<string> circuitRight = new List<string>();
      List<string> breakerRight = new List<string>();
      List<string> phaseARight = new List<string>();
      List<string> phaseBRight = new List<string>();
      List<string> phaseCRight = new List<string>();
      List<string> descriptionRight = new List<string>();

      int panelRow = row + 4;
      int lastRow = panelRow;

      while (selectedWorksheet.Cells[lastRow + 2, col + 6].Value != null)
      {
        lastRow += 2;
      }

      List<bool> descriptionLeftHighlights = new List<bool>();
      List<bool> descriptionRightHighlights = new List<bool>();
      List<bool> breakerLeftHighlights = new List<bool>();
      List<bool> breakerRightHighlights = new List<bool>();

      for (int i = panelRow; i <= lastRow; i++)
      {
        string description = selectedWorksheet.Cells[i, col].Value?.ToString().ToUpper() ?? "SPACE";
        string phaseA = selectedWorksheet.Cells[i, col + 2].Value?.ToString() ?? "0";
        string phaseB = selectedWorksheet.Cells[i, col + 3].Value?.ToString() ?? "0";
        string phaseC = selectedWorksheet.Cells[i, col + 4].Value?.ToString() ?? "0";
        string breaker = selectedWorksheet.Cells[i, col + 5].Value?.ToString() ?? "";
        string circuitL = selectedWorksheet.Cells[i, col + 6].Value?.ToString() ?? "";
        string circuitR = selectedWorksheet.Cells[i, col + 7].Value?.ToString() ?? "";
        string breakerR = selectedWorksheet.Cells[i, col + 8].Value?.ToString() ?? "";
        string phaseAR = selectedWorksheet.Cells[i, col + 9].Value?.ToString() ?? "0";
        string phaseBR = selectedWorksheet.Cells[i, col + 10].Value?.ToString() ?? "0";
        string phaseCR = selectedWorksheet.Cells[i, col + 11].Value?.ToString() ?? "0";
        string descriptionR = selectedWorksheet.Cells[i, col + 12].Value?.ToString().ToUpper() ?? "SPACE";

        bool isLeftHighlighted = selectedWorksheet.Cells[i, col].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isRightHighlighted = selectedWorksheet.Cells[i, col + 12].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isLeftBreakerHighlighted = selectedWorksheet.Cells[i, col + 5].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isRightBreakerHighlighted = selectedWorksheet.Cells[i, col + 8].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";

        descriptionLeft.Add(description);
        phaseALeft.Add(phaseA);
        phaseBLeft.Add(phaseB);
        phaseCLeft.Add(phaseC);
        breakerLeft.Add(breaker);
        circuitLeft.Add(circuitL);
        circuitRight.Add(circuitR);
        breakerRight.Add(breakerR);
        phaseARight.Add(phaseAR);
        phaseBRight.Add(phaseBR);
        phaseCRight.Add(phaseCR);
        descriptionRight.Add(descriptionR);
        descriptionLeftHighlights.Add(isLeftHighlighted);
        descriptionRightHighlights.Add(isRightHighlighted);
        breakerLeftHighlights.Add(isLeftBreakerHighlighted);
        breakerRightHighlights.Add(isRightBreakerHighlighted);
      }

      string panelCellValue = selectedWorksheet.Cells[row, col + 2].Value?.ToString() ?? "";
      if (!panelCellValue.StartsWith("'") || !panelCellValue.EndsWith("'"))
      {
        panelCellValue = "'" + panelCellValue.Trim('\'') + "'";
      }

      var panel = new Dictionary<string, object>
      {
        ["panel"] = panelCellValue,
        ["location"] = selectedWorksheet.Cells[row, col + 5].Value?.ToString() ?? "",
        ["bus_rating"] = selectedWorksheet.Cells[row, col + 9].Value?.ToString() ?? "",
        ["voltage1"] = selectedWorksheet.Cells[row, col + 10].Value?.ToString() ?? "0",
        ["voltage2"] = selectedWorksheet.Cells[row, col + 11].Value?.ToString() ?? "0",
        ["phase"] = selectedWorksheet.Cells[row, col + 12].Value?.ToString() ?? "",
        ["wire"] = selectedWorksheet.Cells[row, col + 13].Value?.ToString() ?? "",
        ["main"] = selectedWorksheet.Cells[row + 1, col + 5].Value?.ToString() ?? "",
        ["mounting"] = selectedWorksheet.Cells[row + 1, col + 12].Value?.ToString() ?? "",
        ["subtotal_a"] = selectedWorksheet.Cells[row + 2, col + 17].Value?.ToString() ?? "0",
        ["subtotal_b"] = selectedWorksheet.Cells[row + 2, col + 18].Value?.ToString() ?? "0",
        ["subtotal_c"] = selectedWorksheet.Cells[row + 2, col + 19].Value?.ToString() ?? "0",
        ["total_va"] = selectedWorksheet.Cells[row + 4, col + 17].Value?.ToString() ?? "0",
        ["lcl"] = selectedWorksheet.Cells[row + 7, col + 17].Value?.ToString() ?? "0",
        ["total_other_load"] = selectedWorksheet.Cells[row + 10, col + 17].Value?.ToString() ?? "0",
        ["kva"] = selectedWorksheet.Cells[row + 13, col + 17].Value?.ToString() ?? "0",
        ["feeder_amps"] = selectedWorksheet.Cells[row + 16, col + 17].Value?.ToString() ?? "0",
        ["existing"] = selectedWorksheet.Cells[row + 2, col + 20].Value?.ToString() ?? "",
        ["description_left_highlights"] = descriptionLeftHighlights,
        ["description_right_highlights"] = descriptionRightHighlights,
        ["breaker_left_highlights"] = breakerLeftHighlights,
        ["breaker_right_highlights"] = breakerRightHighlights,
        ["description_left"] = descriptionLeft,
        ["phase_a_left"] = phaseALeft,
        ["phase_b_left"] = phaseBLeft,
        ["phase_c_left"] = phaseCLeft,
        ["breaker_left"] = breakerLeft,
        ["circuit_left"] = circuitLeft,
        ["circuit_right"] = circuitRight,
        ["breaker_right"] = breakerRight,
        ["phase_a_right"] = phaseARight,
        ["phase_b_right"] = phaseBRight,
        ["phase_c_right"] = phaseCRight,
        ["description_right"] = descriptionRight,
      };

      ReplaceInPanel(panel, "voltage2", "V");
      ReplaceInPanel(panel, "phase", "PH");
      ReplaceInPanel(panel, "wire", "W");

      return panel;
    }

    private static Dictionary<string, object> ProcessTwoPolePanel(ExcelWorksheet selectedWorksheet, int row, int col)
    {
      List<string> descriptionLeft = new List<string>();
      List<string> phaseALeft = new List<string>();
      List<string> phaseBLeft = new List<string>();
      List<string> breakerLeft = new List<string>();
      List<string> circuitLeft = new List<string>();
      List<string> circuitRight = new List<string>();
      List<string> breakerRight = new List<string>();
      List<string> phaseARight = new List<string>();
      List<string> phaseBRight = new List<string>();
      List<string> descriptionRight = new List<string>();
      List<bool> descriptionLeftHighlights = new List<bool>();
      List<bool> descriptionRightHighlights = new List<bool>();
      List<bool> breakerLeftHighlights = new List<bool>();
      List<bool> breakerRightHighlights = new List<bool>();

      int panelRow = row + 4;
      int lastRow = panelRow;

      // check for a circuit column cell value of NULL to end the loop
      while (selectedWorksheet.Cells[lastRow + 2, col + 6].Value != null)
      {
        lastRow += 2;
      }

      // add cell values to lists
      for (int i = panelRow; i <= lastRow; i++)
      {
        string description = selectedWorksheet.Cells[i, col].Value?.ToString().ToUpper() ?? "SPACE";
        string phaseA = selectedWorksheet.Cells[i, col + 3].Value?.ToString() ?? "0";
        string phaseB = selectedWorksheet.Cells[i, col + 4].Value?.ToString() ?? "0";
        string breakerL = selectedWorksheet.Cells[i, col + 5].Value?.ToString() ?? "";
        string circuitL = selectedWorksheet.Cells[i, col + 6].Value?.ToString() ?? "";
        string circuitR = selectedWorksheet.Cells[i, col + 7].Value?.ToString() ?? "";
        string breakerR = selectedWorksheet.Cells[i, col + 8].Value?.ToString() ?? "";
        string phaseAR = selectedWorksheet.Cells[i, col + 9].Value?.ToString() ?? "0";
        string phaseBR = selectedWorksheet.Cells[i, col + 10].Value?.ToString() ?? "0";
        string descriptionR = selectedWorksheet.Cells[i, col + 11].Value?.ToString().ToUpper() ?? "SPACE";

        bool isLeftHighlighted = selectedWorksheet.Cells[i, col].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isRightHighlighted = selectedWorksheet.Cells[i, col + 11].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isLeftBreakerHighlighted = selectedWorksheet.Cells[i, col + 5].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isRightBreakerHighlighted = selectedWorksheet.Cells[i, col + 8].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";

        descriptionLeft.Add(description);
        phaseALeft.Add(phaseA);
        phaseBLeft.Add(phaseB);
        breakerLeft.Add(breakerL);
        circuitLeft.Add(circuitL);
        circuitRight.Add(circuitR);
        breakerRight.Add(breakerR);
        phaseARight.Add(phaseAR);
        phaseBRight.Add(phaseBR);
        descriptionRight.Add(descriptionR);
        descriptionLeftHighlights.Add(isLeftHighlighted);
        descriptionRightHighlights.Add(isRightHighlighted);
        breakerLeftHighlights.Add(isLeftBreakerHighlighted);
        breakerRightHighlights.Add(isRightBreakerHighlighted);
      }

      string panelCellValue = selectedWorksheet.Cells[row, col + 2].Value?.ToString() ?? "";
      if (!panelCellValue.StartsWith("'") || !panelCellValue.EndsWith("'"))
      {
        panelCellValue = "'" + panelCellValue.Trim('\'') + "'";
      }

      var panel = new Dictionary<string, object>
      {
        ["panel"] = panelCellValue,
        ["location"] = selectedWorksheet.Cells[row, col + 5].Value?.ToString() ?? "",
        ["bus_rating"] = selectedWorksheet.Cells[row, col + 9].Value?.ToString() ?? "",
        ["voltage1"] = selectedWorksheet.Cells[row, col + 10].Value?.ToString() ?? "0",
        ["voltage2"] = selectedWorksheet.Cells[row, col + 11].Value?.ToString() ?? "0",
        ["phase"] = selectedWorksheet.Cells[row, col + 12].Value?.ToString() ?? "",
        ["wire"] = selectedWorksheet.Cells[row, col + 13].Value?.ToString() ?? "",
        ["main"] = selectedWorksheet.Cells[row + 1, col + 5].Value?.ToString() ?? "",
        ["mounting"] = selectedWorksheet.Cells[row + 1, col + 12].Value?.ToString() ?? "",
        ["subtotal_a"] = selectedWorksheet.Cells[row + 2, col + 17].Value?.ToString() ?? "0",
        ["subtotal_b"] = selectedWorksheet.Cells[row + 2, col + 18].Value?.ToString() ?? "0",
        ["subtotal_c"] = selectedWorksheet.Cells[row + 2, col + 19].Value?.ToString() ?? "0",
        ["total_va"] = selectedWorksheet.Cells[row + 4, col + 17].Value?.ToString() ?? "0",
        ["lcl"] = selectedWorksheet.Cells[row + 7, col + 17].Value?.ToString() ?? "0",
        ["total_other_load"] = selectedWorksheet.Cells[row + 10, col + 17].Value?.ToString() ?? "0",
        ["kva"] = selectedWorksheet.Cells[row + 13, col + 17].Value?.ToString() ?? "0",
        ["feeder_amps"] = selectedWorksheet.Cells[row + 16, col + 17].Value?.ToString() ?? "0",
        ["existing"] = selectedWorksheet.Cells[row + 2, col + 20].Value?.ToString() ?? "",
        ["description_left_highlights"] = descriptionLeftHighlights,
        ["description_right_highlights"] = descriptionRightHighlights,
        ["breaker_left_highlights"] = breakerLeftHighlights,
        ["breaker_right_highlights"] = breakerRightHighlights,
        ["description_left"] = descriptionLeft,
        ["phase_a_left"] = phaseALeft,
        ["phase_b_left"] = phaseBLeft,
        ["breaker_left"] = breakerLeft,
        ["circuit_left"] = circuitLeft,
        ["circuit_right"] = circuitRight,
        ["breaker_right"] = breakerRight,
        ["phase_a_right"] = phaseARight,
        ["phase_b_right"] = phaseBRight,
        ["description_right"] = descriptionRight,
      };

      ReplaceInPanel(panel, "voltage2", "V");
      ReplaceInPanel(panel, "phase", "PH");
      ReplaceInPanel(panel, "wire", "W");

      return panel;
    }

    private static void ReplaceInPanel(Dictionary<string, object> panel, string key, string toRemove)
    {
      if (panel.ContainsKey(key))
      {
        if (panel[key] is string value)
        {
          value = value.Replace(toRemove, "");
          panel[key] = value;
        }
      }
    }

    private static List<Dictionary<string, object>> ProcessWorksheet(ExcelWorksheet selectedWorksheet)
    {
      var (doc, db, ed) = GetGlobals();
      var panels = new List<Dictionary<string, object>>();
      int rowCount = selectedWorksheet.Dimension.Rows;
      int colCount = selectedWorksheet.Dimension.Columns;

      for (int row = 1; row <= rowCount; row++)
      {
        for (int col = 1; col <= colCount; col++)
        {
          string cellValue = selectedWorksheet.Cells[row, col].Value?.ToString();
          string phaseAMaybe = selectedWorksheet.Cells[row + 3, col + 2].Value?.ToString();
          if (cellValue == "PANEL:" && phaseAMaybe == "PH A")
          {
            panels.Add(ProcessThreePolePanel(selectedWorksheet, row, col));
          }
          else if (cellValue == "PANEL:")
          {
            panels.Add(ProcessTwoPolePanel(selectedWorksheet, row, col));
          }
        }
      }
      return panels;
    }

    private static void HandleExceptions(System.Exception ex)
    {
      Console.WriteLine(ex.Message);
    }

    public void KeepBreakersGivenPoints(Point3d point1, Point3d point2, Point3d point3)
    {
      var (doc, db, ed) = MyCommands.GetGlobals();

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        // Get the active space block table record instead of the paperspace block table record
        var activeSpaceId = db.CurrentSpaceId;
        var btr = (BlockTableRecord)tr.GetObject(activeSpaceId, OpenMode.ForWrite);

        // Rest of the code remains the same...

        // Swap the points if the y-coordinate of the first point is lower than that of the second point
        if (point1.Y < point2.Y)
        {
          Point3d tempPoint = point1;
          point1 = point2;
          point2 = tempPoint;
        }

        var direction = point3.X > point1.X ? 1 : -1;
        var dist = (point1 - point2).Length;

        var line1Start = new Point3d(point1.X + direction * 0.05, point1.Y, 0);
        var line1End = new Point3d(line1Start.X + direction * 0.2, line1Start.Y, 0);
        var line2Start = new Point3d(line1Start.X, point2.Y, 0);
        var line2End = new Point3d(line1End.X, line2Start.Y, 0);

        string layerName = CreateOrGetLayer("E-TEXT", db, tr);

        var line1 = new Line(line1Start, line1End) { Layer = layerName, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2) };
        var line2 = new Line(line2Start, line2End) { Layer = layerName, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2) };

        var mid1 = new Point3d((line1Start.X + line1End.X) / 2, (line1Start.Y + line1End.Y) / 2, 0);
        var mid2 = new Point3d((line2Start.X + line2End.X) / 2, (line2Start.Y + line2End.Y) / 2, 0);
        var mid3 = new Point3d((mid1.X + mid2.X) / 2, (mid1.Y + mid2.Y) / 2, 0);

        var circleTop = new Point3d(mid3.X, mid3.Y + 0.09, 0);
        var circleBottom = new Point3d(mid3.X, mid3.Y - 0.09, 0);

        var line3 = new Line(mid1, circleTop) { Layer = layerName, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2) };
        var line4 = new Line(mid2, circleBottom) { Layer = layerName, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2) };

        if (dist > 0.3)
        {
          btr.AppendEntity(line1);
          btr.AppendEntity(line2);
          tr.AddNewlyCreatedDBObject(line1, true);
          tr.AddNewlyCreatedDBObject(line2, true);
        }

        var blkRef = new BlockReference(mid3, bt["CIRCLEI"]) { Layer = layerName, Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2) };
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

    private static List<Dictionary<string, object>> ImportExcelData()
    {
      var (doc, db, ed) = GetGlobals();

      var openFileDialog = new System.Windows.Forms.OpenFileDialog
      {
        Filter = "Excel Files|*.xlsx;*.xls",
        Title = "Select Excel File"
      };

      List<Dictionary<string, object>> panels = new List<Dictionary<string, object>>();

      if (openFileDialog.ShowDialog() == DialogResult.OK)
      {
        string filePath = openFileDialog.FileName;
        try
        {
          FileInfo fileInfo = GetFileInfo(filePath);
          using (var package = GetExcelPackage(fileInfo))
          {
            ExcelWorkbook workbook = package.Workbook;
            if (ValidateWorkbook(workbook))
            {
              foreach (var selectedWorksheet in workbook.Worksheets)
              {
                if (selectedWorksheet.Name.ToLower().Contains("panel"))
                {
                  panels.AddRange(ProcessWorksheet(selectedWorksheet));
                }
              }
            }
          }
        }
        catch (FileNotFoundException ex)
        {
          HandleExceptions(ex);
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
          HandleExceptions(ex);
        }
      }
      else
      {
        Console.WriteLine("No file selected.");
      }

      return panels;
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

    private void CreateCalculations(BlockTableRecord btr, Transaction tr, Point3d startPoint, Point3d endPoint, Dictionary<string, object> panelData)
    {
      var (_, _, ed) = GetGlobals();
      double kvaValue;
      double feederAmpsValue;
      var kvaParseResult = double.TryParse(panelData["kva"] as string, out kvaValue);
      var feederAmpsParseResult = double.TryParse(panelData["feeder_amps"] as string, out feederAmpsValue);

      ed.WriteMessage($"KVA VALUE: {panelData["kva"] as string}\n");
      ed.WriteMessage($"FEEDER AMPS VALUE: {panelData["feeder_amps"] as string}\n");

      if (kvaParseResult)
      {
        CreateAndPositionRightText(tr, Math.Round(kvaValue, 1).ToString("0.0") + " KVA", "ROMANS", 0.09375, 1, 2, "PNLTXT", new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.785594790702817, 0));
      }
      else
      {
        ed.WriteMessage($"Error: Unable to convert 'kva' value to double: {panelData["kva"]}");
      }

      if (feederAmpsParseResult)
      {
        CreateAndPositionRightText(tr, Math.Round(feederAmpsValue, 1).ToString("0.0") + " A", "ROMANS", 0.09375, 1, 2, "PNLTXT", new Point3d(endPoint.X - 6.70142386189229, endPoint.Y - 0.970762733814496, 0));
      }
      else
      {
        ed.WriteMessage($"Error: Unable to convert 'feeder_amps' value to double: {panelData["feeder_amps"]}");
      }

      // Create the calculation lines
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 0.0846396524177919, endPoint.X - 8.98559999999998, endPoint.Y - 0.0846396524177919, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 1.02063965241777, endPoint.X - 6.17759999999998, endPoint.Y - 0.0846396524177919, "0");
      CreateLine(tr, btr, endPoint.X - 8.98559999999998, endPoint.Y - 1.02063965241777, endPoint.X - 8.98559999999998, endPoint.Y - 0.0846396524177919, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 1.02063965241777, endPoint.X - 8.98559999999998, endPoint.Y - 1.02063965241777, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 0.833439652417809, endPoint.X - 8.98559999999998, endPoint.Y - 0.833439652417809, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 0.64623965241779, endPoint.X - 8.98559999999998, endPoint.Y - 0.64623965241779, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 0.459039652417772, endPoint.X - 8.98559999999998, endPoint.Y - 0.459039652417772, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 0.27183965241781, endPoint.X - 8.98559999999998, endPoint.Y - 0.27183965241781, "0");
      CreateLine(tr, btr, endPoint.X - 6.17759999999998, endPoint.Y - 0.0846396524177919, endPoint.X - 8.98559999999998, endPoint.Y - 0.0846396524177919, "0");

      // Create the text
      CreateAndPositionText(tr, "TOTAL CONNECTED VA", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 8.93821353998555, endPoint.Y - 0.244065644556514, 0));
      CreateAndPositionRightText(tr, panelData["total_va"] as string, "ROMANS", 0.09375, 1, 2, "PNLTXT", new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.222040136230106, 0));
      CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.242614932747216, 0));
      CreateAndPositionText(tr, "LCL @ 125 %          ", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 8.91077927366155, endPoint.Y - 0.432165907882307, 0));
      CreateAndPositionRightText(tr, "0", "ROMANS", 0.09375, 1, 2, "PNLTXT", new Point3d(endPoint.X - 7.59414061117746, endPoint.Y - 0.413648726513742, 0));
      CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.437756414851634, 0));
      CreateAndPositionRightText(tr, panelData["lcl"] as string, "ROMANS", 0.09375, 1, 2, "PNLTXT", new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.413648726513742, 0));
      CreateAndPositionText(tr, "TOTAL OTHER LOAD", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 8.9456956126196, endPoint.Y - 0.616854044919108, 0));
      CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.618180694030713, 0));
      CreateAndPositionRightText(tr, panelData["total_other_load"] as string, "ROMANS", 0.09375, 1, 2, "PNLTXT", new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.597206513223341, 0));
      CreateAndPositionText(tr, "PANEL LOAD", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 8.92075537050664, endPoint.Y - 0.804954308244959, 0));
      CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.809218166102625, 0));
      CreateAndPositionText(tr, "FEEDER AMPS", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 8.9120262857673, endPoint.Y - 0.994381220682413, 0));
      CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.998928989062989, 0));
    }

    private void CreateNotes(BlockTableRecord btr, Transaction tr, Point3d startPoint, Point3d endPoint, string panelType)
    {
      if (panelType.ToLower() == "existing")
      {
        CreateAndPositionText(tr, "(EXISTING PANEL)", "ROMANC", 0.1498, 0.75, 2, "0", new Point3d(startPoint.X + 0.236635303895696, startPoint.Y + 0.113254677317428, 0));

        // Create the text
        CreateAndPositionText(tr, "NOTES:", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 5.96783070435049, endPoint.Y - 0.23875904811004, 0));
        CreateAndPositionText(tr, "DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.405747901076808, 0));
        CreateAndPositionText(tr, "TO MATCH EXISTING.", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.610352149436778, 0));
      }
      else if (panelType.ToLower() == "relocated")
      {
        CreateAndPositionText(tr, "(EXISTING TO BE RELOCATED PANEL)", "ROMANC", 0.1498, 0.75, 2, "0", new Point3d(startPoint.X + 0.236635303895696, startPoint.Y + 0.113254677317428, 0));

        // Create the text
        CreateAndPositionText(tr, "NOTES:", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 5.96783070435049, endPoint.Y - 0.23875904811004, 0));
        CreateAndPositionText(tr, "DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.405747901076808, 0));
        CreateAndPositionText(tr, "TO MATCH EXISTING.", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.610352149436778, 0));
      }
      else
      {
        CreateAndPositionText(tr, "(NEW PANEL)", "ROMANC", 0.1498, 0.75, 2, "0", new Point3d(startPoint.X + 0.236635303895696, startPoint.Y + 0.113254677317428, 0));

        // Create the text
        CreateAndPositionText(tr, "NOTES:", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 5.96783070435049, endPoint.Y - 0.23875904811004, 0));
        CreateAndPositionText(tr, "65 KAIC SERIES RATED OR MATCH FAULT CURRENT AT SITE.", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.405747901076808, 0));
      }

      // Create the lines
      CreateLine(tr, btr, endPoint.X, endPoint.Y - 0.0846396524177919, endPoint.X - 6.07359999999994, endPoint.Y - 0.0846396524177919, "0");
      CreateLine(tr, btr, endPoint.X, endPoint.Y - 0.0846396524177919, endPoint.X, endPoint.Y - 1.02063965241777, "0");
      CreateLine(tr, btr, endPoint.X - 6.07359999999994, endPoint.Y - 1.02063965241777, endPoint.X - 6.07359999999994, endPoint.Y + -0.0846396524177919, "0");
      CreateLine(tr, btr, endPoint.X, endPoint.Y - 1.02063965241777, endPoint.X - 6.07359999999994, endPoint.Y - 1.02063965241777, "0");
      CreateLine(tr, btr, endPoint.X, endPoint.Y - 0.27183965241781, endPoint.X - 6.07359999999994, endPoint.Y - 0.27183965241781, "0");
      CreateLine(tr, btr, endPoint.X, endPoint.Y - 0.459039652417772, endPoint.X - 6.07359999999994, endPoint.Y - 0.459039652417772, "0");
      CreateLine(tr, btr, endPoint.X, endPoint.Y - 0.64623965241779, endPoint.X - 6.07359999999994, endPoint.Y + -0.64623965241779, "0");
      CreateLine(tr, btr, endPoint.X, endPoint.Y + -0.833439652417809, endPoint.X - 6.07359999999994, endPoint.Y + -0.833439652417809, "0");

      // Create the circle
      CreateCircle(btr, tr, new Point3d(endPoint.X - 5.8088, endPoint.Y - 0.3664, 0), 0.09, 2, false);

      // Create the 1
      CreateAndPositionCenteredText(tr, "1", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 5.85897687070053 - 0.145, endPoint.Y - 0.410151417346867, 0));
    }

    private void CreateCenterLines(BlockTableRecord btr, Transaction tr, Point3d startPoint, Point3d endPoint, bool is2Pole)
    {
      // Create horizontal line above
      CreateLine(tr, btr, startPoint.X, endPoint.Y + 0.2533, endPoint.X, endPoint.Y + 0.2533, "0");

      if (is2Pole)
      {
        // Create the slashed line
        CreateLine(tr, btr, endPoint.X - 6.47536463134611, endPoint.Y + 0.0841286798547145, endPoint.X - 6.3618865297326, endPoint.Y + 0.216793591015815, "0");
        CreateLine(tr, btr, endPoint.X - 4.541248855498, endPoint.Y + 0.0739331046861764, endPoint.X - 4.42777075388449, endPoint.Y + 0.206598015847277, "0");

        // Create the vertical center lines
        CreateLine(tr, btr, endPoint.X - 4.67999999999984, endPoint.Y + 0.277148912917994, endPoint.X - 4.67999999999984, startPoint.Y - 0.7488, "0");
        CreateLine(tr, btr, endPoint.X - 4.30668381424084, endPoint.Y + 0.277148912917994, endPoint.X - 4.30668381424084, startPoint.Y - 0.7488, "0");
      }
      else
      {
        // Create the slashed line
        CreateLine(tr, btr, endPoint.X - 6.47536463134611, endPoint.Y + 0.0663322399757078, endPoint.X - 6.36188652973283, endPoint.Y + 0.198997151136808, "0");
        CreateLine(tr, btr, endPoint.X - 4.63023473191822, endPoint.Y + 0.0663322399757078, endPoint.X - 4.51675663030494, endPoint.Y + 0.198997151136808, "0");
        CreateLine(tr, btr, endPoint.X - 2.06710862405464, endPoint.Y + 0.0663322399757078, endPoint.X - 1.95363052244136, endPoint.Y + 0.198997151136808, "0");

        // Create the vertical center lines
        CreateLine(tr, btr, endPoint.X - 4.30559999999991, endPoint.Y + 0.253292187434056, endPoint.X - 4.30559999999991, startPoint.Y - 0.7488, "0");
        CreateLine(tr, btr, endPoint.X - 4.49279999999999, endPoint.Y + 0.253292187434056, endPoint.X - 4.49279999999999, startPoint.Y - 0.7488, "0");
        CreateLine(tr, btr, endPoint.X - 4.67999999999984, endPoint.Y + 0.253292187434056, endPoint.X - 4.67999999999984, startPoint.Y - 0.7488, "0");
      }
      // Create the circle and lines in the center of the panel
      CreateCenterLinePattern(tr, btr, startPoint, endPoint, is2Pole);
    }

    private void CreateCenterLinePattern(Transaction tr, BlockTableRecord btr, Point3d startPoint, Point3d endPoint, bool is2Pole)
    {
      double maxY = endPoint.Y + 0.2533;
      double increaseY = 0.1872;
      double increaseX = 0.1872;
      double baseX = startPoint.X + 4.3056;
      double currentX = baseX;
      double currentY = startPoint.Y - 0.8424;
      int num = 3;

      if (is2Pole)
      {
        baseX = startPoint.X + 4.30560000000014;
        currentX = baseX;
        increaseX = 0.3733;
        num = 2;
      }

      bool conditionMet = false;

      while (currentY >= maxY && !conditionMet)
      {
        for (int i = 0; i < num; i++)
        {
          if (currentY < maxY)
          {
            conditionMet = true;
            break;
          }

          // Create the center line circles
          CreateCircle(btr, tr, new Point3d(currentX, currentY, 0), 0.0312, 7);

          // Create the horizontal center lines
          CreateLine(tr, btr, startPoint.X + 4.22905693965708, currentY, startPoint.X + 4.75654306034312, currentY, "0");

          currentX += increaseX;
          currentY -= increaseY;
        }

        // reset x value
        currentX = baseX;
      }
    }

    public void CreateCircle(BlockTableRecord btr, Transaction tr, Point3d center, double radius, int colorIndex, bool doHatch = true)
    {
      using (Circle circle = new Circle())
      {
        circle.Center = center;
        circle.Radius = radius;
        circle.ColorIndex = colorIndex;  // Setting the color
        circle.Layer = "0";  // Setting the layer to "0"
        btr.AppendEntity(circle);
        tr.AddNewlyCreatedDBObject(circle, true);

        if (doHatch)
        {
          // Creating a Hatch
          using (Hatch hatch = new Hatch())
          {
            hatch.Layer = "0";  // Setting the layer to "0"
            btr.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            hatch.Associative = true;

            // Associating the hatch with the circle
            ObjectIdCollection objIds = new ObjectIdCollection();
            objIds.Add(circle.ObjectId);
            hatch.AppendLoop(HatchLoopTypes.Default, objIds);
            hatch.EvaluateHatch(true);
          }
        }
      }
    }

    private void CreateFooterText(Transaction tr, Point3d endPoint, Dictionary<string, object> panelData, bool is2Pole)
    {
      if (is2Pole)
      {
        CreateAndPositionText(tr, "SUB-TOTAL", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 8.91077927366177, endPoint.Y + 0.0697891578365528, 0));
        CreateAndPositionText(tr, "OA", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 6.45042438923338, endPoint.Y + 0.0920885745242259, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 4.51630861338526, endPoint.Y + 0.0818929993556878, 0));
        CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 6.24591440390805, endPoint.Y + 0.0920885745242259, 0));
        CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 4.32551576122205, endPoint.Y + 0.0818929993556878, 0));
        CreateAndPositionText(tr, panelData["subtotal_a"] as string + "VA", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 6.08199405082502, endPoint.Y + 0.108280531454625, 0));
        CreateAndPositionText(tr, panelData["subtotal_b"] as string + "VA", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 4.15962890179264, endPoint.Y + 0.0980849562860868, 0));
      }
      else
      {
        CreateAndPositionText(tr, "SUB-TOTAL", "Standard", 0.1248, 1, 7, "0", new Point3d(endPoint.X - 8.91077927366155, endPoint.Y + 0.0689855381989162, 0));
        CreateAndPositionText(tr, "OA", "Standard", 0.1248, 0.75, 7, "0", new Point3d(endPoint.X - 6.45042438923338, endPoint.Y + 0.0742921346453898, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.1248, 0.75, 7, "0", new Point3d(endPoint.X - 4.60529448980549, endPoint.Y + 0.0742921346453898, 0));
        CreateAndPositionText(tr, "OC", "Standard", 0.1248, 0.75, 7, "0", new Point3d(endPoint.X - 2.04216838194191, endPoint.Y + 0.0742921346453898, 0));
        CreateAndPositionText(tr, panelData["subtotal_a"] as string + "VA", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 6.07732066030258, endPoint.Y + 0.0948263267698053, 0));
        CreateAndPositionText(tr, panelData["subtotal_b"] as string + "VA", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 4.23219076087469, endPoint.Y + 0.0948263267698053, 0));
        CreateAndPositionText(tr, panelData["subtotal_c"] as string + "VA", "ROMANS", 0.09375, 1, 2, "0", new Point3d(endPoint.X - 1.66906465301099, endPoint.Y + 0.0948263267698053, 0));
        CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 6.24591440390827, endPoint.Y + 0.0742921346453898, 0));
        CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 4.40078450448038, endPoint.Y + 0.0742921346453898, 0));
        CreateAndPositionText(tr, "=", "Standard", 0.1248, 0.75, 256, "0", new Point3d(endPoint.X - 1.8376583966168, endPoint.Y + 0.0742921346453898, 0));
      }
    }

    private double GetEndOfDataY(List<string> list, Point3d startPoint)
    {
      var rowHeight = 0.1872;
      var headerHeight = 0.7488;
      return startPoint.Y - (headerHeight + (rowHeight * ((list.Count + 1) / 2)));
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

    private void CreateRectangle(BlockTableRecord btr, Transaction tr, Point3d topRightCorner, Point3d startPoint, Point3d endPoint, string layerName)
    {
      // Create the rectangle
      var rect = new Autodesk.AutoCAD.DatabaseServices.Polyline(4);
      rect.AddVertexAt(0, new Point2d(startPoint.X, startPoint.Y), 0, 0, 0);
      rect.AddVertexAt(1, new Point2d(startPoint.X, endPoint.Y), 0, 0, 0);
      rect.AddVertexAt(2, new Point2d(endPoint.X, endPoint.Y), 0, 0, 0);
      rect.AddVertexAt(3, new Point2d(endPoint.X, startPoint.Y), 0, 0, 0);
      rect.Closed = true;

      // Set the global width property
      rect.ConstantWidth = 0.02;

      // Set the layer to "0"
      rect.Layer = layerName;

      btr.AppendEntity(rect);
      tr.AddNewlyCreatedDBObject(rect, true);
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

    private void CreateAndPositionFittedText(Transaction tr, string content, string style, double height, double widthFactor, int colorIndex, string layerName, Point3d position, double length, TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft, TextVerticalMode verticalMode = TextVerticalMode.TextBase)
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, (short)colorIndex);
      var textId = CreateText(content, style, horizontalMode, verticalMode, height, widthFactor, color, layerName);
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);

      double naturalWidth = text.GeometricExtents.MaxPoint.X - text.GeometricExtents.MinPoint.X;
      text.WidthFactor = length / naturalWidth; // This will stretch or squeeze text to fit between points
      text.Position = position;
    }

    private void CreateAndPositionCenteredText(Transaction tr, string content, string style, double height, double widthFactor, int colorIndex, string layerName, Point3d position, TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft, TextVerticalMode verticalMode = TextVerticalMode.TextBase)
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, (short)colorIndex);
      var textId = CreateText(content, style, horizontalMode, verticalMode, height, widthFactor, color, layerName);
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
      text.Justify = AttachmentPoint.BaseCenter;
      double x = position.X;
      text.AlignmentPoint = new Point3d(x + 0.1903, position.Y, 0);
    }

    private void CreateAndPositionRightText(Transaction tr, string content, string style, double height, double widthFactor, int colorIndex, string layerName, Point3d position, TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft, TextVerticalMode verticalMode = TextVerticalMode.TextBase)
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, (short)colorIndex);
      var textId = CreateText(content, style, horizontalMode, verticalMode, height, widthFactor, color, layerName);
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
      text.Justify = AttachmentPoint.BaseRight;
      double x = position.X;
      text.AlignmentPoint = new Point3d(x + 0.46, position.Y, 0);
    }

    private void CreateTextsWithoutPanelData(Transaction tr, string layerName, Point3d startPoint, bool is2Pole)
    {
      CreateAndPositionText(tr, "PANEL", "ROMANC", 0.1872, 0.75, 0, layerName, new Point3d(startPoint.X + 0.231944251649111, startPoint.Y - 0.299822699224023, 0));
      CreateAndPositionText(tr, "DESCRIPTION", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 0.305517965881791, startPoint.Y - 0.638118222684739, 0));
      CreateAndPositionText(tr, "W", "Standard", 0.101088, 0.75, 0, layerName, new Point3d(startPoint.X + 8.64365164909793, startPoint.Y - 0.155688865359394, 0));
      CreateAndPositionText(tr, "VOLT AMPS", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 1.9015733562577, startPoint.Y - 0.532524377875689, 0));
      CreateAndPositionText(tr, "L", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 2.97993751651882, startPoint.Y - 0.483601235896458, 0));
      CreateAndPositionText(tr, "T", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 2.97993751651882, startPoint.Y - 0.59526740969153, 0));
      CreateAndPositionText(tr, "G", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 2.97993751651882, startPoint.Y - 0.702157646684782, 0));
      CreateAndPositionText(tr, "R", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 3.20889406685785, startPoint.Y - 0.482921120531671, 0));
      CreateAndPositionText(tr, "E", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 3.20889406685785, startPoint.Y - 0.594587294326715, 0));
      CreateAndPositionText(tr, "C", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 3.20889406685785, startPoint.Y - 0.701477531319966, 0));
      CreateAndPositionText(tr, "M", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 3.43493724520761, startPoint.Y - 0.482921120531671, 0));
      CreateAndPositionText(tr, "I", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 3.4427934214732, startPoint.Y - 0.594587294326715, 0));
      CreateAndPositionText(tr, "S", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 3.43493724520761, startPoint.Y - 0.701477531319966, 0));
      CreateAndPositionText(tr, "BKR", "Standard", 0.09152, 0.75, 256, "0", new Point3d(startPoint.X + 3.63691080609988, startPoint.Y - 0.61662650707666, 0));
      CreateAndPositionText(tr, "CKT", "Standard", 0.0832, 0.75, 256, "0", new Point3d(startPoint.X + 3.94429929014041, startPoint.Y - 0.529332995532684, 0));
      CreateAndPositionText(tr, " NO", "Standard", 0.0832, 0.75, 256, "0", new Point3d(startPoint.X + 3.90688892697108, startPoint.Y - 0.673306258645766, 0));
      CreateAndPositionText(tr, "BUS", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 4.32282163085404, startPoint.Y - 0.527068325709052, 0));
      CreateAndPositionText(tr, "CKT", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 4.88897460099258, startPoint.Y - 0.535275052777223, 0));
      CreateAndPositionText(tr, " NO", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 4.85530527414039, startPoint.Y - 0.664850989579008, 0));
      CreateAndPositionText(tr, "BKR", "Standard", 0.082368, 0.75, 256, "0", new Point3d(startPoint.X + 5.14497871612878, startPoint.Y - 0.612478980835647, 0));
      CreateAndPositionText(tr, "M", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.4736003885796, startPoint.Y - 0.483601235896458, 0));
      CreateAndPositionText(tr, "I", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.48257887574016, startPoint.Y - 0.59526740969153, 0));
      CreateAndPositionText(tr, "S", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.4736003885796, startPoint.Y - 0.702157646684782, 0));
      CreateAndPositionText(tr, "R", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.70588736710022, startPoint.Y - 0.482921120531671, 0));
      CreateAndPositionText(tr, "E", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.70588736710022, startPoint.Y - 0.594587294326715, 0));
      CreateAndPositionText(tr, "C", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.70588736710022, startPoint.Y - 0.701477531319966, 0));
      CreateAndPositionText(tr, "L", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.93367350805136, startPoint.Y - 0.484281352862808, 0));
      CreateAndPositionText(tr, "T", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.93367350805136, startPoint.Y - 0.595947526657881, 0));
      CreateAndPositionText(tr, "G", "Standard", 0.07488, 0.75, 256, "0", new Point3d(startPoint.X + 5.93367350805136, startPoint.Y - 0.702837763651132, 0));
      CreateAndPositionText(tr, "VOLT AMPS", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 6.32453930091015, startPoint.Y - 0.532297673821773, 0));
      CreateAndPositionText(tr, "DESCRIPTION", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 7.68034755863846, startPoint.Y - 0.636791573573134, 0));
      CreateAndPositionText(tr, "LOCATION", "Standard", 0.11232, 0.75, 0, layerName, new Point3d(startPoint.X + 2.32067207718262, startPoint.Y - 0.155059196495415, 0));
      CreateAndPositionText(tr, "MAIN (AMP)", "Standard", 0.11232, 0.75, 0, layerName, new Point3d(startPoint.X + 2.32089885857886, startPoint.Y - 0.338479316609039, 0));
      CreateAndPositionText(tr, "BUS RATING", "Standard", 0.1248, 0.75, 0, layerName, new Point3d(startPoint.X + 5.18507633525223, startPoint.Y - 0.271963067880222, 0));
      CreateAndPositionText(tr, "MOUNTING:", "Standard", 0.11232, 0.75, 0, layerName, new Point3d(startPoint.X + 7.01560982102967, startPoint.Y - 0.329154148660905, 0));
      CreateAndPositionText(tr, "V", "Standard", 0.09984, 0.75, 0, layerName, new Point3d(startPoint.X + 7.80112268015148, startPoint.Y - 0.158231303238949, 0));
      CreateAndPositionText(tr, "O", "Standard", 0.101088, 0.75, 0, layerName, new Point3d(startPoint.X + 8.30325740318381, startPoint.Y - 0.151432601608803, 0));

      if (is2Pole)
      {
        CreateAndPositionText(tr, "OA", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 1.87939466183889, startPoint.Y - 0.720370467604425, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 2.50641160863188, startPoint.Y - 0.720370467604425, 0));
        CreateAndPositionText(tr, "OA", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 4.19245469916268, startPoint.Y - 0.720370467604425, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 4.59766144739842, startPoint.Y - 0.720370467604425, 0));
        CreateAndPositionText(tr, "OA", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 6.2528343633212, startPoint.Y - 0.720370467604425, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.1248, 0.75, 256, "0", new Point3d(startPoint.X + 6.91903366501083, startPoint.Y - 0.720370467604425, 0));
      }
      else
      {
        CreateAndPositionText(tr, "OA", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 1.75923320841673, startPoint.Y - 0.715823582939777, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 2.17200390182074, startPoint.Y - 0.714690056264089, 0));
        CreateAndPositionText(tr, "OC", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 2.57149885762158, startPoint.Y - 0.718725420176355, 0));
        CreateAndPositionText(tr, "OA", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 4.22229040316483, startPoint.Y - 0.714236644953189, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 4.42655098606872, startPoint.Y - 0.714236644953189, 0));
        CreateAndPositionText(tr, "OC", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 4.63417850165774, startPoint.Y - 0.713042660752734, 0));
        CreateAndPositionText(tr, "OA", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 6.22324655852697, startPoint.Y - 0.71537017323044, 0));
        CreateAndPositionText(tr, "OB", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 6.63397621936463, startPoint.Y - 0.714690057865624, 0));
        CreateAndPositionText(tr, "OC", "Standard", 0.11232, 0.75, 256, "0", new Point3d(startPoint.X + 7.03324439586629, startPoint.Y - 0.718272010467018, 0));
      }
    }

    private void CreateTextsWithPanelData(Transaction tr, string layerName, Point3d startPoint, Dictionary<string, object> panelData)
    {
      CreateAndPositionText(tr, panelData["panel"] as string, "ROMANC", 0.1872, 0.75, 2, layerName, new Point3d(startPoint.X + 1.17828457810867, startPoint.Y - 0.299822699224023, 0));
      CreateAndPositionText(tr, panelData["location"] as string, "ROMANS", 0.09375, 1, 2, layerName, new Point3d(startPoint.X + 3.19605976175148, startPoint.Y - 0.137807184107345, 0));
      CreateAndPositionText(tr, panelData["main"] as string, "ROMANS", 0.09375, 1, 2, layerName, new Point3d(startPoint.X + 3.24033367283675, startPoint.Y - 0.32590837886957, 0));
      CreateAndPositionText(tr, panelData["bus_rating"] as string, "ROMANS", 0.12375, 1, 2, layerName, new Point3d(startPoint.X + 6.2073642121926, startPoint.Y - 0.274622599308543, 0));
      CreateAndPositionText(tr, panelData["voltage1"] as string + "/" + panelData["voltage2"] as string, "ROMANS", 0.09375, 1, 2, layerName, new Point3d(startPoint.X + 7.04393671550224, startPoint.Y - 0.141653203021775, 0));
      CreateAndPositionText(tr, panelData["mounting"] as string, "ROMANS", 0.09375, 1, 2, layerName, new Point3d(startPoint.X + 7.87802551675406, startPoint.Y - 0.331292901876935, 0));
      CreateAndPositionText(tr, panelData["phase"] as string, "ROMANS", 0.09375, 1, 2, layerName, new Point3d(startPoint.X + 8.1253996026328, startPoint.Y - 0.141653203021775, 0));
      CreateAndPositionText(tr, panelData["wire"] as string, "ROMANS", 0.09375, 1, 2, layerName, new Point3d(startPoint.X + 8.50104048135836, startPoint.Y - 0.141653203021775, 0));
    }

    private void CreateVerticalLines(Transaction tr, BlockTableRecord btr, Point3d startPoint, double[] distances, double startY, double endY, string layer)
    {
      foreach (double distance in distances)
      {
        var lineStart = new Point3d(startPoint.X + distance, startY, 0);
        var lineEnd = new Point3d(startPoint.X + distance, endY, 0);
        var line = new Line(lineStart, lineEnd);
        line.Layer = layer;

        btr.AppendEntity(line);
        tr.AddNewlyCreatedDBObject(line, true);
      }
    }

    private void CreateLines(Transaction tr, BlockTableRecord btr, IEnumerable<(double startX, double startY, double endX, double endY, string layer)> lines)
    {
      foreach (var (startX, startY, endX, endY, layer) in lines)
      {
        CreateLine(tr, btr, startX, startY, endX, endY, layer);
      }
    }

    private void CreateLine(Transaction tr, BlockTableRecord btr, double startX, double startY, double endX, double endY, string layer)
    {
      var lineStart = new Point3d(startX, startY, 0);
      var lineEnd = new Point3d(endX, endY, 0);
      var line = new Line(lineStart, lineEnd);
      line.Layer = layer;

      btr.AppendEntity(line);
      tr.AddNewlyCreatedDBObject(line, true);
    }

    private void ProcessTextData(Transaction tr, BlockTableRecord btr, Point3d startPoint, Dictionary<string, object> panelData, bool is2Pole)
    {
      var (_, _, ed) = GetGlobals();

      List<string> leftDescriptions = (List<string>)panelData["description_left"];
      List<string> leftBreakers = (List<string>)panelData["breaker_left"];
      List<string> leftCircuits = (List<string>)panelData["circuit_left"];
      List<string> leftPhaseA = (List<string>)panelData["phase_a_left"];
      List<string> leftPhaseB = (List<string>)panelData["phase_b_left"];

      List<string> rightDescriptions = (List<string>)panelData["description_right"];
      List<string> rightBreakers = (List<string>)panelData["breaker_right"];
      List<string> rightCircuits = (List<string>)panelData["circuit_right"];
      List<string> rightPhaseA = (List<string>)panelData["phase_a_right"];
      List<string> rightPhaseB = (List<string>)panelData["phase_b_right"];

      List<bool> leftBreakersHighlight = (List<bool>)panelData["breaker_left_highlights"];
      List<bool> rightBreakersHighlight = (List<bool>)panelData["breaker_right_highlights"];

      if (!is2Pole)
      {
        List<string> leftPhaseC = (List<string>)panelData["phase_c_left"];
        List<string> rightPhaseC = (List<string>)panelData["phase_c_right"];
        // Use the ProcessSideData for the left side
        ProcessSideData(tr, btr, startPoint, leftDescriptions, leftBreakers, leftCircuits, leftPhaseA, leftPhaseB, (List<string>)leftPhaseC, (List<bool>)panelData["description_left_highlights"], (List<bool>)panelData["breaker_left_highlights"], true);

        // Use the ProcessSideData for the right side
        ProcessSideData(tr, btr, startPoint, rightDescriptions, rightBreakers, rightCircuits, rightPhaseA, rightPhaseB, (List<string>)rightPhaseC, (List<bool>)panelData["description_right_highlights"], (List<bool>)panelData["breaker_right_highlights"], false);
      }
      else
      {
        // Use the ProcessSideData for the left side
        ProcessSideData2P(tr, btr, startPoint, leftDescriptions, leftBreakers, leftCircuits, leftPhaseA, leftPhaseB, (List<bool>)panelData["description_left_highlights"], (List<bool>)panelData["breaker_left_highlights"], true);

        // Use the ProcessSideData for the right side
        ProcessSideData2P(tr, btr, startPoint, rightDescriptions, rightBreakers, rightCircuits, rightPhaseA, rightPhaseB, (List<bool>)panelData["description_right_highlights"], (List<bool>)panelData["breaker_right_highlights"], false);
      }

      InsertKeepBreakers(startPoint, leftBreakersHighlight, true);
      InsertKeepBreakers(startPoint, rightBreakersHighlight, false);
    }

    private void InsertKeepBreakers(Point3d startPoint, List<bool> breakersHighlight, bool left)
    {
      var (_, _, ed) = GetGlobals();
      double header_height = 0.7488;
      double panel_width = 8.9856;
      double row_height = 0.1872;
      double left_start_x = startPoint.X;
      double left_start_y = startPoint.Y - header_height;
      double right_start_x = startPoint.X + panel_width;
      double right_start_y = left_start_y;
      double start_x, start_y, displacement;
      Point3d topPoint = new Point3d(0, 0, 0);
      Point3d botPoint = new Point3d(0, 0, 0);
      bool currentlyKeeping = false;

      ed.WriteMessage(breakersHighlight.Count.ToString());
      for (int i = 0; i < breakersHighlight.Count; i += 2)
      {
        if (left)
        {
          start_x = left_start_x;
          start_y = left_start_y;
          displacement = -1;
        }
        else
        {
          start_x = right_start_x;
          start_y = right_start_y;
          displacement = 1;
        }

        ed.WriteMessage($"The value of breakersHighlight[i]: {breakersHighlight[i]} \nThe value of i: {i}\n");

        if (breakersHighlight[i] && !currentlyKeeping)
        {
          topPoint = new Point3d(start_x, start_y - (row_height * (i / 2)), 0);
          currentlyKeeping = true;
        }
        else if (!breakersHighlight[i] && currentlyKeeping)
        {
          botPoint = new Point3d(start_x, start_y - (row_height * (i / 2)), 0);
          currentlyKeeping = false;
          KeepBreakersGivenPoints(topPoint, botPoint, new Point3d(topPoint.X + displacement, topPoint.Y, 0));
          ed.WriteMessage($"\nWent inside to make the breakers, value of i: {i}\n");
          ed.WriteMessage($"\nTop point value: {topPoint}\nBot point value: {botPoint}\n");
        }
        else if (breakersHighlight[i] && currentlyKeeping && i >= breakersHighlight.Count - 2)
        {
          botPoint = new Point3d(start_x, start_y - (row_height * ((i + 2) / 2)), 0);
          KeepBreakersGivenPoints(topPoint, botPoint, new Point3d(topPoint.X + displacement, topPoint.Y, 0));
        }
      }
    }

    private void ProcessSideData(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> breakers, List<string> circuits, List<string> phaseA, List<string> phaseB, List<string> phaseC, List<bool> descriptionHighlights, List<bool> breakerHighlights, bool left)
    {
      var (_, _, ed) = GetGlobals();

      Dictionary<string, double> data = new Dictionary<string, double>();

      data.Add("row height y", 0.1872);
      data.Add("half row height y", 0.0936);
      data.Add("initial half breaker text y", -0.816333638994546);
      data.Add("header height", 0.7488);

      for (int i = 0; i < descriptions.Count; i += 2)
      {
        double phase = GetPhase(breakers, circuits, i);

        if (phase == 0.5)
        {
          CreateHalfBreaker(tr, btr, startPoint, descriptions, phaseA, phaseB, phaseC, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
        }
        else if (phase == 1.0)
        {
          Create1PoleBreaker(tr, btr, startPoint, descriptions, phaseA, phaseB, phaseC, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
        }
        else if (phase == 2.0)
        {
          Create2PoleBreaker(tr, btr, startPoint, descriptions, phaseA, phaseB, phaseC, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
          i += 2;
        }
        else
        {
          Create3PoleBreaker(tr, btr, startPoint, descriptions, phaseA, phaseB, phaseC, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
          i += 4;
        }
      }
    }

    private void ProcessSideData2P(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> breakers, List<string> circuits, List<string> phaseA, List<string> phaseB, List<bool> descriptionHighlights, List<bool> breakerHighlights, bool left)
    {
      var (_, _, ed) = GetGlobals();

      Dictionary<string, double> data = new Dictionary<string, double>
      {
        { "row height y", 0.1872 },
        { "half row height y", 0.0936 },
        { "initial half breaker text y", -0.816333638994546 },
        { "header height", 0.7488 }
      };

      for (int i = 0; i < descriptions.Count; i += 2)
      {
        double phase = GetPhase(breakers, circuits, i);

        if (phase == 0.5)
        {
          CreateHalfBreaker2P(tr, btr, startPoint, descriptions, phaseA, phaseB, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
        }
        else if (phase == 1.0)
        {
          Create1PoleBreaker2P(tr, btr, startPoint, descriptions, phaseA, phaseB, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
        }
        else
        {
          Create2PoleBreaker2P(tr, btr, startPoint, descriptions, phaseA, phaseB, breakers, circuits, descriptionHighlights, breakerHighlights, data, left, i);
          i += 2;
        }
      }
    }

    private void CreateHalfBreaker(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> phaseC, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      List<string> phaseList = GetPhaseList(i, phaseA, phaseB, phaseC);

      double descriptionX = GetDescriptionX(startPoint, left);
      double phaseX = GetPhaseX(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);

      string circuit = circuits[i];

      for (var j = i; j <= i + 1; j++)
      {
        string description = (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD") ? "(E)" + descriptions[j] : descriptions[j];
        string breaker = breakers[j] + "-1";
        string phase = phaseList[j];
        circuit = circuits[j];
        double height = startPoint.Y + (-0.816333638994546 - ((double)j / 2 * 0.1872));

        CreateAndPositionText(tr, description, "ROMANS", 0.046875, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
        if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.046875, 1.0, 2, "0", new Point3d(phaseX, height, 0));
        CreateAndPositionText(tr, breaker, "ROMANS", 0.046875, 1.0, 2, "0", new Point3d(breakerX, height, 0));
        CreateAndPositionText(tr, circuit, "ROMANS", 0.046875, 1.0, 7, "0", new Point3d(circuitX, height, 0));
      }

      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void CreateHalfBreaker2P(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      List<string> phaseList = GetPhaseList2P(i, phaseA, phaseB);

      double descriptionX = GetDescriptionX2P(startPoint, left);
      double phaseX = GetPhaseX2P(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);

      string circuit = circuits[i];

      for (var j = i; j <= i + 1; j++)
      {
        string description = (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD") ? "(E)" + descriptions[j] : descriptions[j];
        string breaker = breakers[j] + "-1";
        string phase = phaseList[j];
        circuit = circuits[j];
        double height = startPoint.Y + (-0.816333638994546 - ((double)j / 2 * 0.1872));

        CreateAndPositionText(tr, description, "ROMANS", 0.046875, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
        if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.046875, 1.0, 2, "0", new Point3d(phaseX, height, 0));
        CreateAndPositionText(tr, breaker, "ROMANS", 0.046875, 1.0, 2, "0", new Point3d(breakerX, height, 0));
        CreateAndPositionText(tr, circuit, "ROMANS", 0.046875, 1.0, 7, "0", new Point3d(circuitX, height, 0));
      }

      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void Create1PoleBreaker(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> phaseC, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      List<string> phaseList = GetPhaseList(i, phaseA, phaseB, phaseC);

      string description = (descriptionHighlights[i] && descriptions[i] != "EXISTING LOAD") ? "(E)" + descriptions[i] : descriptions[i];
      string breaker = breakers[i] + "-1";
      string phase = phaseList[i];
      string circuit = circuits[i];
      double height = startPoint.Y + (-0.890211813771344 - ((i / 2) * 0.1872));
      double descriptionX = GetDescriptionX(startPoint, left);
      double phaseX = GetPhaseX(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.2300;

      CreateAndPositionText(tr, description, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
      if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(phaseX, height, 0));
      if (breaker != "-1") CreateAndPositionFittedText(tr, breaker, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(breakerX, height, 0), length);
      CreateAndPositionText(tr, circuit, "ROMANS", 0.09375, 1.0, 7, "0", new Point3d(circuitX, height, 0));
      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void Create1PoleBreaker2P(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      List<string> phaseList = GetPhaseList2P(i, phaseA, phaseB);

      string description = (descriptionHighlights[i] && descriptions[i] != "EXISTING LOAD") ? "(E)" + descriptions[i] : descriptions[i];
      string breaker = breakers[i] + "-1";
      string phase = phaseList[i];
      string circuit = circuits[i];
      double height = startPoint.Y + (-0.890211813771344 - ((i / 2) * 0.1872));
      double descriptionX = GetDescriptionX2P(startPoint, left);
      double phaseX = GetPhaseX2P(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.2300;

      CreateAndPositionText(tr, description, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
      if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(phaseX, height, 0));
      if (breaker != "-1") CreateAndPositionFittedText(tr, breaker, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(breakerX, height, 0), length);
      CreateAndPositionText(tr, circuit, "ROMANS", 0.09375, 1.0, 7, "0", new Point3d(circuitX, height, 0));
      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void Create2PoleBreaker(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> phaseC, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      double descriptionX = GetDescriptionX(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.14;

      for (var j = i; j <= i + 2; j += 2)
      {
        List<string> phaseList = GetPhaseList(j, phaseA, phaseB, phaseC);
        string description = (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD") ? "(E)" + descriptions[j] : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
          breakerX += 0.16;
          length = 0.07;
        }

        CreateAndPositionText(tr, description, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
        if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(phaseX, height, 0));
        if (breaker != "") CreateAndPositionFittedText(tr, breaker, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(breakerX, height, 0), length);
        CreateAndPositionText(tr, circuit, "ROMANS", 0.09375, 1.0, 7, "0", new Point3d(circuitX, height, 0));
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }

      CreateBreakerLine(startPoint, i, left, tr, btr, 4);
    }

    private void Create2PoleBreaker2P(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      var (_, _, ed) = GetGlobals();
      double descriptionX = GetDescriptionX2P(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.14;

      for (var j = i; j <= i + 2; j += 2)
      {
        List<string> phaseList = GetPhaseList2P(j, phaseA, phaseB);
        string description = (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD") ? "(E)" + descriptions[j] : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX2P(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
          breakerX += 0.16;
          length = 0.07;
        }

        CreateAndPositionText(tr, description, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
        if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(phaseX, height, 0));
        ed.WriteMessage($"The value of breaker is: {breaker}\nThe value of breakerX is: {breakerX}\nThe value of length is: {length}");
        if (breaker != "") CreateAndPositionFittedText(tr, breaker, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(breakerX, height, 0), length);
        CreateAndPositionText(tr, circuit, "ROMANS", 0.09375, 1.0, 7, "0", new Point3d(circuitX, height, 0));
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }

      CreateBreakerLine(startPoint, i, left, tr, btr, 4);
    }

    private void Create3PoleBreaker(Transaction tr, BlockTableRecord btr, Point3d startPoint, List<string> descriptions, List<string> phaseA, List<string> phaseB, List<string> phaseC, List<string> breakers, List<string> circuits, List<bool> descriptionHighlights, List<bool> breakerHighlights, Dictionary<string, double> data, bool left, int i)
    {
      double descriptionX = GetDescriptionX(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.14;

      for (var j = i; j <= i + 4; j += 2)
      {
        List<string> phaseList = GetPhaseList(j, phaseA, phaseB, phaseC);
        string description = (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD") ? "(E)" + descriptions[j] : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
        }
        else if (j == i + 4)
        {
          description = "---";
          breakerX += 0.16;
          length = 0.07;
        }

        CreateAndPositionText(tr, description, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(descriptionX, height, 0));
        if (phase != "0") CreateAndPositionCenteredText(tr, phase, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(phaseX, height, 0));
        if (j != i + 2) CreateAndPositionFittedText(tr, breaker, "ROMANS", 0.09375, 1.0, 2, "0", new Point3d(breakerX, height, 0), length);
        CreateAndPositionText(tr, circuit, "ROMANS", 0.09375, 1.0, 7, "0", new Point3d(circuitX, height, 0));
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }

      CreateBreakerLine(startPoint, i, left, tr, btr, 6);
    }

    private void CreateBreakerLine(Point3d startPoint, int i, bool left, Transaction tr, BlockTableRecord btr, int span)
    {
      double x1, x2;
      double height = startPoint.Y + (-0.7488 - (((i + span) / 2) * 0.1872));
      double y1 = height;
      double y2 = height + (span / 2) * 0.1872;

      if (left)
      {
        x1 = startPoint.X + 3.588;
        x2 = startPoint.X + 3.9;
      }
      else
      {
        x1 = startPoint.X + 5.0856;
        x2 = startPoint.X + 5.3976;
      }

      var lineStart = new Point3d(x1, y1, 0);
      var lineEnd = new Point3d(x2, y2, 0);
      var line = new Line(lineStart, lineEnd);
      line.Layer = "0";
      line.ColorIndex = 2;

      btr.AppendEntity(line);
      tr.AddNewlyCreatedDBObject(line, true);
    }

    private double GetCircuitX(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 3.93681721750636;
      }
      else
      {
        return startPoint.X + 4.87281721750651;
      }
    }

    private double GetBreakerX(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 3.60379818231218;
      }
      else
      {
        return startPoint.X + 5.10947444486385;
      }
    }

    private double GetDescriptionX(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 0.063560431161136;
      }
      else
      {
        return startPoint.X + 7.43528640590171;
      }
    }

    private double GetDescriptionX2P(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 0.0536663060360638;
      }
      else
      {
        return startPoint.X + 7.40509162108179;
      }
    }

    public double GetPhaseX(int i, Point3d startPoint, bool left)
    {
      if (left)
      {
        if (i % 6 == 0)
        {
          return startPoint.X + 1.64526228334811;
        }
        else if (i % 6 == 2)
        {
          return startPoint.X + 2.0792421731542;
        }
        else
        {
          return startPoint.X + 2.50445478897294;
        }
      }
      else
      {
        if (i % 6 == 0)
        {
          return startPoint.X + 6.11211889838299;
        }
        else if (i % 6 == 2)
        {
          return startPoint.X + 6.53328984899773;
        }
        else
        {
          return startPoint.X + 6.96804695722213;
        }
      }
    }

    public double GetPhaseX2P(int i, Point3d startPoint, bool left)
    {
      if (left)
      {
        if (i % 4 == 0)
        {
          return startPoint.X + 1.8390082793234;
        }
        else
        {
          return startPoint.X + 2.39546408826883;
        }
      }
      else
      {
        if (i % 4 == 0)
        {
          return startPoint.X + 6.21960728338948;
        }
        else
        {
          return startPoint.X + 6.83021158846114;
        }
      }
    }

    public List<string> GetPhaseList(int i, List<string> phaseA, List<string> phaseB, List<string> phaseC)
    {
      if (i % 6 == 0)
      {
        return phaseA;
      }
      else if (i % 6 == 2)
      {
        return phaseB;
      }
      else
      {
        return phaseC;
      }
    }

    public List<string> GetPhaseList2P(int i, List<string> phaseA, List<string> phaseB)
    {
      if (i % 4 == 0)
      {
        return phaseA;
      }
      else
      {
        return phaseB;
      }
    }

    private string RemoveLetter(string circuit)
    {
      if (circuit.Contains('A') || circuit.Contains('B'))
      {
        return circuit.Replace("A", "").Replace("B", "");
      }
      return circuit;
    }

    private double GetPhase(List<string> breakers, List<string> circuits, int i)
    {
      // If index i is out of range for the circuits list, return 0.0
      if (i >= circuits.Count) return 0.0;

      // Check if circuit at index i contains 'A' or 'B'
      if (circuits[i].Contains('A') || circuits[i].Contains('B'))
      {
        return 0.5;
      }
      // Check if breakers has a value at [i+2] and if it is '3'
      else if ((i + 4) < breakers.Count && breakers[i + 4] == "3")
      {
        return 3.0;
      }
      // Check if breakers has a value at [i+1] and if it is '2'
      else if ((i + 2) < breakers.Count && breakers[i + 2] == "2")
      {
        return 2.0;
      }
      else
      {
        return 1.0;
      }
    }

    private void ProcessLineData(Transaction tr, BlockTableRecord btr, Point3d startPoint, Point3d endPoint, double endOfDataY, bool is2Pole)
    {
      string layerName = "0";
      double left1, left2, left3, right1, right2, right3;

      if (is2Pole)
      {
        left1 = 1.7549;
        left2 = 2.3023;
        left3 = 0;
        right1 = 6.7259;
        right2 = 0;
        right3 = 7.3632;
        CreateLine(tr, btr, startPoint.X + 1.85445441972615, startPoint.Y - 0.728330362273937, startPoint.X + 1.96793252133921, startPoint.Y - 0.595665451113291, "0");
        CreateLine(tr, btr, startPoint.X + 2.48147136651869, startPoint.Y - 0.728330362273937, startPoint.X + 2.5949494681322, startPoint.Y - 0.595665451113291, "0");
        CreateLine(tr, btr, startPoint.X + 4.16751445704995, startPoint.Y - 0.728330362273937, startPoint.X + 4.280992558663, startPoint.Y - 0.595665451113291, "0");
        CreateLine(tr, btr, startPoint.X + 4.57272120528569, startPoint.Y - 0.728330362273937, startPoint.X + 4.68619930689874, startPoint.Y - 0.595665451113291, "0");
        CreateLine(tr, btr, startPoint.X + 6.22789412120846, startPoint.Y - 0.728330362273937, startPoint.X + 6.34137222282197, startPoint.Y - 0.595665451113291, "0");
        CreateLine(tr, btr, startPoint.X + 6.8940934228981, startPoint.Y - 0.728330362273937, startPoint.X + 7.0075715245116, startPoint.Y - 0.595665451113291, "0");
      }
      else
      {
        left1 = 1.6224;
        left2 = 2.0488;
        left3 = 2.4752;
        right1 = 6.5104;
        right2 = 6.9368;
        right3 = 7.3632;
        CreateLine(tr, btr, startPoint.X + 1.8219640114711, startPoint.Y + -0.587355127494504, startPoint.X + 1.75209866364992, startPoint.Y + -0.732578250164607, "0");
        CreateLine(tr, btr, startPoint.X + 2.2392415498706, startPoint.Y + -0.587355127494504, startPoint.X + 2.16937620204942, startPoint.Y + -0.732578250164607, "0");
        CreateLine(tr, btr, startPoint.X + 2.64064439053459, startPoint.Y + -0.587355127494504, startPoint.X + 2.57077904271353, startPoint.Y + -0.732578250164607, "0");
        CreateLine(tr, btr, startPoint.X + 4.28558110707343, startPoint.Y + -0.581743684459752, startPoint.X + 4.21812491713047, startPoint.Y + -0.728299423953047, "0");
        CreateLine(tr, btr, startPoint.X + 4.4919520727949, startPoint.Y + -0.581743684459752, startPoint.X + 4.42449588285183, startPoint.Y + -0.728299423953047, "0");
        CreateLine(tr, btr, startPoint.X + 4.69832301754843, startPoint.Y + -0.581743684459752, startPoint.X + 4.63086682760547, startPoint.Y + -0.728299423953047, "0");
        CreateLine(tr, btr, startPoint.X + 6.28764850398056, startPoint.Y + -0.586040159740406, startPoint.X + 6.21478297900239, startPoint.Y + -0.730701330926394, "0");
        CreateLine(tr, btr, startPoint.X + 6.69812260049363, startPoint.Y + -0.586040159740406, startPoint.X + 6.62525707551547, startPoint.Y + -0.730701330926394, "0");
        CreateLine(tr, btr, startPoint.X + 7.10859666269549, startPoint.Y + -0.586040159740406, startPoint.X + 7.03573113771732, startPoint.Y + -0.730701330926394, "0");
      }

      var linesData = new (double[] distances, double startY, double endY, string layer)[]
          {
                    (new double[] { 2.2222, 5.0666, 6.9368 }, startPoint.Y, startPoint.Y - 0.3744, layerName),
                    (new double[] { left1, 2.9016, 3.1304, 3.3592, 3.5880, 3.9000, 4.1496, 4.8360, 5.0856, 5.3976, 5.6264, 5.8552, 6.0840, right3 }, startPoint.Y - 0.3744, endOfDataY, layerName),
                    (new double[] { left2, left3, right1, right2 }, startPoint.Y - 0.3744 - (0.3744 / 2), endOfDataY, layerName)
          };

      foreach (var lineData in linesData)
      {
        CreateVerticalLines(tr, btr, startPoint, lineData.distances, lineData.startY, lineData.endY, lineData.layer);
      }

      var linesData2 = new (double startX, double startY, double endX, double endY, string layer)[]
      {
                    (startPoint.X, startPoint.Y - 0.3744, endPoint.X, startPoint.Y - 0.3744, layerName),
                    (startPoint.X, startPoint.Y - 0.7488, endPoint.X, startPoint.Y - 0.7488, layerName),
                    (startPoint.X + 2.2222, startPoint.Y - (0.3744/2), startPoint.X + 5.0666, startPoint.Y - (0.3744/2), layerName),
                    (startPoint.X + 6.9368, startPoint.Y - (0.3744/2), endPoint.X, startPoint.Y - (0.3744/2), layerName),
                    (startPoint.X + left1, startPoint.Y - (0.3744/2) - 0.3744, startPoint.X + 2.9016, startPoint.Y - (0.3744/2) - 0.3744, layerName),
                    (startPoint.X + 6.0840, startPoint.Y - (0.3744/2) - 0.3744, startPoint.X + 7.3632, startPoint.Y - (0.3744/2) - 0.3744, layerName),
                    (startPoint.X + 4.1496, startPoint.Y - (0.3744/2) - 0.3744, startPoint.X + 4.8360, startPoint.Y - (0.3744/2) - 0.3744, layerName),
                    (startPoint.X + 8.28490642235897, startPoint.Y + -0.153581773169606, startPoint.X + 8.37682368466574, startPoint.Y + -0.0461231951291552, "0")
      };

      CreateLines(tr, btr, linesData2);
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

    private double CreateHorizontalLine(double startPointX, double startPointY, string circuitNumber, bool left, Transaction tr, BlockTableRecord btr)
    {
      int circuitNumReducer;
      double lineStartX;
      double lineStartX2;
      int circuitNumInt;
      double deltaY = 0.187200000000021; // Change this value if needed

      if (left)
      {
        circuitNumReducer = 1;
        lineStartX = startPointX;
        lineStartX2 = startPointX + 4.14960000000019;
      }
      else
      {
        circuitNumReducer = 2;
        lineStartX = startPointX + 4.8360;
        lineStartX2 = startPointX + 8.9856;
      }

      if (circuitNumber.Contains('A') || circuitNumber.Contains('B'))
      {
        // Remove 'A' or 'B' from the string
        circuitNumber = circuitNumber.Replace("A", "").Replace("B", "");
        circuitNumInt = int.Parse(circuitNumber);
        CreateCircuitLine(circuitNumInt, circuitNumReducer, startPointY, deltaY, lineStartX, lineStartX2, tr, btr, true);
      }
      else
      {
        circuitNumInt = int.Parse(circuitNumber);
      }

      return CreateCircuitLine(circuitNumInt, circuitNumReducer, startPointY, deltaY, lineStartX, lineStartX2, tr, btr);
    }

    private double CreateCircuitLine(int circuitNumInt, int circuitNumReducer, double startPointY, double deltaY, double lineStartX, double lineStartX2, Transaction tr, BlockTableRecord btr, bool half = false)
    {
      circuitNumInt = (circuitNumInt - circuitNumReducer) / 2;
      double lineStartY = startPointY - (0.935999999999979 + (deltaY * circuitNumInt));
      if (half) lineStartY += deltaY / 2;
      double lineEndY = lineStartY;

      var lineStart = new Point3d(lineStartX, lineStartY, 0);
      var lineEnd = new Point3d(lineStartX2, lineEndY, 0);
      var line = new Line(lineStart, lineEnd)
      {
        Layer = "0"
      };

      btr.AppendEntity(line);
      tr.AddNewlyCreatedDBObject(line, true);

      return line.StartPoint.Y;
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

  public class HotkeyManager
  {
    public static string scale = "1/4";

    [CommandMethod("SETSCALE")]
    public void SetScale()
    {
      Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
      PromptResult result = ed.GetString("\nEnter new scale (e.g., 3/16, 1/8, 3/32, 1/16): ");

      if (result.Status == PromptStatus.OK)
      {
        scale = result.StringResult;
        ed.WriteMessage($"\nScale set to {scale}");
      }
      else
      {
        ed.WriteMessage("\nScale not changed.");
      }
    }

    [CommandMethod("N")]
    public void CreateNote()
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptStringOptions pso = new PromptStringOptions("\nEnter text content for the note: ");
      PromptResult pr = ed.GetString(pso);
      if (pr.Status != PromptStatus.OK) return;

      string textContent = pr.StringResult;

      PromptPointResult ppr = ed.GetPoint("Click where you want the center of the hexagon: ");
      if (ppr.Status != PromptStatus.OK) return;

      Point3d center = ppr.Value;
      double scaleFactor = GetScaleFactor();
      double sideLength = 5.1962 * scaleFactor;
      double textHeight = 4.5 * scaleFactor;

      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        // Create hexagon
        Polyline hexagon = new Polyline(6);
        for (int i = 0; i < 6; i++)
        {
          double angle = (Math.PI / 3.0) * i;
          double x = center.X + sideLength * Math.Cos(angle);
          double y = center.Y + sideLength * Math.Sin(angle);
          hexagon.AddVertexAt(i, new Point2d(x, y), 0, 0, 0);
        }
        hexagon.Closed = true;
        hexagon.Layer = "E-TXT1";
        btr.AppendEntity(hexagon);
        tr.AddNewlyCreatedDBObject(hexagon, true);

        // Create text object
        DBText text = new DBText();
        text.Position = center;
        text.Height = textHeight;
        text.TextString = textContent;
        text.HorizontalMode = TextHorizontalMode.TextCenter;
        text.VerticalMode = TextVerticalMode.TextVerticalMid;
        text.AlignmentPoint = center;
        text.Layer = "E-TXT1";
        text.WidthFactor = 0.85; // Set the width factor to 0.85

        // Set text style
        TextStyleTable tst = (TextStyleTable)tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
        if (tst.Has("rpm"))
        {
          text.TextStyleId = tst["rpm"];
        }

        btr.AppendEntity(text);
        tr.AddNewlyCreatedDBObject(text, true);

        tr.Commit();
      }

      ed.WriteMessage("\nHexagon and text created.");
    }

    [CommandMethod("PU")]
    public void CreatePanelUp()
    {
      DrawPanel(0);
    }

    [CommandMethod("PD")]
    public void CreatePanelDown()
    {
      DrawPanel(180);
    }

    [CommandMethod("PR")]
    public void CreatePanelRight()
    {
      DrawPanel(270);
    }

    [CommandMethod("PL")]
    public void CreatePanelLeft()
    {
      DrawPanel(90);
    }

    private void DrawPanel(double rotationAngle)
    {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptPointResult ppr = ed.GetPoint("Click where you want the base point of the panel: ");
      if (ppr.Status != PromptStatus.OK) return;

      Point3d basePoint = ppr.Value;
      double scaleFactor = GetScaleFactor();

      double width = 20.0;
      double height = 6.0;
      double horizontalLineLength = 26.2940;
      double verticalLineLength = 7.5311;
      double ellipseHeight = 6.3907 * scaleFactor;
      double textHeight = 4.5 * scaleFactor;

      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        // Create a rotation matrix
        Matrix3d rotationMatrix = Matrix3d.Rotation(rotationAngle * (Math.PI / 180), Vector3d.ZAxis, basePoint);

        // Create rectangle
        Polyline rectangle = new Polyline();
        rectangle.AddVertexAt(0, new Point2d(basePoint.X - width / 2, basePoint.Y), 0, 0, 0);
        rectangle.AddVertexAt(1, new Point2d(basePoint.X + width / 2, basePoint.Y), 0, 0, 0);
        rectangle.AddVertexAt(2, new Point2d(basePoint.X + width / 2, basePoint.Y + height), 0, 0, 0);
        rectangle.AddVertexAt(3, new Point2d(basePoint.X - width / 2, basePoint.Y + height), 0, 0, 0);
        rectangle.Closed = true;
        rectangle.Layer = "E-SYM1";
        ObjectId rectId = btr.AppendEntity(rectangle);
        tr.AddNewlyCreatedDBObject(rectangle, true);

        // Hatch the inside of the rectangle
        Hatch hatch = new Hatch();
        hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
        ObjectIdCollection idCol = new ObjectIdCollection { rectId };
        hatch.AppendLoop(HatchLoopTypes.Default, idCol);
        hatch.EvaluateHatch(true);
        hatch.Layer = "E-SYM1";

        // Apply the rotation to both the rectangle and the hatch
        rectangle.TransformBy(rotationMatrix);
        hatch.TransformBy(rotationMatrix);

        // Add hatch to BlockTableRecord
        btr.AppendEntity(hatch);
        tr.AddNewlyCreatedDBObject(hatch, true);

        // Create horizontal line on top of the rectangle
        Line horizontalLine = new Line(
            new Point3d(basePoint.X - horizontalLineLength / 2, basePoint.Y + height, 0),
            new Point3d(basePoint.X + horizontalLineLength / 2, basePoint.Y + height, 0)
        );
        horizontalLine.Layer = "E-SYM1";

        // Create vertical line from the top of the rectangle's height
        Line verticalLine = new Line(
            new Point3d(basePoint.X, basePoint.Y + height, 0),
            new Point3d(basePoint.X, basePoint.Y + height + verticalLineLength, 0)
        );
        verticalLine.Layer = "E-TXT1";

        // Prompt user for text content
        PromptStringOptions stringOptions = new PromptStringOptions("Enter text content: ");
        PromptResult stringResult = ed.GetString(stringOptions);
        string textContent = stringResult.StringResult.ToUpper();

        double baseEllipseHeight = 6.3907;
        double baseEllipseWidth = 10 + (textContent.Length - 1) * 4;
        ellipseHeight = baseEllipseHeight * scaleFactor;
        double ellipseWidth = baseEllipseWidth * scaleFactor;
        Ellipse ellipse = new Ellipse(
            new Point3d(basePoint.X, basePoint.Y + height + verticalLineLength + ellipseHeight / 2, 0),
            Vector3d.ZAxis,
            new Vector3d(ellipseWidth / 2, 0, 0),
            baseEllipseHeight / baseEllipseWidth,
            0,
            Math.PI * 2
        );
        ellipse.Layer = "E-TXT1";

        // Create centered text inside the ellipse
        DBText text = new DBText
        {
          TextString = textContent,
          Height = textHeight,
          Position = new Point3d(basePoint.X, basePoint.Y + height + verticalLineLength + ellipseHeight / 2, 0),
          HorizontalMode = TextHorizontalMode.TextCenter,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          AlignmentPoint = new Point3d(basePoint.X, basePoint.Y + height + verticalLineLength + ellipseHeight / 2, 0),
          Layer = "E-TXT1",
          WidthFactor = 0.85 // Set the width factor to 0.85
        };

        // Set text style
        TextStyleTable tst = (TextStyleTable)tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
        if (tst.Has("rpm"))
        {
          text.TextStyleId = tst["rpm"];
        }

        text.AdjustAlignment(doc.Database);

        // Transform and add the other objects to the BlockTableRecord
        Entity[] entities = { horizontalLine, verticalLine, ellipse, text };
        foreach (Entity entity in entities)
        {
          entity.TransformBy(rotationMatrix);
          btr.AppendEntity(entity);
          tr.AddNewlyCreatedDBObject(entity, true);
        }

        // Create a rotation matrix to rotate the text and ellipse back to 0 degrees
        Matrix3d reverseRotationMatrix = Matrix3d.Rotation(-rotationAngle * (Math.PI / 180), Vector3d.ZAxis, ellipse.Center);

        // Apply the reverse rotation to the text and ellipse
        text.TransformBy(reverseRotationMatrix);
        ellipse.TransformBy(reverseRotationMatrix);

        double shiftAmount = (1.8047 + 2 * (textContent.Length - 1)) * scaleFactor;
        if (shiftAmount < 0) shiftAmount = 0;

        // Apply the shift for the "PR" command (270 degrees rotation)
        if (rotationAngle == 270)
        {
          Matrix3d shiftMatrix = Matrix3d.Displacement(new Vector3d(shiftAmount, 0, 0));
          text.TransformBy(shiftMatrix);
          ellipse.TransformBy(shiftMatrix);
        }

        // Apply the shift for the "PL" command (90 degrees rotation)
        if (rotationAngle == 90)
        {
          Matrix3d shiftMatrix = Matrix3d.Displacement(new Vector3d(-shiftAmount, 0, 0));
          text.TransformBy(shiftMatrix);
          ellipse.TransformBy(shiftMatrix);
        }

        tr.Commit();
      }

      ed.WriteMessage("\nPanel created.");
    }

    private double GetScaleFactor()
    {
      if (scale == "3/16") return 4.0 / 3.0;
      if (scale == "1/8") return 2;
      if (scale == "3/32") return 8.0 / 3.0;
      if (scale == "1/16") return 4;
      return 1;
    }
  }
}