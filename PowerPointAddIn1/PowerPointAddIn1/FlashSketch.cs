﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.ComponentModel.Design;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    class FlashSketch
    {
        public static FlashSketch Instance = new FlashSketch();
        public PowerPoint.Presentation Pres = null;
        public PowerPoint.Application Application = null;
        public PowerPoint.SlideRange SlideSelection = null;
        public PowerPoint.Selection Selection = null;
        public bool LockContraints = false;
        public Office.IRibbonUI Ribbon;

        public FlashSketch()
        {

        }

        public void NewArtboard()
        {
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            var slide = SlideSelection[1];
            Application.StartNewUndoEntry();
            PowerPoint.Shape artboard = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 25, 10, 50, 60);
            artboard.TextFrame.TextRange.InsertAfter("Artboard");
            artboard.Fill.ForeColor.RGB = System.Drawing.Color.White.ToArgb();
            artboard.Line.Visible = Office.MsoTriState.msoFalse;
            artboard.Shadow.Type = Office.MsoShadowType.msoShadow25;
            artboard.Shadow.Blur = 15;
            artboard.Shadow.Transparency = 0.6f;
            artboard.Shadow.Size = 100;
            artboard.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
            artboard.TextEffect.FontSize = 18;
            artboard.TextEffect.Alignment = Office.MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            artboard.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
            artboard.TextFrame.MarginBottom = 0;
            artboard.TextFrame.MarginLeft = 0;
            artboard.TextFrame.MarginRight = 0;
            artboard.TextFrame.MarginTop = 76.8f;
            artboard.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Black.ToArgb();
            artboard.Name = "Artboard " + (artboard.Id-1);
            artboard.Select();
        }


        public void ShapeResize(PowerPoint.Shape shp)
        {
            if (LockContraints && SlideScanner.Instance.LastScan != null)
            {
                float newWidth = shp.Width;
                float newHeight = shp.Height;
                var LastScan = SlideScanner.Instance.LastScan;
                var prevWKey = SnapDetector.Instance.FloatToKey(LastScan[shp.Id].Width);
                var prevHKey = SnapDetector.Instance.FloatToKey(LastScan[shp.Id].Height);
                var currWKey = SnapDetector.Instance.FloatToKey(shp.Width);
                var currHKey = SnapDetector.Instance.FloatToKey(shp.Height);
                shp.Width = LastScan[shp.Id].Width;
                shp.Height = LastScan[shp.Id].Height;
                Application.StartNewUndoEntry();
                if (prevWKey != currWKey)
                {
                    SnapDetector.Instance.SetShapeWidth(SlideScanner.Instance.ScanSlide(SlideSelection[1]), shp.Id, newWidth);
                }
                else
                {
                    SnapDetector.Instance.SetShapeHeight(SlideScanner.Instance.ScanSlide(SlideSelection[1]), shp.Id, newHeight);
                }
                LockContraints = false;
                SlideScanner.Instance.ScanSlide(SlideSelection[1]);
                Ribbon.Invalidate();
            }
            else if (LockContraints)
            {
                LockContraints = false;
                Ribbon.Invalidate();
            }
            else
            {
                if (SlideSelection == null || SlideSelection.Count != 1) return;
                FlashSketch.Instance.Application.StartNewUndoEntry();
                SnapDetector.Instance.UpdateSnapCache(SlideScanner.Instance.ScanSlide(SlideSelection[1]));
            }
        }

        internal void ResizeArtboard(PowerPoint.Shape shp)
        {
            Application.CommandBars.ExecuteMso("Undo");
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            //if (SlideScanner.Instance.LastScan == null) return;
            //SnapDetector.Instance.ResizeShape(SlideScanner.Instance.LastScan, shp.Id, shp.Width, shp.Height);
            //SlideScanner.Instance.ScanSlide(SlideSelection[1]);
        }

        internal void DistributeHorizontally()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count < 3)
            {
                return;
            }
            FlashSketch.Instance.Application.StartNewUndoEntry();
            Selection.ShapeRange.Distribute(MsoDistributeCmd.msoDistributeHorizontally, MsoTriState.msoFalse);
            List<Tuple<int, float>> shapeIds = new List<Tuple<int, float>>();
            foreach (PowerPoint.Shape shape in Selection.ShapeRange)
            {
                shapeIds.Add(new Tuple<int, float>(shape.Id, shape.Left));
            }
            shapeIds.Sort((s1, s2) => s1.Item2.CompareTo(s2.Item2));
            var scan = SlideScanner.Instance.ScanSlide(SlideSelection[1]);
            SnapDetector.Instance.UpdateSnapCacheAfterHDist(scan, shapeIds);
        }

        public void SoftRecomputeConstraints()
        {
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            var objs = SlideScanner.Instance.ScanSlide(SlideSelection[1]);
            SnapDetector.Instance.UpdateSnapCache(objs);
        }

        internal void RecomputeConstraints()
        {
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            var objs = SlideScanner.Instance.ScanSlide(SlideSelection[1]);
            SnapDetector.Instance.UpdateSnapCache(objs);
            FlashSketch.Instance.Application.StartNewUndoEntry();
            SnapDetector.Instance.PrintConstraints(objs);
        }

        internal void DistributeVertically()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count < 3)
            {
                return;
            }
            FlashSketch.Instance.Application.StartNewUndoEntry();
            Selection.ShapeRange.Distribute(MsoDistributeCmd.msoDistributeVertically, MsoTriState.msoFalse);
            List<Tuple<int, float>> shapeIds = new List<Tuple<int, float>>();
            foreach (PowerPoint.Shape shape in Selection.ShapeRange)
            {
                shapeIds.Add(new Tuple<int, float>(shape.Id, shape.Top));
            }
            shapeIds.Sort((s1, s2) => s1.Item2.CompareTo(s2.Item2));
            var scan = SlideScanner.Instance.ScanSlide(SlideSelection[1]);
            SnapDetector.Instance.UpdateSnapCacheAfterVDist(scan, shapeIds);
        }

        internal void MakeSquare()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count != 1)
            {
                return;
            }
            FlashSketch.Instance.Application.StartNewUndoEntry();
            int id = Selection.ShapeRange[1].Id;
            SnapDetector.Instance.MakeSquare(SlideScanner.Instance.ScanSlide(SlideSelection[1]), id);
        }

        internal void EqualizeHeights()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count != 2)
            {
                return;
            }
            FlashSketch.Instance.Application.StartNewUndoEntry();
            int id1 = Selection.ShapeRange[1].Id;
            int id2 = Selection.ShapeRange[2].Id;
            SnapDetector.Instance.EqualizeHeights(SlideScanner.Instance.ScanSlide(SlideSelection[1]), id1, id2);
        }

        internal void EqualizeWidths()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count != 2)
            {
                return;
            }
            FlashSketch.Instance.Application.StartNewUndoEntry();
            int id1 = Selection.ShapeRange[1].Id;
            int id2 = Selection.ShapeRange[2].Id;
            SnapDetector.Instance.EqualizeWidths(SlideScanner.Instance.ScanSlide(SlideSelection[1]), id1, id2);
        }

        public void PrintToNotes(string text)
        {
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            var slide = SlideSelection[1];
            foreach (PowerPoint.Shape shape in slide.NotesPage.Shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    if (shape.Name.StartsWith("Notes"))
                        shape.TextFrame.TextRange.Text += text + "\n";
            }
        }

        public void ClearNotes()
        {
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            var slide = SlideSelection[1];
            foreach (PowerPoint.Shape shape in slide.NotesPage.Shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    if (shape.Name.StartsWith("Notes"))
                        shape.TextFrame.TextRange.Text = "";
            }
        }


        public void EqualizeLineLengths()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count != 2)
            {
                return;
            }
            FlashSketch.Instance.Application.StartNewUndoEntry();
            int id1 = Selection.ShapeRange[1].Id;
            int id2 = Selection.ShapeRange[2].Id;
            SnapDetector.Instance.EqualizeLongerDims(SlideScanner.Instance.ScanSlide(SlideSelection[1]), id1, id2);
        }
    }
}
