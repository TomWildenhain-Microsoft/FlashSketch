using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.ComponentModel.Design;

namespace PowerPointAddIn1
{
    class FlashSketch
    {
        public static FlashSketch Instance = new FlashSketch();
        public PowerPoint.Presentation Pres = null;
        public PowerPoint.Application Application = null;
        public PowerPoint.SlideRange SlideSelection = null;
        public PowerPoint.Selection Selection = null;

        public FlashSketch()
        {

        }

        public void NewArtboard()
        {
            if (SlideSelection == null || SlideSelection.Count != 1) return;
            var slide = SlideSelection[1];
            Application.StartNewUndoEntry();
            PowerPoint.Shape artboard = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 25, 10, 50, 50);
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
            artboard.TextFrame.MarginTop = 66.8f;
            artboard.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Black.ToArgb();
            artboard.Name = "Artboard " + (artboard.Id-1);
            artboard.Select();
        }

        internal void DistributeHorizontally()
        {
            throw new NotImplementedException();
        }

        internal void DistributeVertically()
        {
            throw new NotImplementedException();
        }

        internal void MakeSquare()
        {
            if (Selection == null || Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Selection.ShapeRange.Count != 1)
            {
                return;
            }
            int id = Selection.ShapeRange[1].Id;
            SnapDetector.Instance.MakeSquare(SlideScanner.Instance.ScanSlide(SlideSelection[1]), id);
        }

        internal void EqualizeHeights()
        {
            throw new NotImplementedException();
        }

        internal void EqualizeWidths()
        {
            throw new NotImplementedException();
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
            int id1 = Selection.ShapeRange[1].Id;
            int id2 = Selection.ShapeRange[2].Id;
            SnapDetector.Instance.EqualizeLongerDims(SlideScanner.Instance.ScanSlide(SlideSelection[1]), id1, id2);
        }
    }
}
