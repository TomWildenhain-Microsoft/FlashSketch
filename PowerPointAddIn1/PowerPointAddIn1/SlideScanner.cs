﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.ComponentModel;

namespace PowerPointAddIn1
{
    class SlideScanner
    {
        public static SlideScanner Instance = new SlideScanner();
        public Dictionary<int, SlideObject> LastScan = null;
        public SlideScanner()
        {

        }

        public Dictionary<int, SlideObject> ScanSlide(PowerPoint.Slide slide)
        {
            Dictionary<int, SlideObject> res = new Dictionary<int, SlideObject>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                ScanShape(shape, res, "");
            }
            LastScan = res;
            return res;
        }

        void ScanShape(PowerPoint.Shape shape, Dictionary<int, SlideObject> res, string prefix)
        {
            res[shape.Id] = new SlideObject(shape.Left, shape.Top, shape.Width, shape.Height, shape.Id, shape.Name, shape);
            //FlashSketch.Instance.PrintToNotes(prefix + shape.Name + " " + shape.Left + " " + shape.Top + " " + shape.Width + " " + shape.Height);
            if (shape.Type == MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape child in shape.GroupItems)
                {
                    ScanShape(child, res, prefix + shape.Name + " : ");
                }
            }
        }

        public void ApplyDimsToShapes(Dictionary<int, SlideObject> dimensions, Dictionary<CasVar, float> values)
        {
            foreach (SlideObject obj in dimensions.Values)
            {
                obj.Shape.Width = CasSystem.Instance.Eval(obj.WidthExpr, values);
                obj.Shape.Height = CasSystem.Instance.Eval(obj.HeightExpr, values);
                obj.Shape.Top = CasSystem.Instance.Eval(obj.YExpr, values);
                obj.Shape.Left = CasSystem.Instance.Eval(obj.XExpr, values);
            }
        }
    }

    class SlideObject
    {
        public float X1;
        public float Y1;
        public float Width;
        public float Height;
        public float X2;
        public float Y2;
        public float CY;
        public float CX;
        public int ShapeId;
        public string ShapeName;
        public CasExpr XExpr = null;
        public CasExpr YExpr = null;
        public CasExpr WidthExpr = null;
        public CasExpr HeightExpr = null;
        public PowerPoint.Shape Shape;
        public SlideObject(double x, double y, double width, double height, int id, string shapeName, PowerPoint.Shape shape)
        {
            X1 = (float)x;
            Y1 = (float)y;
            Width = (float)width;
            Height = (float)height;
            X2 = (float)(x + width);
            Y2 = (float)(y + height);
            CX = (float)(x + width / 2);
            CY = (float)(y + height / 2);
            ShapeId = id;
            Shape = shape;
            ShapeName = shapeName;
        }
        public CasExpr LongerDimExpr()
        {
            if (Width > Height) return WidthExpr;
            return HeightExpr;
        }
        public CasExpr ShorterDimExpr()
        {
            if (Width < Height) return WidthExpr;
            return HeightExpr;
        }
    }
}
