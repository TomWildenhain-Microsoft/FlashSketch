using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            FlashSketch.Instance.Application = this.Application;
            this.Application.PresentationNewSlide += new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
            this.Application.AfterShapeSizeChange += Application_AfterShapeSizeChange;
            this.Application.AfterPresentationOpen += Application_AfterPresentationOpen;
            this.Application.AfterNewPresentation += Application_AfterNewPresentation;
            this.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            FlashSketch.Instance.Selection = Sel;
            FlashSketch.Instance.SoftRecomputeConstraints();
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            FlashSketch.Instance.SlideSelection = SldRange;
        }

        private void Application_AfterNewPresentation(PowerPoint.Presentation Pres)
        {
            FlashSketch.Instance.Pres = Pres;
        }

        private void Application_AfterPresentationOpen(PowerPoint.Presentation Pres)
        {
            FlashSketch.Instance.Pres = Pres;
        }

        private void Application_AfterShapeSizeChange(PowerPoint.Shape shp)
        {
            //if (shp.Name.StartsWith("Artboard "))
            //{
            //    Application.StartNewUndoEntry();
            //    shp.TextFrame.MarginTop = shp.Height + 16.8f;
            //    FlashSketch.Instance.ResizeArtboard(shp);
            //}
            FlashSketch.Instance.ShapeResize(shp);
        }

        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            /*PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");*/
            //MessageBox.Show("New slide!");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
