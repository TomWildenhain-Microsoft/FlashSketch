﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PowerPointAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
            
        }
        public void OnNewArtboard(Office.IRibbonControl control)
        {
            FlashSketch.Instance.NewArtboard();
        }
        public void OnEqualizeWidths(Office.IRibbonControl control)
        {
            FlashSketch.Instance.EqualizeWidths();
        }

        public void OnEqualizeHeights(Office.IRibbonControl control)
        {
            FlashSketch.Instance.EqualizeHeights();
        }

        public void OnMakeSquare(Office.IRibbonControl control)
        {
            FlashSketch.Instance.MakeSquare();
        }

        public void OnDistributeHorizontally(Office.IRibbonControl control)
        {
            FlashSketch.Instance.DistributeHorizontally();
        }

        public void OnDistributeVertically(Office.IRibbonControl control)
        {
            FlashSketch.Instance.DistributeVertically();
        }

        public void OnEqualizeLineLengths(Office.IRibbonControl control)
        {
            FlashSketch.Instance.EqualizeLineLengths();
        }

        public void OnRecompute(Office.IRibbonControl control)
        {
            FlashSketch.Instance.RecomputeConstraints();
        }

        public void OnLockConstraints(Office.IRibbonControl control, bool pressed)
        {
            FlashSketch.Instance.LockContraints = pressed;
        }

        public bool GetPressedLockConstraints(Office.IRibbonControl control)
        {
            return FlashSketch.Instance.LockContraints;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointAddIn1.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            FlashSketch.Instance.Ribbon = ribbon;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
