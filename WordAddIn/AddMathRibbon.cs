using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word.Controls;
using Microsoft.Office.Tools.Word.Extensions;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new AddMathRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace WordAddIn
{
    [ComVisible(true)]
    public class AddMathRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private ThisAddIn addIn;

        public AddMathRibbon(ThisAddIn thisAddIn)
        {
            addIn = thisAddIn;
        }

        public Bitmap getImage(Office.IRibbonControl control)
        {
            return Properties.Resources.image;
        }

        public void OnTextButton(Office.IRibbonControl control)
        {
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Text = "This text was added by the Ribbon.";
        }
        public void OnTableButton(Office.IRibbonControl control)
        {
            //Globals.Ribbons.GetRibbon<AddMathRibbon>().
            //object missing = System.Type.Missing;
            //Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            //Word.Table newTable = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
            //currentRange, 3, 4, ref missing, ref missing);

            //// Get all of the borders except for the diagonal borders.
            //Word.Border[] borders = new Word.Border[6];
            //borders[0] = newTable.Borders[Word.WdBorderType.wdBorderLeft];
            //borders[1] = newTable.Borders[Word.WdBorderType.wdBorderRight];
            //borders[2] = newTable.Borders[Word.WdBorderType.wdBorderTop];
            //borders[3] = newTable.Borders[Word.WdBorderType.wdBorderBottom];
            //borders[4] = newTable.Borders[Word.WdBorderType.wdBorderHorizontal];
            //borders[5] = newTable.Borders[Word.WdBorderType.wdBorderVertical];

            //// Format each of the borders.
            //foreach (Word.Border border in borders)
            //{
            //    border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            //    border.Color = Word.WdColor.wdColorBlue;
            //}


            Document extendedDocument = Globals.Factory.GetVstoObject(addIn.Application.ActiveDocument);


            Button salesButton;
            extendedDocument.Paragraphs[1].Range.InsertParagraphBefore();
            //salesButton = extendedDocument.Controls.AddButton(
            //    extendedDocument.Paragraphs[1].Range, 75, 18, "salesButton");
            //salesButton.Text = "Calculate Sales";
            
            //var dd = extendedDocument.Controls.AddDropDownListContentControl(
            //    extendedDocument.Paragraphs[1].Range, "name");
            //for (int i = 0; i < 5; i++)
            //{
            //    dd.DropDownListEntries.Add($"Text {i} Text", $"{i}");
            //} 

            //var dd = extendedDocument.Controls.AddComboBox(extendedDocument.Paragraphs[1].Range, 75, 18, "name");
            //var items = new object[] { "item1", "item2" };
            //dd.Focus();
            //dd.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            
            //dd.Items.AddRange(items);
            //dd.DroppedDown = true;
            
            if (extendedDocument.Controls.Contains("name")) 
                extendedDocument.Controls.Remove("name");
            AddMathUserControl userControl = new AddMathUserControl();
            extendedDocument.Controls.AddControl(userControl, extendedDocument.Paragraphs[1].Range, 75, 200, "name");
            userControl.textBox1.Focus();
            
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAddIn.AddMathRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
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
