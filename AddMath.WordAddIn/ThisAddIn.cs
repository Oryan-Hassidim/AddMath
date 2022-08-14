using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Runtime.InteropServices;

namespace AddMath.WordAddIn
{

    [ComVisible(true)]
    public interface IAddMathObject
    {
        void AddMath();
    }


    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddMathObject : IAddMathObject, IDisposable
    {
        private Suggestions _suggestions;
        private Suggestions Suggestions => _suggestions ??= new Suggestions();
        public string Theme
        {
            get => _suggestions.Theme;
            set => _suggestions.Theme = value;
        }
        public void AddMath()
        {
            Suggestions.Show();
            Suggestions.Activate();
        }

        public void Dispose()
        {
            _suggestions?.Dispose();
        }
    }
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private AddMathObject addMathObject = new AddMathObject();
        public AddMathObject AddMathObject => addMathObject;

        protected override object RequestComAddInAutomationService()
        {
            return AddMathObject;
        }
        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            addMathObject.Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
