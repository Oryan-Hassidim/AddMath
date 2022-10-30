using AddMath.WordAddIn.Properties;
using System.Linq;

namespace AddMath.WordAddIn
{

    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (Settings.Default.Suggestions is null)
            {
                Settings.Default.Suggestions = new();
            }
            if (Settings.Default.Suggestions.Count < 1)
            {
                var defaultCsv = Resources._default;
                var def = new SuggestionsDictionary(Csv.CsvReader.ReadFromText(@defaultCsv)
                    .ToDictionary(r => r.Values[0], r => r.Values[1]));
                Settings.Default.Suggestions.Add("Default", def);
                Settings.Default.Selected = "Default";
                Settings.Default.Save();
                Settings.Default.Reload();
            }
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
