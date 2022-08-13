using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AddMath.WordAddIn.Properties
{


    // This class allows you to handle specific events on the settings class:
    //  The SettingChanging event is raised before a setting's value is changed.
    //  The PropertyChanged event is raised after a setting's value is changed.
    //  The SettingsLoaded event is raised after the setting values are loaded.
    //  The SettingsSaving event is raised before the setting values are saved.
    internal sealed partial class Settings
    {

        public Settings()
        {
            // // To add event handlers for saving and changing settings, uncomment the lines below:
            //
            // this.SettingChanging += this.SettingChangingEventHandler;
            //
            // this.SettingsSaving += this.SettingsSavingEventHandler;
            //
        }

        public SuggestionsDictionaryCollection SuggestionsDictionary
        {
            get
            {
                if (Suggestions is null)
                {
                    Suggestions = new();
                }
                if (Suggestions.Count < 1)
                {
                    var defaultCsv = File.ReadAllText(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())), @"default.txt"));
                    var def = Csv.CsvReader.ReadFromText(@defaultCsv)
                        .ToDictionary(r => r.Values[0], r => r.Values[1]);
                    Suggestions.Add("default", def);
                    Default.Save();
                }
                return Suggestions;
            }
        }

        public Dictionary<string, string> SelectedSuggestions => SuggestionsDictionary[Selected];

        private void SettingChangingEventHandler(object sender, System.Configuration.SettingChangingEventArgs e)
        {
            // Add code to handle the SettingChangingEvent event here.
        }

        private void SettingsSavingEventHandler(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Add code to handle the SettingsSaving event here.
        }
    }
}
