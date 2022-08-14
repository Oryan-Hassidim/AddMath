using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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
            // To add event handlers for saving and changing settings, uncomment the lines below:

            SettingChanging += SettingChangingEventHandler;

            SettingsSaving += SettingsSavingEventHandler;

        }

        public SuggestionsDictionaryCollection SuggestionsDictionary
        {
            get
            {
                if (Suggestions is null)
                {
                    var s = Application.LocalUserAppDataPath;
                    Suggestions = new();
                }
                if (Suggestions.Count < 1)
                {
                    var defaultCsv = File.ReadAllText(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())), @"default.txt"));
                    var def = new SuggestionsDictionary(Csv.CsvReader.ReadFromText(@defaultCsv)
                        .ToDictionary(r => r.Values[0], r => r.Values[1]));
                    Suggestions.Add("default", def);
                    Save();
                }
                return Suggestions;
            }
        }

        public Dictionary<string, string> SelectedSuggestions => SuggestionsDictionary[Selected];
        public event EventHandler<Dictionary<string, string>> SelectedSuggestionsChanged;

        private void SettingChangingEventHandler(object sender, System.Configuration.SettingChangingEventArgs e)
        {
            // Add code to handle the SettingChangingEvent event here.
        }

        protected override void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            base.OnPropertyChanged(sender, e);
            if (e.PropertyName == nameof(Selected))
            {
                SelectedSuggestionsChanged?.Invoke(this, SelectedSuggestions);
            }
        }

        private void SettingsSavingEventHandler(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Add code to handle the SettingsSaving event here.
        }
    }
}
