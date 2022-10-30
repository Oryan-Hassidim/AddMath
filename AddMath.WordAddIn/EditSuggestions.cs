using AddMath.WordAddIn.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace AddMath.WordAddIn
{
    public partial class EditSuggestions : Form
    {
        private readonly BindingList<KeyValue<string, SuggestionsDictionary>> SuggestionsDictionaries = new();
        private readonly SortableList<KeyValue<string, string>> SelectedDictionary = new();
        private KeyValue<string, SuggestionsDictionary> SelectedItem => ThemeComboBox.SelectedItem as KeyValue<string, SuggestionsDictionary>;
        private int lastIndex = 0;
        public EditSuggestions()
        {
            InitializeComponent();
        }

        private void EditSuggestions_Load(object sender, EventArgs e)
        {
            ThemeComboBox.DataSource = SuggestionsDictionaries;
            var bindingSource = new BindingSource(SelectedDictionary, null)
            {
                AllowNew = true
            };
            GridView.DataSource = bindingSource;
            foreach (var kv in Settings.Default.SuggestionsDictionary)
            {
                SuggestionsDictionaries.Add(kv);
            }
            ThemeComboBox.SelectedIndex = 0;
            foreach (var item in SelectedItem.Value)
            {
                SelectedDictionary.Add(item);
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            UpdateLastSelecterd();
            var suggestions = Settings.Default.SuggestionsDictionary;
            suggestions.Clear();
            foreach (var item in SuggestionsDictionaries)
            {
                suggestions.Add(item.Key, item.Value);
            }
            if (!suggestions.ContainsKey(Settings.Default.Selected))
            {
                Settings.Default.Selected = suggestions.Keys.First();
            }
            Settings.Default.Save();
            Close();
        }

        private void UpdateLastSelecterd()
        {
            if (lastIndex < 0 || lastIndex >= SuggestionsDictionaries.Count)
                return;
            var lastSelected = SuggestionsDictionaries[lastIndex].Value;
            lastSelected.Clear();
            foreach (var item in SelectedDictionary)
            {
                lastSelected[item.Key] = item.Value;
            }
        }

        private void ThemeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateLastSelecterd();
            UpdateGridView();
        }

        private void UpdateGridView()
        {
            SelectedDictionary.Clear();
            foreach (var item in SelectedItem.Value)
            {
                SelectedDictionary.Add(item);
            }
            lastIndex = ThemeComboBox.SelectedIndex;
        }

        private void DeleteTheme_Click(object sender, EventArgs e)
        {
            if (SuggestionsDictionaries.Count == 1)
                SuggestionsDictionaries.Add(new("Default", new()));
            SuggestionsDictionaries.RemoveAt(ThemeComboBox.SelectedIndex);
            UpdateGridView();
        }

        private void AddTheme_Click(object sender, EventArgs e)
        {
            UpdateLastSelecterd();
            AddThemeForm addThemeForm = new();
            addThemeForm.NameTextBox.Text = $"Theme {SuggestionsDictionaries.Count + 1}";
            addThemeForm.BasedOnComboBox.Items.Add(new KeyValue<string, SuggestionsDictionary>("None", new()));
            addThemeForm.BasedOnComboBox.Items.AddRange(
                SuggestionsDictionaries
                .Cast<object>()
                .ToArray());
            addThemeForm.BasedOnComboBox.SelectedIndex = 0;
            if (addThemeForm.ShowDialog() == DialogResult.OK)
            {
                var kv = addThemeForm.BasedOnComboBox.SelectedItem as KeyValue<string, SuggestionsDictionary>;
                SuggestionsDictionaries.Add(new(addThemeForm.NameTextBox.Text, new(kv.Value)));
                ThemeComboBox.SelectedItem = SuggestionsDictionaries.Last();
            }
            addThemeForm.Dispose();
        }
    }
    public class KeyValue<TKey, TValue>
    {
        public TKey Key { get; set; }
        public TValue Value { get; set; }
        public KeyValue() { }
        public KeyValue(TKey key, TValue value)
        {
            Key = key;
            Value = value;
        }
        public static implicit operator KeyValuePair<TKey, TValue>(KeyValue<TKey, TValue> kv) => new(kv.Key, kv.Value);
        public static implicit operator KeyValue<TKey, TValue>(KeyValuePair<TKey, TValue> kv) => new(kv.Key, kv.Value);
    }

}
