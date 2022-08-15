using AddMath.WordAddIn.Properties;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddMath.WordAddIn
{
    public partial class Suggestions : Form
    {
        private readonly Dictionary<string, int> newWords = new();
        private readonly Dictionary<Suggestion, string> _Suggestions = new();
        private readonly SortedSet<Suggestion> startsWith = new();
        private readonly SortedSet<Suggestion> contains = new();
        private readonly StringBuilder matrix = new();
        private BindingList<Suggestion> LiveSuggestions { get; set; } = new();
        private Microsoft.Office.Interop.Word.Application instance => Globals.ThisAddIn.Application;

        public string Theme { get; set; }

        private void loadSuggestions()
        {
            _Suggestions.Clear();
            Settings.Default.Reload();
            foreach (var kv in Settings.Default.SelectedSuggestions)
            {
                _Suggestions.Add(new() { Text = kv.Key, Type = SuggestionType.Text }, kv.Value);
            }
            SuggestionsList.BeginUpdate();
            LiveSuggestions.Clear();
            foreach (var item in _Suggestions.Keys.OrderBy(s => s.Text))
            {
                LiveSuggestions.Add(item);
            }
            SuggestionsList.EndUpdate();
        }

        #region Suggestions
        public Suggestions()
        {
            InitializeComponent();
            SuggestionsList.DataSource = LiveSuggestions;
        }
        private async void Suggestions_Load(object sender, EventArgs e)
        {
            loadSuggestions();
            Settings.Default.SelectedSuggestionsChanged += (s, e) => loadSuggestions();
            Settings.Default.SettingsSaving += (s, e) => loadSuggestions();
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            e.Cancel = true;
            Hide();
        }
        private void Suggestions_Activated(object sender, EventArgs e)
        {
            instance.ActiveWindow.GetPoint(out int left, out int top, out _, out _, instance.Selection.Range);
            var zoom = instance.ActiveWindow.View.Zoom.Percentage / 100f;

            Font = new Font(Font.FontFamily, instance.Selection.Font.Size * zoom * 0.9f, Font.Style, GraphicsUnit.Point, Font.GdiCharSet, Font.GdiVerticalFont);
            Top = top - 5;
            Left = left - 5;
            SearchTextBox.Clear();
            SearchTextBox.Focus();
            if (SuggestionsList.Items.Count > 0)
            {
                SuggestionsList.SelectedIndex = 0;
            }
        }
        private void Suggestions_Deactivate(object sender, EventArgs e)
        {
            Hide();
        }

        private void Suggestions_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                case Keys.Tab:
                    if (LiveSuggestions.Any())
                    {
                        if (SuggestionsList.SelectedIndex < 0)
                            SuggestionsList.SelectedIndex = 0;
                        AddSelectedSuggestionFromList();
                    }
                    else
                    {
                        var text = SearchTextBox.Text;
                        if (char.IsDigit(text, 0) && char.IsDigit(text, 1))
                        {
                            AddMatrix();
                        }
                        else
                        {
                            AddSelectedSuggestionFromList();
                        }
                    }
                    break;
                case Keys.Escape:
                    Hide();
                    break;
                default:
                    break;
            }
        }
        #endregion

        #region SearchTextBox
        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            var t = SearchTextBox.Text;
            SuggestionsList.BeginUpdate();
            LiveSuggestions.Clear();
            startsWith.Clear();
            contains.Clear();

            Suggestion key;
            foreach (var item in _Suggestions)
            {
                key = item.Key;
                switch (item.Key.Text.IndexOf(t))
                {
                    case 0:
                        startsWith.Add(key);
                        break;
                    case -1:
                        break;
                    default:
                        contains.Add(key);
                        break;
                }
            }
            foreach (var item in startsWith.Union(contains))
            {
                LiveSuggestions.Add(item);
            }

            var s = t.AsSpan();
            switch (t.Length)
            {
                case 1 when int.TryParse(t, out int height):
                    LiveSuggestions.Add(new()
                    {
                        Text = $"Matrix {height}×1",
                        Type = SuggestionType.Matrix
                    });
                    break;
                case 2 when int.TryParse(t, out int value):
                    int h = value / 10, w = value % 10;
                    LiveSuggestions.Add(new()
                    {
                        Text = $"Matrix {h}×{w}",
                        Type = SuggestionType.Matrix
                    });
                    break;
                case 3 when int.TryParse(s.Slice(0, 2).ToString(), out int value) && s[2] == 'b':
                    h = value / 10;
                    w = value % 10;
                    LiveSuggestions.Add(new()
                    {
                        Text = $"Matrix {h}×{w}|{h}×1",
                        Type = SuggestionType.Matrix
                    });
                    break;
            }
            SuggestionsList.EndUpdate();
        }
        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                case Keys.Tab:
                    if (LiveSuggestions.Any())
                    {
                        SuggestionsList.SelectedIndex = 0;
                        AddSelectedSuggestionFromList();
                    }
                    else
                    {
                        var text = SearchTextBox.Text;
                        if (char.IsDigit(text, 0) && char.IsDigit(text, 1))
                        {
                            AddMatrix();
                        }
                        else
                        {
                            AddSelectedSuggestionFromList();
                        }
                    }
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                case Keys.Down:
                    if (SuggestionsList.SelectedIndex == -1)
                        SuggestionsList.SelectedIndex = 0;
                    SuggestionsList.Focus();
                    if (SuggestionsList.SelectedIndex < SuggestionsList.Items.Count - 1)
                        SuggestionsList.SelectedIndex++;
                    break;
                case Keys.Escape:
                case Keys.Back when SearchTextBox.Text.Length == 0:
                    Hide();
                    SearchTextBox.Clear();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                default:
                    e.Handled = false;
                    break;
            }
        }
        #endregion

        #region SuggestionsList
        private void SuggestionsList_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Up when SuggestionsList.SelectedIndex == 0:
                    SearchTextBox.Focus();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                case Keys.Escape:
                case Keys.Back:
                    Hide();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
                default:
                    break;
            }
        }
        private void SuggestionsList_DoubleClick(object sender, EventArgs e)
        {
            AddSelectedSuggestionFromList();
        }
        #endregion


        public void AddMatrix()
        {
            SearchTextBox.Focus();
            matrix.Clear();
            matrix.Append("■(");

            var t = SearchTextBox.Text;
            var s = t.AsSpan();
            switch (t.Length)
            {
                case 1 when int.TryParse(t, out int height):
                    matrix.AppendJoin("@", Enumerable.Repeat("", height));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    Hide();
                    SendKeys.SendWait(" ");
                    break;
                case 2 when int.TryParse(t, out int value):
                    int h = value / 10, w = value % 10 - 1;
                    var row = string.Concat(Enumerable.Repeat("&", w));
                    matrix.AppendJoin("@", Enumerable.Repeat(row, h));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    Hide();
                    SendKeys.SendWait(" ");
                    break;
                case 3 when int.TryParse(s.Slice(0, 2).ToString(), out int value) && s[2] == 'b':
                    h = value / 10;
                    w = value % 10 - 1;
                    row = string.Concat(Enumerable.Repeat("&", w));
                    matrix.AppendJoin("@", Enumerable.Repeat(row, h));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    Hide();
                    SendKeys.SendWait(" ");

                    matrix.Clear();
                    matrix.Append("|■(");
                    matrix.AppendJoin("@", Enumerable.Repeat("", h));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    SendKeys.SendWait(" ");
                    break;
            }

            SearchTextBox.Clear();
        }
        public async void AddSelectedSuggestionFromList()
        {
            Hide();

            if (SuggestionsList.SelectedIndex == -1 && LiveSuggestions.Any())
                SuggestionsList.SelectedIndex = 0;
            else if (!LiveSuggestions.Any())
            {
                var t = SearchTextBox.Text;
                if (newWords.ContainsKey(t))
                {
                    newWords[t]++;
                    if (newWords[t] >= 2)
                    {
                        var r = MessageBox.Show($"Do you want add \"{t}\" to your suggestions?", "Add new word",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                            MessageBoxDefaultButton.Button1);
                        if (r == DialogResult.Yes)
                        {
                            Settings.Default.SelectedSuggestions.Add(t, t);
                            Settings.Default.Save();
                            loadSuggestions();
                        }
                    }
                }
                else newWords.Add(t, 1);

                instance.Selection.TypeText(@"\" + t);
                SendKeys.SendWait(" ");
                SearchTextBox.Clear();
                return;
            }


            if (SuggestionsList.SelectedValue != null)
            {
                var s = SuggestionsList.SelectedValue as Suggestion;
                switch (s.Type)
                {
                    case SuggestionType.Text:
                        var suggestion = _Suggestions[s];
                        var trimed = suggestion.TrimEnd();
                        instance.Selection.TypeText(trimed);
                        int spaces = suggestion.Length - trimed.Length;
                        for (int i = 0; i < spaces; i++)
                        {
                            SendKeys.SendWait(" ");
                        }
                        SearchTextBox.Clear();
                        break;
                    case SuggestionType.Matrix:
                        AddMatrix();
                        break;
                    default:
                        break;
                }
            }
        }


        private void OKButton_Click(object sender, EventArgs e)
        {
            if (SearchTextBox.Focused)
                SearchTextBox_KeyDown(this, new KeyEventArgs(Keys.Enter));
            else
                Suggestions_KeyDown(this, new KeyEventArgs(Keys.Enter));
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            SearchTextBox_KeyDown(this, new KeyEventArgs(Keys.Escape));
        }

    }
}
