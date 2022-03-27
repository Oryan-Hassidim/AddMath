using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Collections.ObjectModel;
using System.IO;
using System.ComponentModel;
using System.Globalization;
using Forms = System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using CommunityToolkit.Mvvm;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Office.Core;

namespace AddMath
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        private Dictionary<string, int> newWords = new();
        private Dictionary<Suggestion, string> Suggestions = new();
        private SortedSet<Suggestion> startsWith = new(), contains = new();
        private StringBuilder matrix = new();
        private string file;
        private double factor = 1;

        public ObservableCollection<Suggestion> LiveSuggestions { get; set; } = new();

        private Word.Application instance;
        private Process p;

        public async Task loadSuggestions()
        {
            await Task.Run(() =>
            {
                Suggestions.Clear();
                foreach (var item in File.ReadAllLines(file))
                {
                    var a = item.Split("~~~");
                    Suggestions.Add(new() { Text = a[0], Type = SuggestionType.Text }, a[1]);
                }
            });
        }
        public void ActivateMe()
        {
            SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
        }
        public void ActivateWord()
        {
            p.Refresh();
            IntPtr h = p.MainWindowHandle;
            SetForegroundWindow(h);
        }
        private bool FindRunningWord()
        {
            try
            {
                instance = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            }
            catch (COMException)
            {
                MessageBox.Show("There is't Word Running.");
                Close();
                return false;
            }

            p = Process.GetProcessesByName("WinWord").FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);
            if (p is null)
            {
                MessageBox.Show("There is't Word Running.");
                Close();
                return false;
            }
            return true;
        }


        public MainWindow()
        {
            instance = null;
            if (!FindRunningWord()) return;
            factor = Forms.Screen.PrimaryScreen.Bounds.Width / (double)SystemParameters.PrimaryScreenWidth;

            InitializeComponent();
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var t = SearchBox.Text;

            LiveSuggestions.Clear();
            startsWith.Clear();
            contains.Clear();

            Suggestion key;
            foreach (var item in Suggestions)
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
                case 3 when int.TryParse(s.Slice(0, 2), out int value) && s[2] == 'b':
                    h = value / 10;
                    w = value % 10;
                    LiveSuggestions.Add(new()
                    {
                        Text = $"Matrix {h}×{w}|{h}×1",
                        Type = SuggestionType.Matrix
                    });
                    break;
            }
        }

        private void SearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                case Key.Tab:
                    if (LiveSuggestions.Any())
                    {
                        SuggestionsList.SelectedIndex = 0;
                        AddSelectedSuggestionFromList();
                    }
                    else
                    {
                        var text = SearchBox.Text;
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
                case Key.Down:
                    if (SuggestionsList.SelectedIndex == -1)
                        SuggestionsList.SelectedIndex = 0;
                    var listBoxItem = (ListViewItem)SuggestionsList
                        .ItemContainerGenerator
                        .ContainerFromItem(SuggestionsList.SelectedItem);
                    e.Handled = true;
                    listBoxItem.Focus();
                    break;
                case Key.Escape when Keyboard.IsKeyDown(Key.LeftShift):
                    Close();
                    break;
                case Key.Escape:
                case Key.Back when SearchBox.Text.Length == 0:
                    SearchBox.Clear();
                    ActivateWord();
                    break;
                default:
                    e.Handled = false;
                    break;
            }
        }

        public void AddMatrix()
        {
            if (!FindRunningWord()) return;
            SearchBox.Focus();
            matrix.Clear();
            matrix.Append("■(");

            var t = SearchBox.Text;
            var s = t.AsSpan();
            switch (t.Length)
            {
                case 1 when int.TryParse(t, out int height):
                    matrix.AppendJoin('@', Enumerable.Repeat("", height));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    ActivateWord();
                    Forms.SendKeys.SendWait(" ");
                    break;
                case 2 when int.TryParse(t, out int value):
                    int h = value / 10, w = value % 10 - 1;
                    var row = string.Concat(Enumerable.Repeat("&", w));
                    matrix.AppendJoin('@', Enumerable.Repeat(row, h));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    ActivateWord();
                    Forms.SendKeys.SendWait(" ");
                    break;
                case 3 when int.TryParse(s.Slice(0, 2), out int value) && s[2] == 'b':
                    h = value / 10;
                    w = value % 10 - 1;
                    row = string.Concat(Enumerable.Repeat("&", w));
                    matrix.AppendJoin('@', Enumerable.Repeat(row, h));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    ActivateWord();
                    Forms.SendKeys.SendWait(" ");

                    matrix.Clear();
                    matrix.Append("|■(");
                    matrix.AppendJoin('@', Enumerable.Repeat("", h));
                    matrix.Append(')');
                    instance.Selection.TypeText(matrix.ToString());
                    Forms.SendKeys.SendWait(" ");
                    break;
            }

            SearchBox.Clear();
        }

        private void SuggestionsList_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                case Key.Tab:
                    AddSelectedSuggestionFromList();
                    break;
                case Key.Up:
                    if (SuggestionsList.SelectedIndex == 0)
                    {
                        SearchBox.Focus();
                    }
                    break;
            }
        }

        public void AddSelectedSuggestionFromList()
        {
            if (!FindRunningWord()) return;
            SearchBox.Focus();

            if (SuggestionsList.SelectedIndex == -1 && LiveSuggestions.Any())
                SuggestionsList.SelectedIndex = 0;
            else if (!LiveSuggestions.Any())
            {
                var t = SearchBox.Text;
                if (newWords.ContainsKey(t))
                {
                    newWords[t]++;
                    if (newWords[t] >= 2)
                    {
                        var r = MessageBox.Show($"Do you want add \"{t}\" to your suggestions?", "Add new word",
                            MessageBoxButton.YesNo, MessageBoxImage.Question,
                            MessageBoxResult.Yes, MessageBoxOptions.None);
                        if (r == MessageBoxResult.Yes)
                        {
                            File.AppendAllText(file, $"\n{t}~~~\\{t} ");
                            loadSuggestions();
                        }
                    }
                }
                else newWords.Add(t, 1);

                instance.Selection.TypeText("\\" + t);
                ActivateWord();
                Forms.SendKeys.SendWait(" ");
                SearchBox.Clear();
                return;
            }


            if (SuggestionsList.SelectedValue != null)
            {
                var s = (Suggestion)SuggestionsList.SelectedValue;
                switch (s.Type)
                {
                    case SuggestionType.Text:
                        var suggestion = Suggestions[s];
                        var trimed = suggestion.TrimEnd();
                        instance.Selection.TypeText(trimed);
                        int spaces = suggestion.Length - trimed.Length;
                        ActivateWord();
                        for (int i = 0; i < spaces; i++)
                        {
                            Forms.SendKeys.SendWait(" ");
                        }
                        SearchBox.Clear();
                        break;
                    case SuggestionType.Matrix:
                        AddMatrix();
                        break;
                    default:
                        break;
                }
            }
        }

        private void Me_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        [ICommand]
        private async void RefreshSuggestions()
        {
            await loadSuggestions();
            SearchBox.Focus();
        }

        [ICommand]
        private void OpenSuggestionsFile()
        {
            new Process
            {
                StartInfo = new ProcessStartInfo(file)
                {
                    UseShellExecute = true
                }
            }.Start();
        }

        private void Me_Activated(object sender, EventArgs e)
        {
            if (!FindRunningWord()) return;
            instance.ActiveWindow.GetPoint(out int left, out int top, out int width, out int height, instance.Selection.Range);
            var zoom = instance.ActiveWindow.View.Zoom.Percentage / 100.0;
            FontSize = instance.Selection.Font.Size * zoom;
            var rowHeight = height - (instance.Selection.ParagraphFormat.SpaceAfter + instance.Selection.ParagraphFormat.SpaceBefore) * zoom * factor * 1.33;

            var ttop = top + instance.Selection.ParagraphFormat.SpaceBefore * zoom * factor * 1.33 + rowHeight * factor * 0.7 / 2 - FontSize * 1.33 / 2;
            Top = (5 + ttop) / factor;
            Left = (5 + left) / factor - 4;
            SearchBox.Focus();
        }

        private async void Me_Loaded(object sender, RoutedEventArgs e)
        {
            file = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Math.txt");
            await loadSuggestions();
        }

        private void Me_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                case Key.Tab:
                    if (LiveSuggestions.Any())
                    {
                        SuggestionsList.SelectedIndex = 0;
                        AddSelectedSuggestionFromList();
                    }
                    else
                    {
                        var text = SearchBox.Text;
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
                case Key.Escape when Keyboard.IsKeyDown(Key.LeftShift):
                    Close();
                    break;
                case Key.F5:
                    RefreshSuggestions();
                    break;
                case Key.F12:
                    OpenSuggestionsFile();
                    break;
                default:
                    break;
            }
        }


        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
