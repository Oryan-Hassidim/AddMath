using AddMath.WordAddIn.Properties;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AddMath.WordAddIn
{
    public partial class AddMathRibbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            //TODO: multi lingual
            initializeThemeItems();
            Settings.Default.Saved += (s, e) => initializeThemeItems();
        }

        private void initializeThemeItems()
        {
            ThemeDropBox.Items.Clear();
            foreach (var item in Settings.Default.SuggestionsDictionary.Keys)
            {
                var ribbonItem = Factory.CreateRibbonDropDownItem();
                ribbonItem.Label = item;
                ThemeDropBox.Items.Add(ribbonItem);
                if (Settings.Default.Selected == item)
                    ThemeDropBox.SelectedItem = ribbonItem;
            }            
        }

        private void AddMathButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddMathObject.AddMath();
        }

        private void EditSuggestions_Click(object sender, RibbonControlEventArgs e)
        {
            EditSuggestions editSuggestions = new();
            editSuggestions.Show();
        }

        private void InitializeButton_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Templates.Count == 0)
            {
                MessageBox.Show("Must Be an active template!", "Add Math", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var vba = new VBA();
            vba.Show();
        }

        private void ThemeDropBox_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.Selected = ThemeDropBox.SelectedItem.Label;
            Settings.Default.Save();
        }
    }
}
