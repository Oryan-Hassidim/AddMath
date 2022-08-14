namespace AddMath.WordAddIn
{
    partial class AddMathRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddMathRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AddMathGroup = this.Factory.CreateRibbonGroup();
            this.AddMathButton = this.Factory.CreateRibbonButton();
            this.EditSuggestionsButton = this.Factory.CreateRibbonButton();
            this.InitializeButton = this.Factory.CreateRibbonButton();
            this.ThemeDropBox = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.AddMathGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.AddMathGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // AddMathGroup
            // 
            this.AddMathGroup.Items.Add(this.AddMathButton);
            this.AddMathGroup.Items.Add(this.EditSuggestionsButton);
            this.AddMathGroup.Items.Add(this.InitializeButton);
            this.AddMathGroup.Items.Add(this.ThemeDropBox);
            this.AddMathGroup.KeyTip = "M";
            this.AddMathGroup.Label = "Add Math";
            this.AddMathGroup.Name = "AddMathGroup";
            // 
            // AddMathButton
            // 
            this.AddMathButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddMathButton.KeyTip = "M";
            this.AddMathButton.Label = "Add Math";
            this.AddMathButton.Name = "AddMathButton";
            this.AddMathButton.OfficeImageId = "EquationInsertGallery";
            this.AddMathButton.ShowImage = true;
            this.AddMathButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddMathButton_Click);
            // 
            // EditSuggestionsButton
            // 
            this.EditSuggestionsButton.KeyTip = "E";
            this.EditSuggestionsButton.Label = "Edit Suggestions";
            this.EditSuggestionsButton.Name = "EditSuggestionsButton";
            this.EditSuggestionsButton.OfficeImageId = "EditExpression";
            this.EditSuggestionsButton.ShowImage = true;
            this.EditSuggestionsButton.Tag = "";
            this.EditSuggestionsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditSuggestions_Click);
            // 
            // InitializeButton
            // 
            this.InitializeButton.KeyTip = "I";
            this.InitializeButton.Label = "Initialize Shortcuts";
            this.InitializeButton.Name = "InitializeButton";
            this.InitializeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InitializeButton_Click);
            // 
            // ThemeDropBox
            // 
            this.ThemeDropBox.KeyTip = "T";
            this.ThemeDropBox.Label = "Theme";
            this.ThemeDropBox.Name = "ThemeDropBox";
            this.ThemeDropBox.SizeString = "Default 2";
            this.ThemeDropBox.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ThemeDropBox_SelectionChanged);
            // 
            // AddMathRibbon
            // 
            this.Name = "AddMathRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AddMathGroup.ResumeLayout(false);
            this.AddMathGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AddMathGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddMathButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EditSuggestionsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InitializeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ThemeDropBox;
    }

    partial class ThisRibbonCollection
    {
        internal AddMathRibbon Ribbon
        {
            get { return this.GetRibbon<AddMathRibbon>(); }
        }
    }
}
