namespace CharacterCounter
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.openInputDialogue = new System.Windows.Forms.OpenFileDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.OutputFileBox = new System.Windows.Forms.TextBox();
            this.BtnOutputFile = new System.Windows.Forms.Button();
            this.BtnClose = new System.Windows.Forms.Button();
            this.saveExcelDialogue = new System.Windows.Forms.SaveFileDialog();
            this.BtnAnalyse = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
            this.numericCharsAfter = new System.Windows.Forms.NumericUpDown();
            this.label24 = new System.Windows.Forms.Label();
            this.numericCharsBefore = new System.Windows.Forms.NumericUpDown();
            this.label23 = new System.Windows.Forms.Label();
            this.boxContextChars = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.checkGetContext = new System.Windows.Forms.CheckBox();
            this.checkCountCharacters = new System.Windows.Forms.CheckBox();
            this.BtnCheckContextFile = new System.Windows.Forms.Button();
            this.BtnContextCharFile = new System.Windows.Forms.Button();
            this.ContextCharacterFileBox = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.BtnAggregateFile = new System.Windows.Forms.Button();
            this.AggregateStatsBox = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabFonts = new System.Windows.Forms.TabPage();
            this.BtnSaveFontList = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.FontList = new System.Windows.Forms.ListBox();
            this.BtnListFonts = new System.Windows.Forms.Button();
            this.tabStyles = new System.Windows.Forms.TabPage();
            this.BtnSaveStyles = new System.Windows.Forms.Button();
            this.listStyles = new System.Windows.Forms.DataGridView();
            this.Style = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.theDefaultFont = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnGetStyles = new System.Windows.Forms.Button();
            this.ErrorTab = new System.Windows.Forms.TabPage();
            this.label11 = new System.Windows.Forms.Label();
            this.BtnSaveErrorList = new System.Windows.Forms.Button();
            this.listNormalisedErrors = new System.Windows.Forms.DataGridView();
            this.MappedCharacter = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PossibleCharacter = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BtnGetEncoding = new System.Windows.Forms.Button();
            this.EncodingTextBox = new System.Windows.Forms.TextBox();
            this.BtnGetFont = new System.Windows.Forms.Button();
            this.FontBox = new System.Windows.Forms.TextBox();
            this.FontLabel = new System.Windows.Forms.Label();
            this.BtnDecompGlyph = new System.Windows.Forms.Button();
            this.DecompGlyphBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.WriteIndividualFile = new System.Windows.Forms.CheckBox();
            this.BtnSaveAggregateStats = new System.Windows.Forms.Button();
            this.label14 = new System.Windows.Forms.Label();
            this.FileCounter = new System.Windows.Forms.Label();
            this.AggregateStats = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.CombDecomposedChars = new System.Windows.Forms.CheckBox();
            this.AnalyseByFont = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.defaultsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CountCharactersDefault = new System.Windows.Forms.ToolStripMenuItem();
            this.AnalyseContextDefault = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.documentationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CombiningCharacters = new System.Windows.Forms.ToolStripMenuItem();
            this.LicenseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.IndivOrBulk = new System.Windows.Forms.TabControl();
            this.IndividualFile = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.BtnGetInput = new System.Windows.Forms.Button();
            this.InputFileBox = new System.Windows.Forms.TextBox();
            this.BtnSaveXML = new System.Windows.Forms.Button();
            this.XMLFileBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.FontListFileBox = new System.Windows.Forms.TextBox();
            this.StyleListFileBox = new System.Windows.Forms.TextBox();
            this.BtnErrorList = new System.Windows.Forms.Button();
            this.ErrorListBox = new System.Windows.Forms.TextBox();
            this.BtnFontListFile = new System.Windows.Forms.Button();
            this.BtnStyleListFile = new System.Windows.Forms.Button();
            this.BtnXMLFile = new System.Windows.Forms.Button();
            this.Bulk = new System.Windows.Forms.TabPage();
            this.BtnSelectFiles = new System.Windows.Forms.Button();
            this.OutputFileSuffixBox = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.BtnInputFolder = new System.Windows.Forms.Button();
            this.InputFolderBox = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.BtnCharStatFolder = new System.Windows.Forms.Button();
            this.OutputFolderBox = new System.Windows.Forms.TextBox();
            this.BulkFontListFileBox = new System.Windows.Forms.TextBox();
            this.BulkStyleListBox = new System.Windows.Forms.TextBox();
            this.BtnBulkErrorList = new System.Windows.Forms.Button();
            this.BulkErrorListBox = new System.Windows.Forms.TextBox();
            this.BtnBulkFontListFile = new System.Windows.Forms.Button();
            this.BtnBulkStyleListFile = new System.Windows.Forms.Button();
            this.toolStripContainer2 = new System.Windows.Forms.ToolStripContainer();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.saveXMLDialogue = new System.Windows.Forms.SaveFileDialog();
            this.toolTipCombine = new System.Windows.Forms.ToolTip(this.components);
            this.openGlyphFileDialogue = new System.Windows.Forms.OpenFileDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.FolderDialogue = new System.Windows.Forms.FolderBrowserDialog();
            this.openContextCharFileDialogue = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1.SuspendLayout();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericCharsAfter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericCharsBefore)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabFonts.SuspendLayout();
            this.tabStyles.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listStyles)).BeginInit();
            this.ErrorTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listNormalisedErrors)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.IndivOrBulk.SuspendLayout();
            this.IndividualFile.SuspendLayout();
            this.Bulk.SuspendLayout();
            this.toolStripContainer2.ContentPanel.SuspendLayout();
            this.toolStripContainer2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(96, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(453, 37);
            this.label1.TabIndex = 0;
            this.label1.Text = "Count Glyphs in a Document";
            // 
            // openInputDialogue
            // 
            this.openInputDialogue.DefaultExt = "docx";
            this.openInputDialogue.Filter = "Word, Rich Text or Text files |*.doc;*.docx;*.rtf;*.txt|All files| *.*";
            this.openInputDialogue.Title = "Input File";
            this.openInputDialogue.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenInputDialogue_FileOk);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 36);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(94, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Character stats file";
            // 
            // OutputFileBox
            // 
            this.OutputFileBox.Location = new System.Drawing.Point(122, 33);
            this.OutputFileBox.Name = "OutputFileBox";
            this.OutputFileBox.Size = new System.Drawing.Size(408, 20);
            this.OutputFileBox.TabIndex = 5;
            // 
            // BtnOutputFile
            // 
            this.BtnOutputFile.Location = new System.Drawing.Point(537, 31);
            this.BtnOutputFile.Name = "BtnOutputFile";
            this.BtnOutputFile.Size = new System.Drawing.Size(75, 23);
            this.BtnOutputFile.TabIndex = 6;
            this.BtnOutputFile.Text = "Browse";
            this.BtnOutputFile.UseVisualStyleBackColor = true;
            this.BtnOutputFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(631, 649);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(101, 30);
            this.BtnClose.TabIndex = 7;
            this.BtnClose.Text = "Close";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // saveExcelDialogue
            // 
            this.saveExcelDialogue.Filter = "Excel WorkBook | *.xlsx";
            // 
            // BtnAnalyse
            // 
            this.BtnAnalyse.Enabled = false;
            this.BtnAnalyse.Location = new System.Drawing.Point(605, 395);
            this.BtnAnalyse.Name = "BtnAnalyse";
            this.BtnAnalyse.Size = new System.Drawing.Size(101, 33);
            this.BtnAnalyse.TabIndex = 8;
            this.BtnAnalyse.Text = "Analyse";
            this.BtnAnalyse.UseVisualStyleBackColor = true;
            this.BtnAnalyse.Click += new System.EventHandler(this.BtnAnalyse_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(4, 671);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(219, 22);
            this.statusStrip1.TabIndex = 9;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(200, 16);
            // 
            // toolStripContainer1
            // 
            // 
            // toolStripContainer1.ContentPanel
            // 
            this.toolStripContainer1.ContentPanel.AutoScroll = true;
            this.toolStripContainer1.ContentPanel.Controls.Add(this.numericCharsAfter);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label24);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.numericCharsBefore);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label23);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.boxContextChars);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label22);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.checkGetContext);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.checkCountCharacters);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnCheckContextFile);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnContextCharFile);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.ContextCharacterFileBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label20);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnAggregateFile);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.AggregateStatsBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label13);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.tabControl1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnGetEncoding);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.EncodingTextBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnAnalyse);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnGetFont);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.FontBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.FontLabel);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnDecompGlyph);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.DecompGlyphBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label9);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.WriteIndividualFile);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnSaveAggregateStats);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label14);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.FileCounter);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.BtnClose);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.AggregateStats);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label12);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.CombDecomposedChars);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.AnalyseByFont);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.textBox1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.statusStrip1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.menuStrip1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.IndivOrBulk);
            this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(838, 724);
            this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer1.LeftToolStripPanelVisible = false;
            this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer1.Name = "toolStripContainer1";
            this.toolStripContainer1.RightToolStripPanelVisible = false;
            this.toolStripContainer1.Size = new System.Drawing.Size(838, 724);
            this.toolStripContainer1.TabIndex = 10;
            this.toolStripContainer1.Text = "toolStripContainer1";
            this.toolStripContainer1.TopToolStripPanelVisible = false;
            // 
            // numericCharsAfter
            // 
            this.numericCharsAfter.Location = new System.Drawing.Point(474, 478);
            this.numericCharsAfter.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericCharsAfter.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericCharsAfter.Name = "numericCharsAfter";
            this.numericCharsAfter.Size = new System.Drawing.Size(30, 20);
            this.numericCharsAfter.TabIndex = 85;
            this.numericCharsAfter.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.numericCharsAfter.ValueChanged += new System.EventHandler(this.Numeric_ValueChanged);
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(386, 482);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(82, 13);
            this.label24.TabIndex = 84;
            this.label24.Text = "Characters after";
            // 
            // numericCharsBefore
            // 
            this.numericCharsBefore.Location = new System.Drawing.Point(350, 479);
            this.numericCharsBefore.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericCharsBefore.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericCharsBefore.Name = "numericCharsBefore";
            this.numericCharsBefore.Size = new System.Drawing.Size(30, 20);
            this.numericCharsBefore.TabIndex = 83;
            this.numericCharsBefore.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(252, 480);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(91, 13);
            this.label23.TabIndex = 82;
            this.label23.Text = "Characters before";
            // 
            // boxContextChars
            // 
            this.boxContextChars.Location = new System.Drawing.Point(137, 479);
            this.boxContextChars.Name = "boxContextChars";
            this.boxContextChars.Size = new System.Drawing.Size(103, 20);
            this.boxContextChars.TabIndex = 81;
            this.toolTip1.SetToolTip(this.boxContextChars, "Space separated values in the form U+nnnn");
            this.boxContextChars.Validating += new System.ComponentModel.CancelEventHandler(this.BoxContextChars_Validating);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(19, 479);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(102, 13);
            this.label22.TabIndex = 79;
            this.label22.Text = "Context character(s)";
            // 
            // checkGetContext
            // 
            this.checkGetContext.AutoSize = true;
            this.checkGetContext.Enabled = false;
            this.checkGetContext.Location = new System.Drawing.Point(604, 362);
            this.checkGetContext.Name = "checkGetContext";
            this.checkGetContext.Size = new System.Drawing.Size(87, 17);
            this.checkGetContext.TabIndex = 78;
            this.checkGetContext.Text = "Get Contexts";
            this.checkGetContext.UseVisualStyleBackColor = true;
            this.checkGetContext.CheckedChanged += new System.EventHandler(this.CheckBox_CheckedChanged);
            // 
            // checkCountCharacters
            // 
            this.checkCountCharacters.AutoSize = true;
            this.checkCountCharacters.Enabled = false;
            this.checkCountCharacters.Location = new System.Drawing.Point(605, 331);
            this.checkCountCharacters.Name = "checkCountCharacters";
            this.checkCountCharacters.Size = new System.Drawing.Size(108, 17);
            this.checkCountCharacters.TabIndex = 77;
            this.checkCountCharacters.Text = "Count Characters";
            this.checkCountCharacters.UseVisualStyleBackColor = true;
            this.checkCountCharacters.CheckedChanged += new System.EventHandler(this.CheckBox_CheckedChanged);
            // 
            // BtnCheckContextFile
            // 
            this.BtnCheckContextFile.Enabled = false;
            this.BtnCheckContextFile.Location = new System.Drawing.Point(641, 538);
            this.BtnCheckContextFile.Name = "BtnCheckContextFile";
            this.BtnCheckContextFile.Size = new System.Drawing.Size(97, 24);
            this.BtnCheckContextFile.TabIndex = 76;
            this.BtnCheckContextFile.Text = "Check";
            this.BtnCheckContextFile.UseVisualStyleBackColor = true;
            this.BtnCheckContextFile.Click += new System.EventHandler(this.BtnCheckContextFile_Click);
            // 
            // BtnContextCharFile
            // 
            this.BtnContextCharFile.Location = new System.Drawing.Point(557, 540);
            this.BtnContextCharFile.Name = "BtnContextCharFile";
            this.BtnContextCharFile.Size = new System.Drawing.Size(75, 23);
            this.BtnContextCharFile.TabIndex = 57;
            this.BtnContextCharFile.Text = "Browse";
            this.BtnContextCharFile.UseVisualStyleBackColor = true;
            this.BtnContextCharFile.Click += new System.EventHandler(this.BtnGetInput_Click);
            // 
            // ContextCharacterFileBox
            // 
            this.ContextCharacterFileBox.Location = new System.Drawing.Point(137, 541);
            this.ContextCharacterFileBox.Name = "ContextCharacterFileBox";
            this.ContextCharacterFileBox.Size = new System.Drawing.Size(408, 20);
            this.ContextCharacterFileBox.TabIndex = 56;
            this.ContextCharacterFileBox.TextChanged += new System.EventHandler(this.ContextCharacterFileBox_TextChanged);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(17, 545);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(122, 13);
            this.label20.TabIndex = 75;
            this.label20.Text = "Context character list file";
            // 
            // BtnAggregateFile
            // 
            this.BtnAggregateFile.Location = new System.Drawing.Point(557, 598);
            this.BtnAggregateFile.Name = "BtnAggregateFile";
            this.BtnAggregateFile.Size = new System.Drawing.Size(75, 23);
            this.BtnAggregateFile.TabIndex = 64;
            this.BtnAggregateFile.Text = "Browse";
            this.BtnAggregateFile.UseVisualStyleBackColor = true;
            this.BtnAggregateFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // AggregateStatsBox
            // 
            this.AggregateStatsBox.Location = new System.Drawing.Point(137, 599);
            this.AggregateStatsBox.Name = "AggregateStatsBox";
            this.AggregateStatsBox.Size = new System.Drawing.Size(408, 20);
            this.AggregateStatsBox.TabIndex = 63;
            this.toolTip1.SetToolTip(this.AggregateStatsBox, "File that aggregates statistics from aggregate input files.\r\n");
            this.AggregateStatsBox.TextChanged += new System.EventHandler(this.AggregateStatsBox_TextChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(17, 603);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(97, 13);
            this.label13.TabIndex = 62;
            this.label13.Text = "Aggregate stats file";
            this.label13.UseMnemonic = false;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabFonts);
            this.tabControl1.Controls.Add(this.tabStyles);
            this.tabControl1.Controls.Add(this.ErrorTab);
            this.tabControl1.Location = new System.Drawing.Point(13, 290);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(573, 183);
            this.tabControl1.TabIndex = 61;
            // 
            // tabFonts
            // 
            this.tabFonts.Controls.Add(this.BtnSaveFontList);
            this.tabFonts.Controls.Add(this.label4);
            this.tabFonts.Controls.Add(this.FontList);
            this.tabFonts.Controls.Add(this.BtnListFonts);
            this.tabFonts.Location = new System.Drawing.Point(4, 22);
            this.tabFonts.Name = "tabFonts";
            this.tabFonts.Padding = new System.Windows.Forms.Padding(3);
            this.tabFonts.Size = new System.Drawing.Size(565, 157);
            this.tabFonts.TabIndex = 0;
            this.tabFonts.Text = "Get fonts";
            this.tabFonts.UseVisualStyleBackColor = true;
            // 
            // BtnSaveFontList
            // 
            this.BtnSaveFontList.Enabled = false;
            this.BtnSaveFontList.Location = new System.Drawing.Point(10, 98);
            this.BtnSaveFontList.Name = "BtnSaveFontList";
            this.BtnSaveFontList.Size = new System.Drawing.Size(119, 45);
            this.BtnSaveFontList.TabIndex = 21;
            this.BtnSaveFontList.Text = "Save font list";
            this.BtnSaveFontList.UseVisualStyleBackColor = true;
            this.BtnSaveFontList.Click += new System.EventHandler(this.BtnListFonts_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(-2, 5);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(131, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "The fonts in the document";
            // 
            // FontList
            // 
            this.FontList.FormattingEnabled = true;
            this.FontList.Location = new System.Drawing.Point(135, 5);
            this.FontList.Name = "FontList";
            this.FontList.Size = new System.Drawing.Size(424, 147);
            this.FontList.TabIndex = 19;
            // 
            // BtnListFonts
            // 
            this.BtnListFonts.Enabled = false;
            this.BtnListFonts.Location = new System.Drawing.Point(10, 47);
            this.BtnListFonts.Name = "BtnListFonts";
            this.BtnListFonts.Size = new System.Drawing.Size(119, 45);
            this.BtnListFonts.TabIndex = 18;
            this.BtnListFonts.Text = "List the fonts";
            this.BtnListFonts.UseVisualStyleBackColor = true;
            this.BtnListFonts.Click += new System.EventHandler(this.BtnListFonts_Click);
            // 
            // tabStyles
            // 
            this.tabStyles.Controls.Add(this.BtnSaveStyles);
            this.tabStyles.Controls.Add(this.listStyles);
            this.tabStyles.Controls.Add(this.label5);
            this.tabStyles.Controls.Add(this.BtnGetStyles);
            this.tabStyles.Location = new System.Drawing.Point(4, 22);
            this.tabStyles.Name = "tabStyles";
            this.tabStyles.Padding = new System.Windows.Forms.Padding(3);
            this.tabStyles.Size = new System.Drawing.Size(565, 157);
            this.tabStyles.TabIndex = 1;
            this.tabStyles.Text = "Get Styles";
            this.tabStyles.UseVisualStyleBackColor = true;
            // 
            // BtnSaveStyles
            // 
            this.BtnSaveStyles.Enabled = false;
            this.BtnSaveStyles.Location = new System.Drawing.Point(3, 125);
            this.BtnSaveStyles.Name = "BtnSaveStyles";
            this.BtnSaveStyles.Size = new System.Drawing.Size(102, 24);
            this.BtnSaveStyles.TabIndex = 3;
            this.BtnSaveStyles.Text = "Save styles";
            this.BtnSaveStyles.UseVisualStyleBackColor = true;
            this.BtnSaveStyles.Click += new System.EventHandler(this.BtnGetStyles_Click);
            // 
            // listStyles
            // 
            this.listStyles.AllowUserToAddRows = false;
            this.listStyles.AllowUserToDeleteRows = false;
            this.listStyles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.listStyles.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Style,
            this.theDefaultFont});
            this.listStyles.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2;
            this.listStyles.Location = new System.Drawing.Point(118, 4);
            this.listStyles.MultiSelect = false;
            this.listStyles.Name = "listStyles";
            this.listStyles.ReadOnly = true;
            this.listStyles.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.listStyles.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;
            this.listStyles.RowTemplate.ReadOnly = true;
            this.listStyles.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.listStyles.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.listStyles.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.listStyles.ShowCellToolTips = false;
            this.listStyles.ShowEditingIcon = false;
            this.listStyles.Size = new System.Drawing.Size(484, 150);
            this.listStyles.TabIndex = 2;
            this.listStyles.TabStop = false;
            // 
            // Style
            // 
            this.Style.HeaderText = "Style";
            this.Style.Name = "Style";
            this.Style.ReadOnly = true;
            this.Style.Width = 230;
            // 
            // theDefaultFont
            // 
            this.theDefaultFont.HeaderText = "Default Font";
            this.theDefaultFont.Name = "theDefaultFont";
            this.theDefaultFont.ReadOnly = true;
            this.theDefaultFont.Width = 230;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(113, 92);
            this.label5.TabIndex = 2;
            this.label5.Text = "This lets you inspect\r\nthe styles and their\r\ndefault fonts. It is\r\nthere so you c" +
    "an\r\ncheck if the\r\napplication is\r\nworking properly\r\n";
            this.label5.UseCompatibleTextRendering = true;
            // 
            // BtnGetStyles
            // 
            this.BtnGetStyles.Enabled = false;
            this.BtnGetStyles.Location = new System.Drawing.Point(3, 95);
            this.BtnGetStyles.Name = "BtnGetStyles";
            this.BtnGetStyles.Size = new System.Drawing.Size(102, 24);
            this.BtnGetStyles.TabIndex = 0;
            this.BtnGetStyles.Text = "Get styles";
            this.BtnGetStyles.UseVisualStyleBackColor = true;
            this.BtnGetStyles.Click += new System.EventHandler(this.BtnGetStyles_Click);
            // 
            // ErrorTab
            // 
            this.ErrorTab.Controls.Add(this.label11);
            this.ErrorTab.Controls.Add(this.BtnSaveErrorList);
            this.ErrorTab.Controls.Add(this.listNormalisedErrors);
            this.ErrorTab.Location = new System.Drawing.Point(4, 22);
            this.ErrorTab.Name = "ErrorTab";
            this.ErrorTab.Padding = new System.Windows.Forms.Padding(3);
            this.ErrorTab.Size = new System.Drawing.Size(565, 157);
            this.ErrorTab.TabIndex = 2;
            this.ErrorTab.Text = "Normalisation suggestions";
            this.ErrorTab.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(14, 16);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(71, 65);
            this.label11.TabIndex = 2;
            this.label11.Text = "This lists the\r\nsuggested\r\ncharacters to\r\nuse after\r\nnormalisation.";
            // 
            // BtnSaveErrorList
            // 
            this.BtnSaveErrorList.Enabled = false;
            this.BtnSaveErrorList.Location = new System.Drawing.Point(7, 98);
            this.BtnSaveErrorList.Name = "BtnSaveErrorList";
            this.BtnSaveErrorList.Size = new System.Drawing.Size(87, 53);
            this.BtnSaveErrorList.TabIndex = 1;
            this.BtnSaveErrorList.Text = "Save data";
            this.BtnSaveErrorList.UseVisualStyleBackColor = true;
            this.BtnSaveErrorList.Click += new System.EventHandler(this.BtnSaveErrorList_Click);
            // 
            // listNormalisedErrors
            // 
            this.listNormalisedErrors.AllowUserToAddRows = false;
            this.listNormalisedErrors.AllowUserToDeleteRows = false;
            this.listNormalisedErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.listNormalisedErrors.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.MappedCharacter,
            this.PossibleCharacter});
            this.listNormalisedErrors.Location = new System.Drawing.Point(107, 3);
            this.listNormalisedErrors.Name = "listNormalisedErrors";
            this.listNormalisedErrors.ReadOnly = true;
            this.listNormalisedErrors.Size = new System.Drawing.Size(462, 150);
            this.listNormalisedErrors.TabIndex = 0;
            // 
            // MappedCharacter
            // 
            this.MappedCharacter.HeaderText = "Mapped Character";
            this.MappedCharacter.Name = "MappedCharacter";
            this.MappedCharacter.ReadOnly = true;
            this.MappedCharacter.Width = 250;
            // 
            // PossibleCharacter
            // 
            this.PossibleCharacter.HeaderText = "Possible Character";
            this.PossibleCharacter.Name = "PossibleCharacter";
            this.PossibleCharacter.ReadOnly = true;
            this.PossibleCharacter.Width = 200;
            // 
            // BtnGetEncoding
            // 
            this.BtnGetEncoding.Location = new System.Drawing.Point(17, 568);
            this.BtnGetEncoding.Name = "BtnGetEncoding";
            this.BtnGetEncoding.Size = new System.Drawing.Size(75, 23);
            this.BtnGetEncoding.TabIndex = 59;
            this.BtnGetEncoding.Text = "Encoding";
            this.BtnGetEncoding.UseVisualStyleBackColor = true;
            this.BtnGetEncoding.Click += new System.EventHandler(this.BtnGetEncoding_Click);
            // 
            // EncodingTextBox
            // 
            this.EncodingTextBox.Enabled = false;
            this.EncodingTextBox.Location = new System.Drawing.Point(137, 569);
            this.EncodingTextBox.Name = "EncodingTextBox";
            this.EncodingTextBox.Size = new System.Drawing.Size(180, 20);
            this.EncodingTextBox.TabIndex = 60;
            // 
            // BtnGetFont
            // 
            this.BtnGetFont.Enabled = false;
            this.BtnGetFont.Location = new System.Drawing.Point(557, 568);
            this.BtnGetFont.Name = "BtnGetFont";
            this.BtnGetFont.Size = new System.Drawing.Size(75, 23);
            this.BtnGetFont.TabIndex = 58;
            this.BtnGetFont.Text = "Browse";
            this.BtnGetFont.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.toolTip1.SetToolTip(this.BtnGetFont, "Get the font you want to use.");
            this.BtnGetFont.UseVisualStyleBackColor = true;
            this.BtnGetFont.Click += new System.EventHandler(this.BtnGetFont_Click);
            // 
            // FontBox
            // 
            this.FontBox.Enabled = false;
            this.FontBox.Location = new System.Drawing.Point(361, 569);
            this.FontBox.Name = "FontBox";
            this.FontBox.Size = new System.Drawing.Size(184, 20);
            this.FontBox.TabIndex = 57;
            this.FontBox.Text = "Calibri";
            // 
            // FontLabel
            // 
            this.FontLabel.AutoSize = true;
            this.FontLabel.Enabled = false;
            this.FontLabel.Location = new System.Drawing.Point(328, 573);
            this.FontLabel.Name = "FontLabel";
            this.FontLabel.Size = new System.Drawing.Size(28, 13);
            this.FontLabel.TabIndex = 56;
            this.FontLabel.Text = "Font";
            // 
            // BtnDecompGlyph
            // 
            this.BtnDecompGlyph.Location = new System.Drawing.Point(557, 508);
            this.BtnDecompGlyph.Name = "BtnDecompGlyph";
            this.BtnDecompGlyph.Size = new System.Drawing.Size(75, 23);
            this.BtnDecompGlyph.TabIndex = 55;
            this.BtnDecompGlyph.Text = "Browse";
            this.toolTip1.SetToolTip(this.BtnDecompGlyph, "Browse for the font for text files");
            this.BtnDecompGlyph.UseVisualStyleBackColor = true;
            this.BtnDecompGlyph.Click += new System.EventHandler(this.BtnGetInput_Click);
            // 
            // DecompGlyphBox
            // 
            this.DecompGlyphBox.Location = new System.Drawing.Point(137, 509);
            this.DecompGlyphBox.Name = "DecompGlyphBox";
            this.DecompGlyphBox.Size = new System.Drawing.Size(408, 20);
            this.DecompGlyphBox.TabIndex = 54;
            this.toolTip1.SetToolTip(this.DecompGlyphBox, "An Excel files with the decomposed characters and their fonts in cells A2 downwar" +
        "ds.  It is only useful if you are analysing by font.");
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(17, 513);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(66, 13);
            this.label9.TabIndex = 53;
            this.label9.Text = "Diacritics file";
            // 
            // WriteIndividualFile
            // 
            this.WriteIndividualFile.AutoSize = true;
            this.WriteIndividualFile.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.WriteIndividualFile.Checked = true;
            this.WriteIndividualFile.CheckState = System.Windows.Forms.CheckState.Checked;
            this.WriteIndividualFile.Location = new System.Drawing.Point(127, 627);
            this.WriteIndividualFile.Name = "WriteIndividualFile";
            this.WriteIndividualFile.Size = new System.Drawing.Size(114, 17);
            this.WriteIndividualFile.TabIndex = 51;
            this.WriteIndividualFile.Text = "Write individual file";
            this.WriteIndividualFile.UseVisualStyleBackColor = true;
            // 
            // BtnSaveAggregateStats
            // 
            this.BtnSaveAggregateStats.Enabled = false;
            this.BtnSaveAggregateStats.Location = new System.Drawing.Point(631, 595);
            this.BtnSaveAggregateStats.Name = "BtnSaveAggregateStats";
            this.BtnSaveAggregateStats.Size = new System.Drawing.Size(101, 26);
            this.BtnSaveAggregateStats.TabIndex = 50;
            this.BtnSaveAggregateStats.Text = "Save aggregate";
            this.toolTip1.SetToolTip(this.BtnSaveAggregateStats, "Save aggregate statistics");
            this.BtnSaveAggregateStats.UseVisualStyleBackColor = true;
            this.BtnSaveAggregateStats.Click += new System.EventHandler(this.BtnSaveAggregateStats_Click);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(259, 650);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(70, 13);
            this.label14.TabIndex = 49;
            this.label14.Text = "Files counted";
            // 
            // FileCounter
            // 
            this.FileCounter.AutoSize = true;
            this.FileCounter.Location = new System.Drawing.Point(335, 650);
            this.FileCounter.Name = "FileCounter";
            this.FileCounter.Size = new System.Drawing.Size(13, 13);
            this.FileCounter.TabIndex = 48;
            this.FileCounter.Text = "0";
            // 
            // AggregateStats
            // 
            this.AggregateStats.AutoSize = true;
            this.AggregateStats.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.AggregateStats.Enabled = false;
            this.AggregateStats.Location = new System.Drawing.Point(127, 648);
            this.AggregateStats.Name = "AggregateStats";
            this.AggregateStats.Size = new System.Drawing.Size(121, 17);
            this.AggregateStats.TabIndex = 44;
            this.AggregateStats.Text = "Aggregate File Stats";
            this.AggregateStats.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(61, 61);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(525, 13);
            this.label12.TabIndex = 36;
            this.label12.Text = "This program runs hidden copies of Word and Excel. Don\'t close instances of them " +
    "before closing this program";
            // 
            // CombDecomposedChars
            // 
            this.CombDecomposedChars.AutoSize = true;
            this.CombDecomposedChars.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.CombDecomposedChars.Checked = true;
            this.CombDecomposedChars.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CombDecomposedChars.Cursor = System.Windows.Forms.Cursors.Default;
            this.CombDecomposedChars.Location = new System.Drawing.Point(402, 627);
            this.CombDecomposedChars.Name = "CombDecomposedChars";
            this.CombDecomposedChars.Size = new System.Drawing.Size(184, 17);
            this.CombDecomposedChars.TabIndex = 29;
            this.CombDecomposedChars.Text = "Combine decomposed characters";
            this.toolTipCombine.SetToolTip(this.CombDecomposedChars, "Some characters have to be mapped to two Unicode characters that are displayed to" +
        "gether as a single glyph. \r\nChecking this box causes the program to try to count" +
        " them as a single character.");
            this.CombDecomposedChars.UseVisualStyleBackColor = true;
            this.CombDecomposedChars.CheckStateChanged += new System.EventHandler(this.CombDecomposedChars_CheckStateChanged);
            // 
            // AnalyseByFont
            // 
            this.AnalyseByFont.AutoSize = true;
            this.AnalyseByFont.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.AnalyseByFont.Checked = true;
            this.AnalyseByFont.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AnalyseByFont.Location = new System.Drawing.Point(262, 626);
            this.AnalyseByFont.Name = "AnalyseByFont";
            this.AnalyseByFont.Size = new System.Drawing.Size(101, 17);
            this.AnalyseByFont.TabIndex = 14;
            this.AnalyseByFont.Text = "Analyse by Font";
            this.toolTip1.SetToolTip(this.AnalyseByFont, "Analyse by Font is slightly slower.");
            this.AnalyseByFont.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(64, 64);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(452, 13);
            this.textBox1.TabIndex = 13;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.defaultsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(4, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(151, 24);
            this.menuStrip1.TabIndex = 10;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // defaultsToolStripMenuItem
            // 
            this.defaultsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.CountCharactersDefault,
            this.AnalyseContextDefault});
            this.defaultsToolStripMenuItem.Name = "defaultsToolStripMenuItem";
            this.defaultsToolStripMenuItem.Size = new System.Drawing.Size(62, 20);
            this.defaultsToolStripMenuItem.Text = "Defaults";
            // 
            // CountCharactersDefault
            // 
            this.CountCharactersDefault.CheckOnClick = true;
            this.CountCharactersDefault.Name = "CountCharactersDefault";
            this.CountCharactersDefault.Size = new System.Drawing.Size(166, 22);
            this.CountCharactersDefault.Text = "Count Characters";
            this.CountCharactersDefault.CheckedChanged += new System.EventHandler(this.Default_CheckedChanged);
            // 
            // AnalyseContextDefault
            // 
            this.AnalyseContextDefault.CheckOnClick = true;
            this.AnalyseContextDefault.Name = "AnalyseContextDefault";
            this.AnalyseContextDefault.Size = new System.Drawing.Size(166, 22);
            this.AnalyseContextDefault.Text = "Analyse Contexts";
            this.AnalyseContextDefault.CheckedChanged += new System.EventHandler(this.Default_CheckedChanged);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.documentationToolStripMenuItem,
            this.CombiningCharacters,
            this.LicenseMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // documentationToolStripMenuItem
            // 
            this.documentationToolStripMenuItem.Name = "documentationToolStripMenuItem";
            this.documentationToolStripMenuItem.Size = new System.Drawing.Size(291, 22);
            this.documentationToolStripMenuItem.Text = "Documentation";
            this.documentationToolStripMenuItem.Click += new System.EventHandler(this.DocumentationToolStripMenuItem_Click);
            // 
            // CombiningCharacters
            // 
            this.CombiningCharacters.Name = "CombiningCharacters";
            this.CombiningCharacters.Size = new System.Drawing.Size(291, 22);
            this.CombiningCharacters.Text = "Legacy Combining Characters Workbook";
            this.CombiningCharacters.Click += new System.EventHandler(this.CombiningCharacters_Click);
            // 
            // LicenseMenuItem
            // 
            this.LicenseMenuItem.Name = "LicenseMenuItem";
            this.LicenseMenuItem.Size = new System.Drawing.Size(291, 22);
            this.LicenseMenuItem.Text = "License";
            this.LicenseMenuItem.Click += new System.EventHandler(this.LicenseMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(291, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
            // 
            // IndivOrBulk
            // 
            this.IndivOrBulk.Controls.Add(this.IndividualFile);
            this.IndivOrBulk.Controls.Add(this.Bulk);
            this.IndivOrBulk.Location = new System.Drawing.Point(10, 81);
            this.IndivOrBulk.Name = "IndivOrBulk";
            this.IndivOrBulk.SelectedIndex = 0;
            this.IndivOrBulk.Size = new System.Drawing.Size(732, 203);
            this.IndivOrBulk.TabIndex = 52;
            // 
            // IndividualFile
            // 
            this.IndividualFile.Controls.Add(this.label2);
            this.IndividualFile.Controls.Add(this.BtnGetInput);
            this.IndividualFile.Controls.Add(this.InputFileBox);
            this.IndividualFile.Controls.Add(this.BtnSaveXML);
            this.IndividualFile.Controls.Add(this.XMLFileBox);
            this.IndividualFile.Controls.Add(this.label3);
            this.IndividualFile.Controls.Add(this.label6);
            this.IndividualFile.Controls.Add(this.label7);
            this.IndividualFile.Controls.Add(this.label10);
            this.IndividualFile.Controls.Add(this.label8);
            this.IndividualFile.Controls.Add(this.BtnOutputFile);
            this.IndividualFile.Controls.Add(this.OutputFileBox);
            this.IndividualFile.Controls.Add(this.FontListFileBox);
            this.IndividualFile.Controls.Add(this.StyleListFileBox);
            this.IndividualFile.Controls.Add(this.BtnErrorList);
            this.IndividualFile.Controls.Add(this.ErrorListBox);
            this.IndividualFile.Controls.Add(this.BtnFontListFile);
            this.IndividualFile.Controls.Add(this.BtnStyleListFile);
            this.IndividualFile.Controls.Add(this.BtnXMLFile);
            this.IndividualFile.Location = new System.Drawing.Point(4, 22);
            this.IndividualFile.Name = "IndividualFile";
            this.IndividualFile.Padding = new System.Windows.Forms.Padding(3);
            this.IndividualFile.Size = new System.Drawing.Size(724, 177);
            this.IndividualFile.TabIndex = 0;
            this.IndividualFile.Text = "Individual";
            this.IndividualFile.UseVisualStyleBackColor = true;
            this.IndividualFile.Enter += new System.EventHandler(this.Individual_Entered);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 29;
            this.label2.Text = "Input Document";
            // 
            // BtnGetInput
            // 
            this.BtnGetInput.Location = new System.Drawing.Point(537, 6);
            this.BtnGetInput.Name = "BtnGetInput";
            this.BtnGetInput.Size = new System.Drawing.Size(75, 23);
            this.BtnGetInput.TabIndex = 45;
            this.BtnGetInput.Text = "Browse";
            this.BtnGetInput.UseVisualStyleBackColor = true;
            this.BtnGetInput.Click += new System.EventHandler(this.BtnGetInput_Click);
            // 
            // InputFileBox
            // 
            this.InputFileBox.Location = new System.Drawing.Point(123, 6);
            this.InputFileBox.Name = "InputFileBox";
            this.InputFileBox.Size = new System.Drawing.Size(408, 20);
            this.InputFileBox.TabIndex = 44;
            this.InputFileBox.TextChanged += new System.EventHandler(this.InputFileBox_TextChanged);
            // 
            // BtnSaveXML
            // 
            this.BtnSaveXML.Enabled = false;
            this.BtnSaveXML.Location = new System.Drawing.Point(626, 137);
            this.BtnSaveXML.Name = "BtnSaveXML";
            this.BtnSaveXML.Size = new System.Drawing.Size(92, 30);
            this.BtnSaveXML.TabIndex = 28;
            this.BtnSaveXML.Text = "Save XML";
            this.BtnSaveXML.UseVisualStyleBackColor = true;
            this.BtnSaveXML.Click += new System.EventHandler(this.BtnSaveXML_Click);
            // 
            // XMLFileBox
            // 
            this.XMLFileBox.Location = new System.Drawing.Point(122, 141);
            this.XMLFileBox.Name = "XMLFileBox";
            this.XMLFileBox.Size = new System.Drawing.Size(408, 20);
            this.XMLFileBox.TabIndex = 26;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(11, 59);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Font list file";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 94);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(61, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "Style list file";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(11, 120);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(60, 13);
            this.label10.TabIndex = 33;
            this.label10.Text = "Error list file";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(11, 146);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "XML file";
            // 
            // FontListFileBox
            // 
            this.FontListFileBox.Location = new System.Drawing.Point(123, 60);
            this.FontListFileBox.Name = "FontListFileBox";
            this.FontListFileBox.Size = new System.Drawing.Size(408, 20);
            this.FontListFileBox.TabIndex = 20;
            this.FontListFileBox.TextChanged += new System.EventHandler(this.FontListFileBox_TextChanged);
            // 
            // StyleListFileBox
            // 
            this.StyleListFileBox.Location = new System.Drawing.Point(122, 87);
            this.StyleListFileBox.Name = "StyleListFileBox";
            this.StyleListFileBox.Size = new System.Drawing.Size(408, 20);
            this.StyleListFileBox.TabIndex = 23;
            this.StyleListFileBox.TextChanged += new System.EventHandler(this.StyleListFileBox_TextChanged);
            // 
            // BtnErrorList
            // 
            this.BtnErrorList.Location = new System.Drawing.Point(537, 115);
            this.BtnErrorList.Name = "BtnErrorList";
            this.BtnErrorList.Size = new System.Drawing.Size(75, 23);
            this.BtnErrorList.TabIndex = 35;
            this.BtnErrorList.Text = "Browse";
            this.BtnErrorList.UseVisualStyleBackColor = true;
            this.BtnErrorList.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // ErrorListBox
            // 
            this.ErrorListBox.Location = new System.Drawing.Point(122, 114);
            this.ErrorListBox.Name = "ErrorListBox";
            this.ErrorListBox.Size = new System.Drawing.Size(408, 20);
            this.ErrorListBox.TabIndex = 34;
            // 
            // BtnFontListFile
            // 
            this.BtnFontListFile.Location = new System.Drawing.Point(537, 59);
            this.BtnFontListFile.Name = "BtnFontListFile";
            this.BtnFontListFile.Size = new System.Drawing.Size(75, 23);
            this.BtnFontListFile.TabIndex = 21;
            this.BtnFontListFile.Text = "Browse";
            this.BtnFontListFile.UseVisualStyleBackColor = true;
            this.BtnFontListFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // BtnStyleListFile
            // 
            this.BtnStyleListFile.Location = new System.Drawing.Point(537, 86);
            this.BtnStyleListFile.Name = "BtnStyleListFile";
            this.BtnStyleListFile.Size = new System.Drawing.Size(75, 23);
            this.BtnStyleListFile.TabIndex = 24;
            this.BtnStyleListFile.Text = "Browse";
            this.BtnStyleListFile.UseVisualStyleBackColor = true;
            this.BtnStyleListFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // BtnXMLFile
            // 
            this.BtnXMLFile.Location = new System.Drawing.Point(537, 142);
            this.BtnXMLFile.Name = "BtnXMLFile";
            this.BtnXMLFile.Size = new System.Drawing.Size(75, 23);
            this.BtnXMLFile.TabIndex = 27;
            this.BtnXMLFile.Text = "Browse";
            this.BtnXMLFile.UseVisualStyleBackColor = true;
            this.BtnXMLFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // Bulk
            // 
            this.Bulk.Controls.Add(this.BtnSelectFiles);
            this.Bulk.Controls.Add(this.OutputFileSuffixBox);
            this.Bulk.Controls.Add(this.label21);
            this.Bulk.Controls.Add(this.label15);
            this.Bulk.Controls.Add(this.BtnInputFolder);
            this.Bulk.Controls.Add(this.InputFolderBox);
            this.Bulk.Controls.Add(this.label16);
            this.Bulk.Controls.Add(this.label17);
            this.Bulk.Controls.Add(this.label18);
            this.Bulk.Controls.Add(this.label19);
            this.Bulk.Controls.Add(this.BtnCharStatFolder);
            this.Bulk.Controls.Add(this.OutputFolderBox);
            this.Bulk.Controls.Add(this.BulkFontListFileBox);
            this.Bulk.Controls.Add(this.BulkStyleListBox);
            this.Bulk.Controls.Add(this.BtnBulkErrorList);
            this.Bulk.Controls.Add(this.BulkErrorListBox);
            this.Bulk.Controls.Add(this.BtnBulkFontListFile);
            this.Bulk.Controls.Add(this.BtnBulkStyleListFile);
            this.Bulk.Location = new System.Drawing.Point(4, 22);
            this.Bulk.Name = "Bulk";
            this.Bulk.Padding = new System.Windows.Forms.Padding(3);
            this.Bulk.Size = new System.Drawing.Size(724, 177);
            this.Bulk.TabIndex = 1;
            this.Bulk.Text = "Bulk";
            this.Bulk.UseVisualStyleBackColor = true;
            this.Bulk.Enter += new System.EventHandler(this.Bulk_Entered);
            // 
            // BtnSelectFiles
            // 
            this.BtnSelectFiles.Enabled = false;
            this.BtnSelectFiles.Location = new System.Drawing.Point(630, 5);
            this.BtnSelectFiles.Name = "BtnSelectFiles";
            this.BtnSelectFiles.Size = new System.Drawing.Size(87, 35);
            this.BtnSelectFiles.TabIndex = 68;
            this.BtnSelectFiles.Text = "Select files";
            this.BtnSelectFiles.UseVisualStyleBackColor = true;
            this.BtnSelectFiles.Click += new System.EventHandler(this.BtnSelectFiles_Click);
            // 
            // OutputFileSuffixBox
            // 
            this.OutputFileSuffixBox.Location = new System.Drawing.Point(502, 41);
            this.OutputFileSuffixBox.Name = "OutputFileSuffixBox";
            this.OutputFileSuffixBox.Size = new System.Drawing.Size(122, 20);
            this.OutputFileSuffixBox.TabIndex = 65;
            this.OutputFileSuffixBox.TextChanged += new System.EventHandler(this.OutputFileSuffixBox_TextChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(401, 45);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(79, 13);
            this.label21.TabIndex = 64;
            this.label21.Text = "File name suffix";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(14, 16);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(60, 13);
            this.label15.TabIndex = 58;
            this.label15.Text = "Input folder";
            // 
            // BtnInputFolder
            // 
            this.BtnInputFolder.Location = new System.Drawing.Point(549, 13);
            this.BtnInputFolder.Name = "BtnInputFolder";
            this.BtnInputFolder.Size = new System.Drawing.Size(75, 23);
            this.BtnInputFolder.TabIndex = 63;
            this.BtnInputFolder.Text = "Browse";
            this.BtnInputFolder.UseVisualStyleBackColor = true;
            this.BtnInputFolder.Click += new System.EventHandler(this.BtnInputFolder_Click);
            // 
            // InputFolderBox
            // 
            this.InputFolderBox.Location = new System.Drawing.Point(135, 13);
            this.InputFolderBox.Name = "InputFolderBox";
            this.InputFolderBox.Size = new System.Drawing.Size(407, 20);
            this.InputFolderBox.TabIndex = 62;
            this.InputFolderBox.TextChanged += new System.EventHandler(this.InputFolderBox_TextChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(14, 45);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(107, 13);
            this.label16.TabIndex = 46;
            this.label16.Text = "Character stats folder";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(14, 75);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(59, 13);
            this.label17.TabIndex = 49;
            this.label17.Text = "Font list file";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(14, 110);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(61, 13);
            this.label18.TabIndex = 52;
            this.label18.Text = "Style list file";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(14, 146);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(60, 13);
            this.label19.TabIndex = 59;
            this.label19.Text = "Error list file";
            // 
            // BtnCharStatFolder
            // 
            this.BtnCharStatFolder.Location = new System.Drawing.Point(310, 40);
            this.BtnCharStatFolder.Name = "BtnCharStatFolder";
            this.BtnCharStatFolder.Size = new System.Drawing.Size(75, 23);
            this.BtnCharStatFolder.TabIndex = 48;
            this.BtnCharStatFolder.Text = "Browse";
            this.BtnCharStatFolder.UseVisualStyleBackColor = true;
            this.BtnCharStatFolder.Click += new System.EventHandler(this.BtnCharStatFolder_Click);
            // 
            // OutputFolderBox
            // 
            this.OutputFolderBox.Location = new System.Drawing.Point(134, 41);
            this.OutputFolderBox.Name = "OutputFolderBox";
            this.OutputFolderBox.Size = new System.Drawing.Size(170, 20);
            this.OutputFolderBox.TabIndex = 47;
            // 
            // BulkFontListFileBox
            // 
            this.BulkFontListFileBox.Location = new System.Drawing.Point(135, 76);
            this.BulkFontListFileBox.Name = "BulkFontListFileBox";
            this.BulkFontListFileBox.Size = new System.Drawing.Size(408, 20);
            this.BulkFontListFileBox.TabIndex = 50;
            // 
            // BulkStyleListBox
            // 
            this.BulkStyleListBox.Location = new System.Drawing.Point(134, 103);
            this.BulkStyleListBox.Name = "BulkStyleListBox";
            this.BulkStyleListBox.Size = new System.Drawing.Size(408, 20);
            this.BulkStyleListBox.TabIndex = 53;
            // 
            // BtnBulkErrorList
            // 
            this.BtnBulkErrorList.Location = new System.Drawing.Point(549, 141);
            this.BtnBulkErrorList.Name = "BtnBulkErrorList";
            this.BtnBulkErrorList.Size = new System.Drawing.Size(75, 23);
            this.BtnBulkErrorList.TabIndex = 61;
            this.BtnBulkErrorList.Text = "Browse";
            this.BtnBulkErrorList.UseVisualStyleBackColor = true;
            this.BtnBulkErrorList.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // BulkErrorListBox
            // 
            this.BulkErrorListBox.Location = new System.Drawing.Point(134, 140);
            this.BulkErrorListBox.Name = "BulkErrorListBox";
            this.BulkErrorListBox.Size = new System.Drawing.Size(408, 20);
            this.BulkErrorListBox.TabIndex = 60;
            this.BulkErrorListBox.TextChanged += new System.EventHandler(this.BulkErrorListbox_TextChanged);
            // 
            // BtnBulkFontListFile
            // 
            this.BtnBulkFontListFile.Location = new System.Drawing.Point(549, 75);
            this.BtnBulkFontListFile.Name = "BtnBulkFontListFile";
            this.BtnBulkFontListFile.Size = new System.Drawing.Size(75, 23);
            this.BtnBulkFontListFile.TabIndex = 51;
            this.BtnBulkFontListFile.Text = "Browse";
            this.BtnBulkFontListFile.UseVisualStyleBackColor = true;
            this.BtnBulkFontListFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // BtnBulkStyleListFile
            // 
            this.BtnBulkStyleListFile.Location = new System.Drawing.Point(549, 102);
            this.BtnBulkStyleListFile.Name = "BtnBulkStyleListFile";
            this.BtnBulkStyleListFile.Size = new System.Drawing.Size(75, 23);
            this.BtnBulkStyleListFile.TabIndex = 54;
            this.BtnBulkStyleListFile.Text = "Browse";
            this.BtnBulkStyleListFile.UseVisualStyleBackColor = true;
            this.BtnBulkStyleListFile.Click += new System.EventHandler(this.BtnGetOutput_Click);
            // 
            // toolStripContainer2
            // 
            // 
            // toolStripContainer2.ContentPanel
            // 
            this.toolStripContainer2.ContentPanel.AutoScroll = true;
            this.toolStripContainer2.ContentPanel.Controls.Add(this.toolStripContainer1);
            this.toolStripContainer2.ContentPanel.Size = new System.Drawing.Size(838, 724);
            this.toolStripContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer2.LeftToolStripPanelVisible = false;
            this.toolStripContainer2.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer2.Name = "toolStripContainer2";
            this.toolStripContainer2.RightToolStripPanelVisible = false;
            this.toolStripContainer2.Size = new System.Drawing.Size(838, 724);
            this.toolStripContainer2.TabIndex = 11;
            this.toolStripContainer2.Text = "toolStripContainer2";
            this.toolStripContainer2.TopToolStripPanelVisible = false;
            // 
            // toolTip1
            // 
            this.toolTip1.ToolTipTitle = "Analyse by Font";
            // 
            // saveXMLDialogue
            // 
            this.saveXMLDialogue.DefaultExt = "xml";
            this.saveXMLDialogue.Filter = "XML File | *.xml";
            // 
            // toolTipCombine
            // 
            this.toolTipCombine.BackColor = System.Drawing.Color.Khaki;
            this.toolTipCombine.ToolTipTitle = "Combine decomposed characters";
            // 
            // OpenGlyphFileDialogue
            // 
            this.openGlyphFileDialogue.DefaultExt = "xlsm";
            this.openGlyphFileDialogue.Filter = "Excel Files | *.xlsx|Excel Macro Enabled Files |*.xlsm";
            this.openGlyphFileDialogue.FilterIndex = 2;
            this.openGlyphFileDialogue.Title = "Decomposed Glyph File";
            // 
            // fontDialog1
            // 
            this.fontDialog1.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // FolderDialogue
            // 
            this.FolderDialogue.Description = "Select the input folder you want to analyse.";
            this.FolderDialogue.RootFolder = System.Environment.SpecialFolder.MyComputer;
            this.FolderDialogue.ShowNewFolderButton = false;
            // 
            // openContextCharFileDialogue
            // 
            this.openContextCharFileDialogue.DefaultExt = "xlsx";
            this.openContextCharFileDialogue.Filter = "Excel Files | *.xlsx|Excel Macro Enabled Files |*.xlsm";
            this.openContextCharFileDialogue.Title = "Context Character File";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(838, 724);
            this.Controls.Add(this.toolStripContainer2);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Count Glyphs";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.ContentPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericCharsAfter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericCharsBefore)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabFonts.ResumeLayout(false);
            this.tabFonts.PerformLayout();
            this.tabStyles.ResumeLayout(false);
            this.tabStyles.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listStyles)).EndInit();
            this.ErrorTab.ResumeLayout(false);
            this.ErrorTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listNormalisedErrors)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.IndivOrBulk.ResumeLayout(false);
            this.IndividualFile.ResumeLayout(false);
            this.IndividualFile.PerformLayout();
            this.Bulk.ResumeLayout(false);
            this.Bulk.PerformLayout();
            this.toolStripContainer2.ContentPanel.ResumeLayout(false);
            this.toolStripContainer2.ResumeLayout(false);
            this.toolStripContainer2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openInputDialogue;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox OutputFileBox;
        private System.Windows.Forms.Button BtnOutputFile;
        private System.Windows.Forms.Button BtnClose;
        private System.Windows.Forms.SaveFileDialog saveExcelDialogue;
        private System.Windows.Forms.Button BtnAnalyse;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripContainer toolStripContainer1;
        private System.Windows.Forms.ToolStripContainer toolStripContainer2;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem documentationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LicenseMenuItem;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckBox AnalyseByFont;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button BtnFontListFile;
        private System.Windows.Forms.TextBox FontListFileBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button BtnXMLFile;
        private System.Windows.Forms.TextBox XMLFileBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button BtnStyleListFile;
        private System.Windows.Forms.TextBox StyleListFileBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.SaveFileDialog saveXMLDialogue;
        private System.Windows.Forms.Button BtnSaveXML;
        private System.Windows.Forms.CheckBox CombDecomposedChars;
        private System.Windows.Forms.ToolTip toolTipCombine;
        private System.Windows.Forms.OpenFileDialog openGlyphFileDialogue;
        private System.Windows.Forms.Button BtnErrorList;
        private System.Windows.Forms.TextBox ErrorListBox;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label FileCounter;
        private System.Windows.Forms.CheckBox AggregateStats;
        private System.Windows.Forms.Button BtnSaveAggregateStats;
        private System.Windows.Forms.CheckBox WriteIndividualFile;
        private System.Windows.Forms.ToolStripMenuItem CombiningCharacters;
        private System.Windows.Forms.TabControl IndivOrBulk;
        private System.Windows.Forms.TabPage IndividualFile;
        private System.Windows.Forms.Button BtnGetInput;
        private System.Windows.Forms.TextBox InputFileBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabPage Bulk;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabFonts;
        private System.Windows.Forms.Button BtnSaveFontList;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox FontList;
        private System.Windows.Forms.Button BtnListFonts;
        private System.Windows.Forms.TabPage tabStyles;
        private System.Windows.Forms.Button BtnSaveStyles;
        internal System.Windows.Forms.DataGridView listStyles;
        private System.Windows.Forms.DataGridViewTextBoxColumn Style;
        private System.Windows.Forms.DataGridViewTextBoxColumn theDefaultFont;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button BtnGetStyles;
        private System.Windows.Forms.TabPage ErrorTab;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button BtnSaveErrorList;
        private System.Windows.Forms.DataGridView listNormalisedErrors;
        private System.Windows.Forms.DataGridViewTextBoxColumn MappedCharacter;
        private System.Windows.Forms.DataGridViewTextBoxColumn PossibleCharacter;
        private System.Windows.Forms.Button BtnGetEncoding;
        private System.Windows.Forms.TextBox EncodingTextBox;
        private System.Windows.Forms.Button BtnGetFont;
        private System.Windows.Forms.TextBox FontBox;
        private System.Windows.Forms.Label FontLabel;
        private System.Windows.Forms.Button BtnDecompGlyph;
        private System.Windows.Forms.TextBox DecompGlyphBox;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox OutputFileSuffixBox;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button BtnInputFolder;
        private System.Windows.Forms.TextBox InputFolderBox;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Button BtnCharStatFolder;
        private System.Windows.Forms.TextBox OutputFolderBox;
        private System.Windows.Forms.TextBox BulkFontListFileBox;
        private System.Windows.Forms.TextBox BulkStyleListBox;
        private System.Windows.Forms.Button BtnBulkErrorList;
        private System.Windows.Forms.TextBox BulkErrorListBox;
        private System.Windows.Forms.Button BtnBulkFontListFile;
        private System.Windows.Forms.Button BtnBulkStyleListFile;
        private System.Windows.Forms.Button BtnAggregateFile;
        private System.Windows.Forms.TextBox AggregateStatsBox;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button BtnSelectFiles;
        private System.Windows.Forms.FolderBrowserDialog FolderDialogue;
        private System.Windows.Forms.OpenFileDialog openContextCharFileDialogue;
        private System.Windows.Forms.Button BtnContextCharFile;
        private System.Windows.Forms.TextBox ContextCharacterFileBox;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Button BtnCheckContextFile;
        private System.Windows.Forms.NumericUpDown numericCharsAfter;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.NumericUpDown numericCharsBefore;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.TextBox boxContextChars;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.CheckBox checkGetContext;
        private System.Windows.Forms.CheckBox checkCountCharacters;
        private System.Windows.Forms.ToolStripMenuItem defaultsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem CountCharactersDefault;
        private System.Windows.Forms.ToolStripMenuItem AnalyseContextDefault;
    }
}

