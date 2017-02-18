/*
 *  This program counts characters in MS Word documents and writes the font, character code, glyph and count to an Excel workbook.
 *  It counts the characters in headers, footers, text boxes and other such containers as well as the main text.
 *  By storing the character information in a dictionary, we count the characters associated with each font in the document.
 *  It queries the XML representation of a Word document as this is much faster than using the Word object model and allows us to
 *  count the occurrences of characters entered using Insert Symbol.
 *  
 *  The copyright is owned by MissionAssist as the work was carried out on their behalf.
 * 
 *  Written by Stephen Palmstrom, last modified 6 February 2017
 *  
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.IO;
using System.Globalization;
using System.Diagnostics;
using System.Xml;
using System.Xml.XPath;
using System.Runtime.InteropServices;

using Microsoft.Win32;
using WordApp = Microsoft.Office.Interop.Word._Application;
using WordRoot = Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word._Document;
using ExcelApp = Microsoft.Office.Interop.Excel._Application;
using ExcelRoot = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel.Application;
using WorkBook = Microsoft.Office.Interop.Excel._Workbook;


namespace CharacterCounter
{
    public partial class Form1 : Form
    {
        // Some global variables
        private WordApp wrdApp;
        private ExcelApp excelApp;
        object missing = Type.Missing;
        private Document theDocument = null;
        private string InputDir = "";
        private string OutputDir = "";
        private string StyleDir = "";
        private string FontDir = "";
        private string XMLDir = "";
        private string GlyphDir = "";
        private string ContextDir = "";
        private string ErrorDir = "";
        private string AggregateDir = "";

        private string theFirstFont = "";

        private bool AggregateSaved = false;
        private bool Individual = true;
        private bool AnalyseContext = false;

        private int FileType = WordDoc;
        //private XmlDocument theXMLDocument = new XmlDocument();  // make it global so we can save it.
        //private string theTextDocument = "";  // This is to load a taxt document
        // needed for XML lookup
        const string wordmlNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";
        const string wordmlxNamespace = "http://schemas.microsoft.com/office/word/2003/auxHint";
        const string userRoot = "HKEY_CURRENT_USER";
        const string subkey = "Software\\MissionAssist\\CountCharacters";
        const string keyName = userRoot + "\\" + subkey;

        private Dictionary<string, string> theStyleDictionary = new Dictionary<string, string>(10); // to hold all defined styles
        private Dictionary<string, string> theDefaultStyleDictionary = new Dictionary<string, string>(5); // to hold all default styles
        private Dictionary<string, string> theBreakDictionary = new Dictionary<string, string>(3); // to hold the characters corresponding to breaks
        private Dictionary<string, string> theGlyphDictionary = new Dictionary<string, string>(5); // to hold the regular expressions for decomposed characters by font.
        private Dictionary<string, Encoding> theEncodingDictionary = new Dictionary<string, Encoding>(30); // to hold the encoding dictionary
        private Dictionary<CharacterDescriptor, int> theAggregateDictionary =
            new Dictionary<CharacterDescriptor, int>(1000, new CharacterEqualityComparer());
        private Dictionary<CharacterDescriptor, int> theAggregateSummaryDictionary =
            new Dictionary<CharacterDescriptor, int>(255, new CharacterEqualityComparer());
        private Dictionary<string, bool> theControlDictionary = new Dictionary<string, bool>(30); // To hold the state of the various controls
        //
        // Handle context analysis
        //
        private Dictionary<ContextDescriptor, int> theContextDictionary = 
            new Dictionary<ContextDescriptor, int>(100, new ContextEqualityComparer());
        private Dictionary<ContextDescriptor, int> theAggregateContextDictionary = 
            new Dictionary<ContextDescriptor, int>(1000, new ContextEqualityComparer());
        private Dictionary<ContextDescriptor, int> theAggregateContextSummaryDictionary =
            new Dictionary<ContextDescriptor, int>(255, new ContextEqualityComparer());
        private List<TargetDescriptor> TargetList = new List<TargetDescriptor>(5); // to hold a list of targets for which context is wanted.

        private bool Working = false;  // Flag if we are working

        private string AggregateFileList = "";


        private Dictionary<string, XmlDocument> theXMLDictionary = new Dictionary<string, XmlDocument>(10); // To hold the XML documents
        private Dictionary<string, string> theTextDictionary = new Dictionary<string, string>(10);  // to hold the text documents

        private string InputFileName = null;
        private bool GlyphsLoaded = false;
        private Encoding theEncoding = null;  // A place to store the encoding
        private XmlNamespaceManager nsManager;

        //
        //  Special characters
        //
        string[] SpecialCharacterKeys = { "w:endnoteRef", "w:endnote", "w:footnoteRef", "w:footnote", "w:tab", "w:noBreakHyphen",
                                            "w:softHyphen", "w:separator", "w:continuationSeparator"};
        string[] SpecialCharacterValues =
            {
                Convert.ToString("\x0002"),
                Convert.ToString("\x0002"),
                Convert.ToString("\x0002"),
                Convert.ToString("\x0002"),
                "\t",
                Convert.ToString("\x001E"),
                Convert.ToString("\x001F"),
                Convert.ToString("\x0003"),
                Convert.ToString("\x0004")
           };
        /*
         * Variables for handling Pause and Resume
         */
        //private bool CloseApp = false;

        /*
         * Character codes can be displayed as decimal, hexadecimal or USV
         */
        const int dec = 0;
        const int hex = 1;
        const int USV = 2;
        /*
         * File types
         */
        const int TextDoc = 0;
        const int WordDoc = 1;

        public Form1()
        {
            InitializeComponent();
            Application.ApplicationExit += new EventHandler(this.CloseApps);
            // Start Word and Excel

            wrdApp = new Word();
            // Turn off as much as possible.
            wrdApp.Visible = false;
            excelApp = new Excel();
            excelApp.Visible = false;
            /*
             * If the registry subkey doesn't exist, create it
             */
            if (Registry.CurrentUser.OpenSubKey(subkey, true) == null)
            {
                Registry.CurrentUser.CreateSubKey(subkey);
            }
            Registry.CurrentUser.Close(); // Close it
            //
            // Some registry settings
            //
            try
            {
                InputDir = GetDirectory("InputDir");
                OutputDir = GetDirectory("OutputDir", InputDir);
                StyleDir = GetDirectory("StyleDir", OutputDir);
                FontDir = GetDirectory("FontDir", OutputDir);
                XMLDir = GetDirectory("XMLDir", OutputDir);
                ErrorDir = GetDirectory("ErrorDir", OutputDir);
                GlyphDir = GetDirectory("GlyphDir", InputDir);
                ContextDir = GetDirectory("ContextDir", InputDir);
                AggregateDir = GetDirectory("AggregateDir", OutputDir);
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Failed to get directories", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseApps();
            }
            //
            //  Types of break
            //

            theBreakDictionary.Add("page", "\f");
            theBreakDictionary.Add("column", Convert.ToString("\x000E"));
            theBreakDictionary.Add("text-wrapping", "\v");



        }
        private string GetDirectory(string ValueName, string DefaultPath = "")
        {
            string theDirectory = "";
            if (DefaultPath == "")
            {
                DefaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            try
            {
                if (Registry.GetValue(keyName, ValueName, DefaultPath) != null)
                {
                    theDirectory = Registry.GetValue(keyName, ValueName, DefaultPath).ToString();
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace + "\rkeyName " + keyName + "\rValueName " + ValueName +
                "\rDefaultPath " + DefaultPath, "Can't read registry", MessageBoxButtons.OK);
                CloseApps();
            }

            return theDirectory;
        }






        private void btnGetInput_Click(object sender, EventArgs e)
        {
            Control theControl = (Control)sender;
            switch (theControl.Name)
            {
                case "btnGetInput":
                    openInputDialogue.Multiselect = false;
                    openInputDialogue.InitialDirectory = InputDir;
                    if (openInputDialogue.ShowDialog() == DialogResult.OK)
                    {
                        InputFileBox.Text = openInputDialogue.FileName;
                    }
                    break;
                case "btnDecompGlyph":
                    OpenGlyphFileDialogue.InitialDirectory = GlyphDir;
                    if (OpenGlyphFileDialogue.ShowDialog() == DialogResult.OK)
                    {
                        DecompGlyphBox.Text = OpenGlyphFileDialogue.FileName;
                        GlyphDir = Path.GetDirectoryName(OpenGlyphFileDialogue.FileName);
                        Registry.SetValue(keyName, "GlyphDir", GlyphDir);
                    }
                    break;
                case "btnContextCharFile":
                    openContextCharFileDialogue.InitialDirectory = ContextDir;
                    if (openContextCharFileDialogue.ShowDialog() == DialogResult.OK)
                    {
                        ContextCharacterFileBox.Text = openContextCharFileDialogue.FileName;
                        ContextDir = Path.GetDirectoryName(openContextCharFileDialogue.FileName);
                        Registry.SetValue(keyName, "ContextDir", ContextDir);
                    }
                    break;

            }
        }
        private int GetFileType(string FileName, bool JustType = true)
        {
            switch (Path.GetExtension(FileName).ToLower())
            {
                // There may be times we just want the file type and nothing else.
                case ".doc":
                case ".docx":
                case ".rtf":
                    if (!JustType)
                    {
                        btnListFonts.Enabled = true;
                        btnGetStyles.Enabled = true;
                        btnSaveFontList.Enabled = true && (Individual || FontListFileBox.Text != "");
                        btnSaveStyles.Enabled = true && (Individual || StyleListFileBox.Text != "");
                        btnSaveXML.Enabled = true;
                        btnGetFont.Enabled = false && !Individual;
                        btnGetEncoding.Enabled = false && !Individual;
                        if (Individual)
                        {
                            AnalyseByFont.Enabled = true;  // Don't turn on if we are doing a bulk analysis
                        }
                        FontLabel.Enabled = false;
                        FontBox.Text = "";
                    }
                    return WordDoc;
                default:
                    if (!JustType)
                    {
                        btnGetFont.Enabled = true;
                        btnGetEncoding.Enabled = true;
                        if (Individual)
                        {
                            AnalyseByFont.Checked = false;  // Don't turn off if we are doing a bulk analysis.
                        }
                        AnalyseByFont.Enabled = false;
                        btnListFonts.Enabled = false;
                        btnSaveXML.Enabled = false;
                        btnSaveFontList.Enabled = false;
                        btnGetStyles.Enabled = false;
                        btnSaveStyles.Enabled = false;
                        FontLabel.Enabled = true;
                        if (FontBox.Text == "")
                        {
                            FontBox.Text = Registry.GetValue(keyName, "Font", "Calibri").ToString();
                        }
                        if (EncodingTextBox.Text == "")
                        {
                            EncodingTextBox.Text = Registry.GetValue(keyName, "Encoding", "Western European (Windows)").ToString();
                        }
                    }
                    return TextDoc;

            }


        }

        private void btnGetOutput_Click(object sender, EventArgs e)
        {
            //
            // Handle the output files
            //
            Control theControl = (Control)sender;  // cast the sender as a control.
            SaveFileDialog theDialogue = null;
            string theDirectory = "";
            string ValueName = "";
            TextBox theTextBox = null;
            Button theButton = null;
            switch (theControl.Name)
            {
                case "btnOutputFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = OutputFileBox;
                    theDialogue.InitialDirectory = OutputDir;
                    theButton = btnAnalyse;
                    ValueName = "OutputDir";
                    break;
                case "btnStyleListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = StyleListFileBox;
                    theDialogue.InitialDirectory = StyleDir;
                    theButton = btnSaveStyles;
                    ValueName = "StyleDir";
                    break;
                case "btnBulkStyleListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = BulkStyleListBox;
                    theDialogue.InitialDirectory = StyleDir;
                    theButton = btnSaveStyles;
                    ValueName = "StyleDir";
                    break;
                case "btnBulkFontListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = BulkFontListFileBox;
                    theDialogue.InitialDirectory = FontDir;
                    theButton = btnSaveFontList;
                    ValueName = "FontDir";
                    break;
                case "btnFontListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = FontListFileBox;
                    theDialogue.InitialDirectory = FontDir;
                    theButton = btnSaveFontList;
                    ValueName = "FontDir";
                    break;
                case "btnXMLFile":
                    theDialogue = saveXMLDialogue;
                    theTextBox = XMLFileBox;
                    theDialogue.InitialDirectory = XMLDir;
                    theButton = btnSaveXML;
                    ValueName = "XMLDir";
                    break;
                case "btnErrorList":
                    theDialogue = saveExcelDialogue;
                    theTextBox = ErrorListBox;
                    theDialogue.InitialDirectory = ErrorDir;
                    theButton = btnSaveErrorList;
                    ValueName = "ErrorDir";
                    break;
                case "btnBulkErrorList":
                    theDialogue = saveExcelDialogue;
                    theTextBox = BulkErrorListBox;
                    theDialogue.InitialDirectory = ErrorDir;
                    theButton = btnSaveErrorList;
                    ValueName = "ErrorDir";
                    break;

                case "btnAggregateFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = AggregateStatsBox;
                    theDialogue.InitialDirectory = AggregateDir;
                    theButton = btnAggregateFile;
                    ValueName = "AggregateDir";
                    break;

            }

            theDialogue.FileName = theTextBox.Text;
            if (theDialogue.ShowDialog() == DialogResult.OK)
            {
                theTextBox.Text = theDialogue.FileName;
                theDirectory = Path.GetDirectoryName(theDialogue.FileName);
                Registry.SetValue(keyName, ValueName, theDirectory);
                if (theButton != null)
                {
                    theButton.Enabled = true;
                }
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            //CloseApp = true;
            Application.DoEvents();
            this.Close();
            Application.Exit();
        }
        private void CloseApps(object sender = null, EventArgs e = null)
        {
            toolStripStatusLabel1.Text = "Shutting down...";
            Working = MarkWorking(false, Working, theControlDictionary);
            Application.DoEvents();
            // Close Excel and Word, but don't flag an error if they are already closed.
            try
            {
                //NAR(wrdApp);  // release any objects like documents to make sure we can quit.
                wrdApp.Quit();
                wrdApp = null;
            }
            catch
            {
            }
            try
            {
                if (excelApp != null)
                {
                    // Close any open workbooks without saving
                    try
                    {

                        foreach (WorkBook theWorkBook in excelApp.Workbooks)
                        {
                            theWorkBook.Close(false);  // Close without saving
                        }

                    }
                    catch
                    { }

                }


                excelApp.Quit();
                NAR(excelApp);  // release any objects like workbooks because Excel doesn't always quit.
                System.Threading.Thread.Sleep(5000); // and sleep five seconds
                excelApp.Quit(); // try again
                excelApp = null;

            }
            catch
            {
            }
            this.Close();

        }
        private void NAR(object o)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0) ;
            }
            catch { }
            finally
            {
                o = null;  // clear the object
            }
        }


        private void btnAnalyse_Click(object sender, EventArgs e)
        {
            /*
             * Here is where we start the analysis
             * 
             * The dictionary holds the character counts.  Its key is the character, including details of the associated font.
             * This means that we can output the counts for a given code for each associated font and therefore glyph.
             * 
             * The CharacterDescriptor class holds both the font and text information.
             */
            Button theButton = (Button)sender;
            Working = MarkWorking(true, Working, theControlDictionary); // Mark as waiting
            try
            {
                Stopwatch theStopwatch = new Stopwatch();
                listNormalisedErrors.Rows.Clear();  // Reset the normalised errors list
                // Get the context targets if the file is there and we haven't loaded it already
                if (AnalyseContext && TargetList.Count == 0)
                {
                    LoadTargets(TargetList, excelApp);
                }
                foreach (string theFile in openInputDialogue.FileNames)
                {
                    Dictionary<CharacterDescriptor, int> theCharDictionary =
                            new Dictionary<CharacterDescriptor, int>(500, new CharacterEqualityComparer());
                    toolStripProgressBar1.Value = 0;
                    // This is so we can remember the textboxes as we find them.
                    TimeSpan TimeToFinish = TimeSpan.Zero;
                    toolStripStatusLabel1.Text = "Analysing " + theFile + "...";
                    Application.DoEvents();
                    /*
                     * Open the Word file
                     */
                    //EnableButtons(false, FileType); // Disable a whole lot of buttons
                    //lngLink = theDocument.Sections[1].Headers[theIndex].Range.StoryType;
                    int CharacterCount = 0;  // Counter to measure progress
                    // Count the total characters
                    toolStripStatusLabel1.Text = theFile + " opened... ";
                    theFirstFont = "";
                    toolStripProgressBar1.Value = 0;
                    Application.DoEvents();
                    /*
                     * We now load the document and analyse it
                     */
                    string theFileName = Path.GetFileName(theFile);
                    bool Success = false;  // Be pessimistic!
                    switch (GetFileType(theFile))
                    {
                        case WordDoc:

                            if (DocumentLoader(theFile, theXMLDictionary, ref nsManager))
                            {
                                Success = true;
                                CharacterCount = AnalyseDocument(theCharDictionary, theContextDictionary, theXMLDictionary[theFileName], ref theFirstFont, nsManager, CharacterCount,
                                    theStopwatch);
                            }
                            break;
                        default:
                            if (DocumentLoader(theFile, theTextDictionary))
                            {
                                Success = true;
                                CharacterCount = AnalyseText(theCharDictionary, theContextDictionary, theTextDictionary[theFileName], FontBox.Text, CharacterCount, theStopwatch);
                                theFirstFont = FontBox.Text;
                            }
                            break;
                    }
                    if (Success)
                    {
                        /*
                         * Load the aggregate file statistics dictionary
                         */
                        if (AggregateStats.Checked)
                        {
                            AggregateFileList = LoadAggregateStats(theAggregateDictionary, theAggregateSummaryDictionary, theCharDictionary, 
                                theAggregateContextDictionary, theAggregateContextSummaryDictionary, theContextDictionary,
                                AggregateFileList, theFileName);
                            AggregateSaved = false;  // the list has changed.
                            theControlDictionary[btnSaveAggregateStats.Name] = !AggregateSaved && AggregateStatsBox.Text != "" && FileCounter.Text != "0";  // Enable when we finish working and have counted some files
                        }

                        /*
                          * Create the Excel worksheet
                          */
                        toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                        if (WriteIndividualFile.Checked)
                        {
                            string OutputFile = "";
                            if (Individual)
                            {
                                OutputFile = OutputFileBox.Text;
                            }
                            else
                            {
                                OutputFile = Path.Combine(OutputDir, Path.GetFileNameWithoutExtension(theFile) +
                                OutputFileSuffixBox.Text + ".xlsx");
                            }
                            WriteOutput(theCharDictionary, theContextDictionary, theFirstFont, OutputDir, OutputFile, theFile);
                        }
                        Working = MarkWorking(false, Working, theControlDictionary);
                        EnableButtons(true, FileType);
                        btnSaveErrorList.Enabled = (listNormalisedErrors.Rows.Count > 0) && ((Individual && ErrorListBox.Text != "") || (!Individual && BulkErrorListBox.Text != ""));
                        theCharDictionary = null;  // release it.
                    }
                }
                if (!Individual && AggregateStats.Checked)
                {
                    btnSaveAggregateStats_Click(sender, e);  // Pretend we clicked the Save Aggregate Stats button
                }
                theStopwatch.Stop();
                toolStripStatusLabel1.Text = "Finished in " + theStopwatch.Elapsed.ToString(@"hh\:mm\:ss");
                AnalyseByFont.Enabled = true;
                CombDecomposedChars.Enabled = true;
                toolStripProgressBar1.Value = 0;
                Registry.SetValue(keyName, "OutputDir", OutputDir);
                System.Media.SystemSounds.Beep.Play();  // and beep

            }
            catch (Exception theException)
            {
                // Catch any unexpected errors
                MessageBox.Show(theException.Message + theException.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseApps(this);
            }
            theButton.Enabled = false;  // disable so you can't analyse the same file twice by mistake.


        }
        private void EnableButtons(bool Enable, int theFileType)
        {
            // Disable or enable a number of buttons
            //btnAnalyse.Enabled = Enable;
            //btnErrorList.Enabled = Enable;
            btnSaveErrorList.Enabled = Enable && ((Individual && ErrorListBox.Text != "") || (!Individual && BulkErrorListBox.Text != ""));
            //CombDecomposedChars.Enabled = Enable;
            //AnalyseByFont.Enabled = Enable;
            btnDecompGlyph.Enabled = Enable && CombDecomposedChars.Checked;
            if (theFileType == WordDoc)
            {
                // These only apply to Word documents
                btnOutputFile.Enabled = Enable;
                btnGetInput.Enabled = Enable;
                btnDecompGlyph.Enabled = Enable;
                btnFontListFile.Enabled = Enable;
                /*
                 * Only enable the save buttons if there the file has been specified.
                 */
                btnSaveFontList.Enabled = Enable && ((Individual && FontListFileBox.Text != "") || (!Individual && BulkFontListFileBox.Text != ""));
                btnStyleListFile.Enabled = Enable;
                btnSaveStyles.Enabled = Enable && ((Individual && StyleListFileBox.Text != "") || (!Individual && BulkStyleListBox.Text != ""));
                btnXMLFile.Enabled = Enable && Individual && XMLFileBox.Text != "";
                btnListFonts.Enabled = Enable;
                btnGetStyles.Enabled = Enable;
            }
        }

        private bool WriteOutput(Dictionary<CharacterDescriptor, int> theCharDictionary, Dictionary<ContextDescriptor, int> theContextDictionary, string theFont, string OutputDir, string OutputFile,
            string InputFile, Dictionary<CharacterDescriptor, int> theSummaryCharDictionary = null, Dictionary<ContextDescriptor, int> theSummaryContextDictionary = null)
        {
            Stopwatch theStopwatch = new Stopwatch();
            theStopwatch.Start();
            // We write the results to Excel.
            Application.DoEvents();
            if (!Path.IsPathRooted(OutputFile))
            {
                OutputFile = Path.Combine(OutputDir, OutputFile);  // Make it a complete directory
            }
            toolStripStatusLabel1.Text = "Writing data to Excel workbook " + Path.GetFileName(OutputFile) + "...";
            Application.DoEvents();
            if (!DeleteFile(OutputFile))
            {
                return false; // Do no more if not successfully deleted.
            }
            bool Retry = true;
            ExcelRoot.Workbook theWorkbook = null;
            while (Retry)
            {
                try
                {
                    theWorkbook = excelApp.Workbooks.Add();  // Create it
                    Retry = false;
                }
                catch (COMException ComEx)
                {
                    if (ComEx.ErrorCode == -2147023174)  // RPC Exception - we've lost Excel
                    {
                        excelApp = new Excel();  // Recreate it
                        Retry = true;
                    }
                    else
                    {
                        MessageBox.Show(ComEx.Message + "\r" + ComEx.StackTrace, "Failed to open Excel", MessageBoxButtons.OK);
                        CloseApps();
                        return false;
                    }
                }
            }

            Working = MarkWorking(true, Working, theControlDictionary); // Disable buttons and textboxes

            ExcelRoot.Worksheet theSheet = theWorkbook.ActiveSheet;
            // Write the first output sheet.
            WriteOutputSheet(theCharDictionary, theFont, theSheet, "Statistics", theSummaryCharDictionary != null);  //

            if (theSummaryCharDictionary != null)  // We have aggregate results
            {
                theSheet = theWorkbook.Sheets.Add(missing, theSheet, 1, ExcelRoot.XlSheetType.xlWorksheet);  // add the summary sheet
                toolStripStatusLabel1.Text = "Writing summary statistics to Excel workbook " + Path.GetFileName(OutputFile) + "...";
                WriteOutputSheet(theSummaryCharDictionary, theFont, theSheet, "Summary Statistics", false);  //

            }
            // Now look at context

            if (AnalyseContext)
            {
                theSheet = theWorkbook.Sheets.Add(missing, theSheet, 1, ExcelRoot.XlSheetType.xlWorksheet);
                WriteContextOutputSheet(theContextDictionary, theFont, theSheet, "Context", theSummaryContextDictionary != null);
                if (theSummaryContextDictionary != null)  // We have aggregate results
                {
                    theSheet = theWorkbook.Sheets.Add(missing, theSheet, 1, ExcelRoot.XlSheetType.xlWorksheet);  // add the summary sheet
                    toolStripStatusLabel1.Text = "Writing context summary statistics to Excel workbook " + Path.GetFileName(OutputFile) + "...";
                    WriteContextOutputSheet(theSummaryContextDictionary, theFont, theSheet, "Summary Context Statistics", false);  //

                }

            }
            theWorkbook.SaveAs(OutputFile);
            //
            //Create a new worksheet for the metadata and write to it
            //
            theSheet = theWorkbook.Sheets.Add(missing, theSheet, 1, ExcelRoot.XlSheetType.xlWorksheet);
            theSheet.Name = "MetaData";
            theSheet.Range["A1"].Value = "Character Counter Version";
            theSheet.Range["B1"].Value = String.Format("{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
            theSheet.Range["A2"].Value = "Filename(s)";
            theSheet.Range["B2"].Value = InputFile;
            if (FileType == TextDoc)
            {
                theSheet.Range["A3"].Value = "Encoding";
                theSheet.Range["B3"].Value = EncodingTextBox.Text;
            }
            theSheet.Columns["A"].ColumnWidth = 25;
            theWorkbook.Sheets["Statistics"].Activate();  // go to the statistics sheet

            // and save it
            theWorkbook.Save();
            theWorkbook.Close();
            theStopwatch.Stop();
            //System.Media.SystemSounds.Beep.Play();  // and beep
            toolStripStatusLabel1.Text = "Finished writing to Excel workbook " + Path.GetFileName(OutputFile) + " in " +
                ((float)theStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            toolStripProgressBar1.Value = 0; // reset
            Working = MarkWorking(false, Working, theControlDictionary);
            return true;

        }
        private void WriteOutputSheet(Dictionary<CharacterDescriptor, int> theDictionary, string theFont, ExcelRoot.Worksheet theSheet, string SheetName, bool Aggregate)
        {

            theSheet.Name = SheetName;

            /*
            * Column headers
            */
            int Column = 1;
            string FontColumnLetter = "A";
            string ValueColumnLetter = "B";
            string ColumnLetter = "E";
            string GlyphLetter = "D";
            int FontNumber = (int)Convert.ToChar(FontColumnLetter);
            int ValueNumber = (int)Convert.ToChar(ValueColumnLetter);
            int ColumnNumber = (int)Convert.ToChar(ColumnLetter);
            int GlyphNumber = (int)Convert.ToChar(GlyphLetter);
            toolStripProgressBar1.Value = 0; // reset
            toolStripProgressBar1.Maximum = theDictionary.Count;
            if (Aggregate)
            {
                // we are writing an aggregate file and will need a summary worksheet, too.
                theSheet.Cells[1, Column++].Value = "Filename";
                // Increment the column numbers for glyph and the final column
                ValueNumber++;
                FontNumber++;
                ColumnNumber++;
                GlyphNumber++;
            }
            if (AnalyseByFont.Checked)
            {
                theSheet.Cells[1, Column++].Value = "Font";
                ValueNumber++;
                ColumnNumber++;
                GlyphNumber++;
            }
            ValueColumnLetter = Convert.ToChar(ValueNumber).ToString();
            ColumnLetter = Convert.ToChar(ColumnNumber).ToString();
            GlyphLetter = Convert.ToChar(GlyphNumber).ToString();
            FontColumnLetter = Convert.ToChar(FontNumber).ToString();
            string RangeString = "A1:" + ColumnLetter + "1";

            theSheet.Cells[1, Column++].Value = "Dec";
            theSheet.Cells[1, Column++].value = "MS Hex";
            theSheet.Cells[1, Column++].Value = "USV";
            theSheet.Cells[1, Column++].Value = "Glyph";
            theSheet.Cells[1, Column].Value = "Count";
            theSheet.Range[RangeString].Font.Bold = true;  // Make the headings bold.
            theSheet.Range[RangeString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;  // Centre

            int theRow = 2;
            foreach (KeyValuePair<CharacterDescriptor, int> kvp in theDictionary)
            {
                /*
                    * Go through the dictionary writing to the worksheet
                    */
                Column = 1;
                if (Aggregate)
                {
                    theSheet.Cells[theRow, Column++].Value = kvp.Key.FileName;
                }
                if (AnalyseByFont.Checked)
                {
                    theSheet.Cells[theRow, Column++].Value = kvp.Key.Font;
                }
                theSheet.Cells[theRow, Column].Value = GetCodes(kvp.Key.Text, dec);  // Decimal
                theSheet.Cells[theRow, Column++].NumberFormat = "@";  // Text format
                theSheet.Cells[theRow, Column].NumberFormat = "@";
                theSheet.Cells[theRow, Column++].Value = GetCodes(kvp.Key.Text, hex);  // Hexadecimal
                theSheet.Cells[theRow, Column++].Value = GetCodes(kvp.Key.Text, USV);  // USV
                if (AnalyseByFont.Checked)
                {
                    theSheet.Cells[theRow, Column].Font.Name = kvp.Key.Font;  // and set the font.
                }
                else
                {
                    theSheet.Cells[theRow, Column].Font.Name = theFont;  // Set the cell to the first font we found
                }

                theSheet.Cells[theRow, Column++].Value = kvp.Key.Text;  // Write the glyph
                theSheet.Cells[theRow, Column].Value = kvp.Value;  // the count
                theRow++;
                toolStripProgressBar1.Value = theRow - 2;
                Application.DoEvents();
            }
            string theRowString = (theRow - 1).ToString();
            /*
                * Now sort by USV value and, if analysing by font, font name
                */

            ExcelRoot.Range CharStats = excelApp.get_Range("A1", ColumnLetter + theRowString);
            // Now format all cells
            if (AnalyseByFont.Checked)
            {
                CharStats.Sort(CharStats.Columns[4], ExcelRoot.XlSortOrder.xlAscending,
                    CharStats.Columns[1], missing, ExcelRoot.XlSortOrder.xlAscending,
                    missing, ExcelRoot.XlSortOrder.xlAscending, ExcelRoot.XlYesNoGuess.xlYes);
                string FontRange = FontColumnLetter + ":" + FontColumnLetter;
                theSheet.Range[FontRange].EntireColumn.ColumnWidth = 30;  // allow for 30 character font names
                theSheet.Range[FontRange].EntireColumn.HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            }
            else
            {
                CharStats.Sort(CharStats.Columns[3], ExcelRoot.XlSortOrder.xlAscending,
                    missing, missing, ExcelRoot.XlSortOrder.xlAscending,
                    missing, ExcelRoot.XlSortOrder.xlAscending, ExcelRoot.XlYesNoGuess.xlYes);

            }
            theSheet.Range[ValueColumnLetter + "1:" + ColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            theSheet.Range["A1:" + ColumnLetter + theRowString].VerticalAlignment = ExcelRoot.XlVAlign.xlVAlignBottom;
            // and the counts
            theSheet.Range[GlyphLetter + "2:" + GlyphLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;
            theSheet.Range[ColumnLetter + "2:" + ColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignRight;
            theSheet.Range[ColumnLetter + "2:" + ColumnLetter + theRowString].NumberFormat = "#,##0";
            // and return to A1
            theSheet.Range["A1"].Select();
            // and freeze the top row
            excelApp.ActiveWindow.SplitColumn = 0;
            excelApp.ActiveWindow.SplitRow = 1;
            excelApp.ActiveWindow.FreezePanes = true;
            toolStripProgressBar1.Value = 0;
        }
        private void WriteContextOutputSheet(Dictionary<ContextDescriptor, int> theDictionary, string theFont, ExcelRoot.Worksheet theSheet, string SheetName, bool Aggregate)
        {

            theSheet.Name = SheetName;

            /*
            * Column headers
            */
            int Column = 1;
            string FontColumnLetter = "A";
            string TargetUSVColumnLetter = "A";
            string ContextUSVColumnLetter = "B";
            string TargetLetter = "C";
            string ContextColumnLetter = "D";
            string CountColumnLetter = "E";

            int FontNumber = (int)Convert.ToChar(FontColumnLetter);
            int TargetUSVNumber = (int)Convert.ToChar(TargetUSVColumnLetter);
            int ContextUSVColumnNumber = (int)Convert.ToChar(ContextUSVColumnLetter);
            int ContextColumnNumber = (int)Convert.ToChar(ContextColumnLetter);
            int CountColumnNumber = (int)Convert.ToChar(CountColumnLetter);
            int TargetNumber = (int)Convert.ToChar(TargetLetter);
            toolStripProgressBar1.Value = 0; // reset
            toolStripProgressBar1.Maximum = theDictionary.Count;
            if (Aggregate)
            {
                // we are writing an aggregate file and will need a summary worksheet, too.
                theSheet.Cells[1, Column++].Value = "Filename";
                // Increment the column numbers for glyph and the final column
                TargetUSVNumber++;
                ContextUSVColumnNumber++;
                FontNumber++;
                ContextColumnNumber++;
                CountColumnNumber++;
                TargetNumber++;
            }
            if (AnalyseByFont.Checked)
            {
                theSheet.Cells[1, Column++].Value = "Font";
                TargetUSVNumber++;
                ContextColumnNumber++;
                ContextUSVColumnNumber++;
                CountColumnNumber++;
                TargetNumber++;
            }
            TargetUSVColumnLetter = Convert.ToChar(TargetUSVNumber).ToString();
            ContextColumnLetter = Convert.ToChar(ContextColumnNumber).ToString();
            ContextUSVColumnLetter = Convert.ToChar(ContextUSVColumnNumber).ToString();
            CountColumnLetter = Convert.ToChar(CountColumnNumber).ToString();
            TargetLetter = Convert.ToChar(TargetNumber).ToString();
            FontColumnLetter = Convert.ToChar(FontNumber).ToString();


            theSheet.Cells[1, Column++].Value = "Target (USV)";
            theSheet.Cells[1, Column++].value = "Context(USV)";
            theSheet.Cells[1, Column++].Value = "Target (Glyph)";
            theSheet.Cells[1, Column++].Value = "Context";
            theSheet.Cells[1, Column].Value = "Count";
            // Now format some cells
            string theRowString = (theDictionary.Count + 1).ToString();
            theSheet.Range[TargetUSVColumnLetter + "1:" + ContextUSVColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            theSheet.Range["A1:" + ContextColumnLetter + theRowString].VerticalAlignment = ExcelRoot.XlVAlign.xlVAlignBottom;
            // and the counts
            theSheet.Range[TargetLetter + "2:" + TargetLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;
            theSheet.Range[ContextColumnLetter + "2:" + ContextColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            theSheet.Range[CountColumnLetter + "2:" + CountColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignRight;
            theSheet.Range[CountColumnLetter + "2:" + CountColumnLetter + theRowString].NumberFormat = "#,##0";
            // Align the top row
            string RangeString = "A1:" + CountColumnLetter + "1";
            theSheet.Range[RangeString].Font.Bold = true;  // Make the headings bold.
            theSheet.Range[RangeString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;  // Centre
            theSheet.Range[RangeString].VerticalAlignment = ExcelRoot.XlVAlign.xlVAlignTop; // Align at the top
            theSheet.Range[RangeString].WrapText = true;
            // Size some column widths
            theSheet.Range[ContextColumnLetter + ":" + ContextColumnLetter].EntireColumn.ColumnWidth = 30;  // allow for 30 character font names
            theSheet.Range[ContextUSVColumnLetter + ":" + ContextUSVColumnLetter].EntireColumn.ColumnWidth = 90;  // allow for 30 character font names

            TargetDescriptor theTargetDescriptor = new TargetDescriptor();  // To expose its functions
            int theRow = 2;
            foreach (KeyValuePair<ContextDescriptor, int> kvp in theDictionary)
            {
                /*
                    * Go through the dictionary writing to the worksheet
                    */
                Column = 1;
                if (Aggregate)
                {
                    theSheet.Cells[theRow, Column++].Value = kvp.Key.FileName;
                }
                if (AnalyseByFont.Checked)
                {
                    theSheet.Cells[theRow, Column++].Value = kvp.Key.Font;
                }
                theSheet.Cells[theRow, Column++].Value = theTargetDescriptor.GetHex(kvp.Key.Target, "U+", " "); // Target USV
                theSheet.Cells[theRow, Column++].Value = theTargetDescriptor.GetHex(kvp.Key.Context, "U+", " "); // Context USV

                theSheet.Cells[theRow, Column++].Value = kvp.Key.Target;  // Write the glyph

                theSheet.Cells[theRow, Column++].Value = kvp.Key.Context;  // Write the context as glyphs
                if (AnalyseByFont.Checked)
                {
                    theSheet.Cells[theRow, Column - 2].Font.Name = kvp.Key.Font;  // and set the font.
                    theSheet.Cells[theRow, Column - 1].Font.Name = kvp.Key.Font;  // and set the font.
                }
                else
                {
                    theSheet.Cells[theRow, Column - 2].Font.Name = theFont;  // Set the cell to the first font we found
                    theSheet.Cells[theRow, Column - 1].Font.Name = theFont;  // Set the cell to the first font we found
                }
                theSheet.Cells[theRow, Column++].Value = kvp.Value;  // the count
                theRow++;
                toolStripProgressBar1.Value = theRow - 2;
                Application.DoEvents();
            }

            /*
                * Now sort by USV value and, if analysing by font, font name
                */

            ExcelRoot.Range CharStats = excelApp.get_Range("A1", ContextColumnLetter + theRowString);
            // Now format all cells
            if (AnalyseByFont.Checked)
            {
                CharStats.Sort(CharStats.Columns[4], ExcelRoot.XlSortOrder.xlAscending,
                    CharStats.Columns[1], missing, ExcelRoot.XlSortOrder.xlAscending,
                    missing, ExcelRoot.XlSortOrder.xlAscending, ExcelRoot.XlYesNoGuess.xlYes);
                string FontRange = FontColumnLetter + ":" + FontColumnLetter;
                theSheet.Range[FontRange].EntireColumn.ColumnWidth = 30;  // allow for 30 character font names
                theSheet.Range[FontRange].EntireColumn.HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            }
            else
            {
                CharStats.Sort(CharStats.Columns[3], ExcelRoot.XlSortOrder.xlAscending,
                    missing, missing, ExcelRoot.XlSortOrder.xlAscending,
                    missing, ExcelRoot.XlSortOrder.xlAscending, ExcelRoot.XlYesNoGuess.xlYes);

            }


            // and return to A1
            theSheet.Range["A1"].Select();
            // and freeze the top row
            excelApp.ActiveWindow.SplitColumn = 0;
            excelApp.ActiveWindow.SplitRow = 1;
            excelApp.ActiveWindow.FreezePanes = true;
            toolStripProgressBar1.Value = 0;
        }
        private bool DeleteFile(string theFileName)
        {
            DialogResult Success = DialogResult.Retry;
            while (Success == DialogResult.Retry)
            {
                try
                {
                    File.Delete(theFileName); // Delete the existing file
                    Success = DialogResult.OK;  // we succeeded
                }
                catch (Exception Ex)
                {
                    Success = MessageBox.Show(Ex.Message, "Failed to delete Excel file", MessageBoxButtons.RetryCancel);
                    if (Success == DialogResult.Cancel)
                    {
                        return false; // Don't try to save to Excel
                    }
                }
            }
            return true;
        }
        private void WriteFontList(string theFileName, string theHeader, ListBox theListBox)
        {
            Working = MarkWorking(true, Working, theControlDictionary);
            int RowCounter = 2;
            // Write the contents of a list box to Excel
            if (!DeleteFile(theFileName))
            {
                return;
            }
            toolStripStatusLabel1.Text = "Writing to " + Path.GetFileName(theFileName);
            toolStripProgressBar1.Maximum = theListBox.Items.Count;
            toolStripProgressBar1.Value = 0;
            ExcelRoot.Workbook theWorkBook = excelApp.Workbooks.Add();
            ExcelRoot.Worksheet theSheet = theWorkBook.ActiveSheet;
            theSheet.Range["A1"].Value = theHeader;
            theSheet.Range["A1"].Font.Bold = true;
            foreach (string theItem in theListBox.Items)
            {
                theSheet.Range["A" + RowCounter.ToString()].Value = theItem;
                RowCounter++;
            }
            theWorkBook.SaveAs(theFileName);
            theWorkBook.Close();
            toolStripStatusLabel1.Text = "Finished writing to " + Path.GetFileName(theFileName);
            toolStripProgressBar1.Value = 0;
            Working = MarkWorking(false, Working, theControlDictionary);
            System.Media.SystemSounds.Beep.Play();  // and beep
            return;
        }
        private void WriteDataGridView(string theFileName, DataGridView theDataGridView)
        {
            Working = MarkWorking(true, Working, theControlDictionary);
            int RowCounter = 2;
            // Write the contents of a list box to Excel
            if (!DeleteFile(theFileName))
            {
                return;
            }
            toolStripStatusLabel1.Text = "Writing to " + Path.GetFileName(theFileName);
            toolStripProgressBar1.Maximum = theDataGridView.Rows.Count;
            toolStripProgressBar1.Value = 0;
            ExcelRoot.Workbook theWorkBook = excelApp.Workbooks.Add();
            ExcelRoot.Worksheet theSheet = theWorkBook.ActiveSheet;
            theSheet.Range["A1"].Value = theDataGridView.Columns[0].HeaderText;
            theSheet.Range["B1"].Value = theDataGridView.Columns[1].HeaderText;
            theSheet.Range["A1:B1"].Font.Bold = true;
            theSheet.Range["A1:B1"].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;
            theSheet.Range["A:B"].EntireColumn.ColumnWidth = 100;

            foreach (DataGridViewRow theRow in theDataGridView.Rows)
            {

                theSheet.Range["A" + RowCounter.ToString()].Value = theRow.Cells[0].Value;
                theSheet.Range["B" + RowCounter.ToString()].Value = theRow.Cells[1].Value;
                toolStripProgressBar1.Value++;
                Application.DoEvents();
                RowCounter++;
            }
            theWorkBook.SaveAs(theFileName);
            theWorkBook.Close();
            toolStripStatusLabel1.Text = "Finished writing to " + Path.GetFileName(theFileName);
            toolStripProgressBar1.Value = 0;
            System.Media.SystemSounds.Beep.Play();  // and beep
            Working = MarkWorking(false, Working, theControlDictionary);
            return;
        }
        private int AnalyseText(Dictionary<CharacterDescriptor, int> theCharacterDictionary, Dictionary<ContextDescriptor, int> theContextDictionary,
            string theText, string theFont, int CharacterCount, Stopwatch theStopwatch)
        {
            // Analyse a text document
            int CharactersInText = theText.Length;
            toolStripStatusLabel1.Text = "Counting characters";
            CharacterCount += AnalyseString(theCharacterDictionary, theContextDictionary,
                theFont, theText, CharactersInText, CharacterCount, theStopwatch);
            listNormalisedErrors.Sort(listNormalisedErrors.Columns[0], ListSortDirection.Ascending);  // Sort first
            return CharacterCount;


        }
        private int AnalyseDocument(Dictionary<CharacterDescriptor, int> theCharacterDictionary, Dictionary<ContextDescriptor, int> theContextDictionary,
            XmlDocument theXMLDocument, ref string theFirstFont, XmlNamespaceManager nsManager, int CharacterCount, Stopwatch theStopwatch)
        {
            // Analyse the contents of a character string
            // We repeat ourselves to avoid having to do a logic test through each iteration of the loop.
            string TextString = "";
            string FontName = "";
            int RangeCharacterCount = 0;
            int TextCount = 0;  // Count separately for troubleshooting purposes
            int OtherCount = 0;
            theCharacterDictionary.Clear();  // Clear the dictionary
            theContextDictionary.Clear();
            //
            //  Count the characters
            //
            toolStripStatusLabel1.Text = "Counting characters";
            XmlNode theRoot = theXMLDocument.DocumentElement;
            XmlNodeList theNodeList = theRoot.SelectNodes(@"(//w:body//w:r/w:t | //w:body//w:r/w:sym | //w:body//w:r/w:tab | //w:body//w:r/w:noBreakHyphen | //w:body//w:r/w:softHyphen | //w:body//w:r/w:br)", nsManager);

            foreach (XmlNode theData in theNodeList)
            {
                // we look the range structures
                switch (theData.Name)
                {
                    case "w:t":
                        // we have text
                        TextCount += theData.InnerText.Length;
                        break;
                    case "w:sym":
                        // we have a symbol
                        TextCount++;
                        break;

                    default:
                        // Anything else we simply increment the counter
                        OtherCount++;
                        break;

                }
            }
            // now count paragraphs and section and breaks
            theNodeList = theRoot.SelectNodes(@"(//w:body//w:p | //w:body//w:sectPr)", nsManager);
            if (theNodeList != null)
            {
                OtherCount += theNodeList.Count;
            }
            RangeCharacterCount = TextCount + OtherCount;

            toolStripStatusLabel1.Text = "Counted " + RangeCharacterCount + " characters in "
                + ((float)theStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";

            toolStripProgressBar1.Maximum = RangeCharacterCount;  // To show progress, but the count isn't accurate.
            Application.DoEvents();
            /*
                * Look for text or symbols in the document
            */
            try
            {
                // Get the styles in use
                GetStylesInUse(theRoot, nsManager, theStyleDictionary);

                // Load decomposed glyphs if we have specified a file
                if (!GlyphsLoaded && CombDecomposedChars.Checked && DecompGlyphBox.Text != "")
                {
                    GlyphsLoaded = LoadDecomposedGlyphs(theGlyphDictionary, excelApp);
                }

                if (AnalyseByFont.Checked)
                {
                    try
                    {
                        string theParagraphFont = "";
                        theNodeList = theRoot.SelectNodes(@"//w:body//w:p", nsManager);  // Find the paragraphs
                        foreach (XmlNode theParagraphData in theNodeList)
                        {
                            // Check if the paragraph has a font - we ignore this, it gave misleading results.

                            //FontName = XmlLookup(theParagraphData, "w:pPr/w:rPr/wx:font", nsManager, "wx:val", "");
                            //FontName = "";
                            //if (FontName == "")
                            //{
                            // Determine the paragraph's style
                            string theParagraphStyleID = XmlLookup(theParagraphData, "w:pPr/w:pStyle", nsManager, "w:val", "DefaultParagraphFont");
                            if (theStyleDictionary.Keys.Contains(theParagraphStyleID))
                            {
                                FontName = theStyleDictionary[theParagraphStyleID];
                            }
                            else
                            {
                                FontName = GetDefaultFont(theStyleDictionary, theParagraphData);
                            }
                            //}
                            theParagraphFont = FontName;  // Remember the paragraph font for the end of line.
                            XmlNodeList theRanges = theParagraphData.SelectNodes("w:r", nsManager);
                            TextString = "";
                            /*
                             * We go through the document a range at a time.  If we find a symbol whose font is the same as that of an existing range
                             * we concatenate the symbol to that range.
                             */
                            string OldFontName = "";
                            foreach (XmlNode theRangeData in theRanges)
                            {
                                XmlNode theSymbol = theRangeData.SelectSingleNode("w:sym", nsManager);
                                if (theSymbol != null)
                                {
                                    // we have a symbol
                                    FontName = theSymbol.Attributes["w:font"].Value;
                                    string theSymbolValue = theSymbol.Attributes["w:char"].Value;
                                    char theChar = Convert.ToChar(Convert.ToUInt16(theSymbolValue, 16));  // get the character number
                                    if (FontName == OldFontName)
                                    {
                                        // Concatenate the text string
                                        TextString += Convert.ToString(theChar); // make it a string concatenating it with previous symbols.
                                    }
                                    else
                                    {
                                        // Analyse the text string, then remember the old font and start a new text string
                                        CharacterCount = AnalyseString(theCharacterDictionary, theContextDictionary, OldFontName, TextString,
                                            RangeCharacterCount, CharacterCount, theStopwatch);
                                        OldFontName = FontName;
                                        TextString = Convert.ToString(theChar); // make it a string concatenating it with previous symbols. 
                                    }

                                }
                                else
                                {

                                    // See if there is a font defined in the range and use that
                                    FontName = XmlLookup(theRangeData, "w:rPr/wx:font", nsManager, "wx:val", "");
                                    if (FontName == "")
                                    {
                                        string theStyleID = XmlLookup(theRangeData, "w:rPr/w:rStyle", nsManager, "w:val", "");
                                        if (theStyleID != "" && theStyleDictionary.Keys.Contains(theStyleID))
                                        {
                                            // If we have no style nor do we have a font for the style, we do nothing
                                            // Otherwise we get the font name for the style.
                                            FontName = theStyleDictionary[theStyleID];
                                        }
                                        else
                                        {
                                            FontName = theParagraphFont; // we pick up the paragraph font
                                        }
                                    }

                                    // Look for text
                                    XmlNode theText = theRangeData.SelectSingleNode("w:t", nsManager);
                                    if (theText != null)
                                    {
                                        if (FontName == OldFontName)
                                        {
                                            TextString += theText.InnerText;
                                        }
                                        else
                                        {
                                            CharacterCount = AnalyseString(theCharacterDictionary, theContextDictionary,  OldFontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                                            OldFontName = FontName;
                                            TextString = theText.InnerText;
                                        }
                                    }
                                    for (int Counter = 0; Counter < SpecialCharacterKeys.Count(); Counter++)
                                    {
                                        XmlNode theSpecialChar = theRangeData.SelectSingleNode(SpecialCharacterKeys[Counter], nsManager);
                                        if (theSpecialChar != null)
                                        {
                                            // We've found a special character
                                            TextString += SpecialCharacterValues[Counter];
                                        }
                                    }
                                    //
                                    // Look for break characters
                                    //
                                    XmlNode theBreak = theRangeData.SelectSingleNode("w:br", nsManager);
                                    if (theBreak != null)
                                    {
                                        if (theBreak.Attributes.Count > 0)
                                        {
                                            try
                                            {
                                                TextString += theBreakDictionary[theBreak.Attributes["w:type"].Value];
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        else
                                        {
                                            // a break with nothing else is U+000B = \v
                                            TextString += "\v";
                                        }
                                    }
                                    // Now look for a section break
                                    XmlNode theSectionBreak = theParagraphData.SelectSingleNode("w:pPr/w:sectPr", nsManager);
                                    if (theSectionBreak != null)
                                    {
                                        TextString += "\f";
                                    }
                                }
                            }
                            if (TextString != "")
                            {
                                // We have some text to process
                                CharacterCount = AnalyseString(theCharacterDictionary, theContextDictionary, FontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                                TextString = "";
                                // Now add the end of line marker
                            }
                            CharacterCount = AnalyseString(theCharacterDictionary, theContextDictionary, theParagraphFont, "\r", RangeCharacterCount, CharacterCount, theStopwatch);


                        }
                        // Now look for any more sections
                        theNodeList = theRoot.SelectNodes(@"//w:body/wx:sect", nsManager);
                        if (theNodeList != null)
                        {
                            TextString = "";
                            for (int Counter = 0; Counter < theNodeList.Count; Counter++)
                            {
                                TextString += "\f";  // section/page break gives a form feed.
                            }
                            CharacterCount = AnalyseString(theCharacterDictionary, theContextDictionary, FontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                            TextString = "";
                        }
                    }
                    catch (Exception Ex)
                    {

                        MessageBox.Show(Ex.Message + "\r\r" + Ex.StackTrace, "Error in character counting - analysed by font", MessageBoxButtons.OK);
                        CloseApps(this);

                    }


                }
                else
                {
                    // We aren't analysing by font
                    try
                    {
                        theNodeList = theRoot.SelectNodes(@"//w:body//w:p", nsManager);  // Find the paragraphs
                        theFirstFont = theStyleDictionary["DefaultParagraphFont"]; // The default font
                        TextString = "";
                        foreach (XmlNode theParagraphData in theNodeList)
                        {
                            // Determine the paragraph's type

                            XmlNodeList theRanges = theParagraphData.SelectNodes("w:r", nsManager);
                            foreach (XmlNode theRangeData in theRanges)
                            {
                                XmlNode theSymbol = theRangeData.SelectSingleNode("w:sym", nsManager);
                                if (theSymbol != null)
                                {
                                    // we have a symbol
                                    string theSymbolValue = theSymbol.Attributes["w:char"].Value;
                                    char theChar = Convert.ToChar(Convert.ToUInt16(theSymbolValue, 16));  // get the character number
                                    TextString += Convert.ToString(theChar); // make it a string

                                }
                                else
                                {
                                    // Look for text
                                    XmlNode theText = theRangeData.SelectSingleNode("w:t", nsManager);
                                    if (theText != null)
                                    {
                                        TextString += theText.InnerText;
                                    }
                                    for (int Counter = 0; Counter < SpecialCharacterKeys.Count(); Counter++)
                                    {
                                        XmlNode theSpecialChar = theRangeData.SelectSingleNode(SpecialCharacterKeys[Counter], nsManager);
                                        if (theSpecialChar != null)
                                        {
                                            // We've found a special character
                                            TextString += SpecialCharacterValues[Counter];
                                        }
                                    }
                                    //
                                    // Look for break characters
                                    //
                                    XmlNode theBreak = theRangeData.SelectSingleNode("w:br", nsManager);
                                    if (theBreak != null)
                                    {
                                        if (theBreak.Attributes.Count > 0)
                                        {
                                            try
                                            {
                                                TextString += theBreakDictionary[theBreak.Attributes["w:type"].Value];
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        else
                                        {
                                            // a break with nothing else is U+000B = \v
                                            TextString += "\v";
                                        }
                                    }
                                }
                            }
                            // Now look for a section break
                            XmlNode theSectionBreak = theParagraphData.SelectSingleNode("w:pPr/w:sectPr", nsManager);
                            if (theSectionBreak != null)
                            {
                                TextString += "\f"; // the break character
                            }

                            // Now add the end of line marker
                            TextString += "\r";

                        }
                        /*
                        // Now look for sections and page breaks
                        theNodeList = theRoot.SelectNodes(@"//w:body/wx:sect", nsManager);
                        if (theNodeList != null)
                        {
                            for (int Counter = 0; Counter < theNodeList.Count - 1; Counter++)
                            {
                                TextString += "\f";  // section/page break gives a form feed.
                            }
                        }
                        */
                        CharacterCount = AnalyseString(theCharacterDictionary, theContextDictionary, "", TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                        TextString = "";  // clear the text string.

                    }
                    catch (Exception Ex)
                    {

                        MessageBox.Show(Ex.Message + "\r\r" + Ex.StackTrace, "Error in character counting - not analysed by font", MessageBoxButtons.OK);
                        CloseApps(this);

                    }

                }
            }

            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Error in analysing text", MessageBoxButtons.OK);
                CloseApps(this);
            }
            listNormalisedErrors.Sort(listNormalisedErrors.Columns[0], ListSortDirection.Ascending);  // Sort first
            return CharacterCount;
        }

        private string XmlLookup(XmlNode theNode, string theSearchPath, XmlNamespaceManager nsManager, string theValueID, string InputString = "")
        {
            /*
             * This looks up something in Xml and updates the input string with the returned value.  Otherwise
             * it returns the input string.  The idea is to update some data with new information
             */
            XmlNode theChildNode = theNode.SelectSingleNode(theSearchPath, nsManager);
            if (theChildNode == null)
            {
                // We didn't find anything, so return the input string
                return InputString;
            }
            else
            {
                try
                {
                    string tmpString = theChildNode.Attributes[theValueID].Value;
                    return tmpString;
                }
                catch (Exception Ex)
                {
                    // Something went wrong
                    string theError = Ex.Message;
                    return InputString;
                }
            }
        }
        private string GetDefaultFont(Dictionary<string, string> theStyleDictionary, XmlNode theNode)
        {
            string NodePath = GetNodePath(theNode, "");
            string theDefaultID = "DefaultParagraphFont";  // we assume a normal paragraph
            if (NodePath.Contains("w:tbl"))
            {
                // we have a table
                theDefaultID = "Default Table";
            }

            return theStyleDictionary[theDefaultID];
        }
        private void GetStylesInUse(XmlNode theRoot, XmlNamespaceManager nsManager, Dictionary<string, string> theStyleDictionary)
        {                // Load a list of current styles and their fonts
            XmlNodeList theNodeList = theRoot.SelectNodes(@"//w:styles/w:style", nsManager);
            theStyleDictionary.Clear();  // Empty the style dictionary
            // First look for the styles that have fonts
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                // For some we can't search on wx:font so we have to iterate
                if (theFont != null)
                {
                    string theFontName = theFont.Attributes["wx:val"].Value;
                    theStyleDictionary.Add(theStyleID, theFontName);
                }
            }
            // Now look for the default fonts - we do this as a second pass in case they don't appear first
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                // For some we can't search on wx:font so we have to iterate
                if (theFont != null)
                {
                    if (theStyle.Attributes.Count == 3)
                    {
                        // check to see if this is the default.
                        bool IsDefault = false;
                        try
                        {
                            IsDefault = (theStyle.Attributes[@"w:default"].Value == "on" /*&& theStyle.Attributes[@"w:type"].Value == "paragraph"*/);
                        }
                        catch
                        {
                        }
                        if (IsDefault)
                        {


                            string theDefaultStyle = theStyle.Attributes[@"w:styleId"].Value;
                            // We have found a default style so we look up its font and add to the nominal styles.
                            switch (theStyle.Attributes[@"w:type"].Value)
                            {
                                case "paragraph":
                                    theStyleDictionary["DefaultParagraphFont"] = theStyleDictionary[theDefaultStyle];
                                    break;
                                case "table":
                                    theStyleDictionary["Default Table"] = theStyleDictionary[theDefaultStyle];
                                    break;
                                case "character":
                                    theStyleDictionary["Default Character"] = theStyleDictionary[theDefaultStyle];
                                    break;

                            }
                        }
                    }
                }
            }
            // Now the styles that don't - we have to get the font of the style on which they are based.
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                if (theFont == null)
                {
                    // Get the font based on what the style is based on
                    XmlNode theBasedOnStyle = theStyle.SelectSingleNode("w:basedOn", nsManager);
                    if (theBasedOnStyle != null)
                    {
                        // Save the font name for the style on which this one is based.
                        string basedOnStyleID = theBasedOnStyle.Attributes["w:val"].Value;
                        string theFontName = theStyleDictionary[basedOnStyleID];
                        theStyleDictionary[theStyleID] = theFontName;
                    }
                    else
                    {
                        // Use the default paragraph font
                        theStyleDictionary[theStyleID] = theStyleDictionary["DefaultParagraphFont"];
                    }
                }
            }
        }

        private string GetNodePath(XmlNode theNode, string InputType)
        {
            // Iteratively walk up the nodes.
            XmlNode theParent = theNode.ParentNode;
            if (theParent != null)
            {
                string tmpString = theParent.Name + "/" + InputType;
                GetNodePath(theParent, tmpString);
                return tmpString;
            }
            else
            {
                return InputType;
            }
        }
        private int AnalyseString(Dictionary<CharacterDescriptor, int> theCharacterDictionary, Dictionary<ContextDescriptor, int> theContextDictionary,
            string FontName, string TextString, int RangeCharacterCount, int CharacterCount, Stopwatch theStopwatch)
        {
            /*
             * We shall first use the data for legacy decomposed characters to count them before we use the built-in functions that handle
             * decomposed Unicode characters.
             */

            CharacterDescriptor theKey = null;
            string tmpString = "";
            if (AnalyseByFont.Checked && CombDecomposedChars.Checked && theGlyphDictionary.Keys.Contains(FontName))
            {
                // We will count the glyphs loaded as single characters if we have the data
                try
                {
                    Regex theGlyphs = new Regex(theGlyphDictionary[FontName]);

                    MatchCollection theMatches = theGlyphs.Matches(TextString);
                    foreach (Match theMatch in theMatches)
                    {
                        string theString = theMatch.Value.ToString();
                        theKey = new CharacterDescriptor(FontName, theString);
                        IncrementCharCount(theCharacterDictionary, theKey);
                        CharacterCount += theString.Length;
                        ReportProgress(CharacterCount, RangeCharacterCount, theStopwatch);

                    }

                    // now remove all those characters
                    tmpString = theGlyphs.Replace(TextString, "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\r" + ex.StackTrace);
                    string tmpMessage = ex.Message;
                    CloseApps(null, null);
                }
            }
            else
            {
                tmpString = TextString;
            }
            /*
             * We shall now use the built-in Unicode functions to find Unicode decomposed characters
             */
            if (CombDecomposedChars.Checked)
            {
                TextElementEnumerator theTextElements = StringInfo.GetTextElementEnumerator(tmpString);
                while (theTextElements.MoveNext())
                {
                    string theString = theTextElements.GetTextElement();
                    if (AnalyseByFont.Checked)
                    {
                        theKey = new CharacterDescriptor(FontName, theString);
                    }
                    else
                    {
                        theKey = new CharacterDescriptor(theString);
                    }
                    CharacterCount += theString.Length;
                    IncrementCharCount(theCharacterDictionary, theKey);
                    ReportProgress(CharacterCount, RangeCharacterCount, theStopwatch);
                    if (theString.Length > 1)
                    {
                        CheckNormalisation(theString);  // Check to see if the string is normalised
                    }
                }
            }
            else
            {
                for (int i = 0; i < tmpString.Length; i++)
                {
                    if (AnalyseByFont.Checked)
                    {
                        theKey = new CharacterDescriptor(FontName, tmpString[i].ToString());
                    }
                    else
                    {
                        theKey = new CharacterDescriptor(tmpString[i].ToString());
                    }
                    IncrementCharCount(theCharacterDictionary, theKey);
                    ReportProgress(CharacterCount, RangeCharacterCount, theStopwatch);
                    CharacterCount++;
                }
            }
            // We shall now look for contexts
            if (AnalyseContext)
            {
                foreach(TargetDescriptor theTarget in TargetList)
                {
                    string pattern = theTarget.GetRegEx();
                    MatchCollection contexts = Regex.Matches(TextString, pattern);
                    ContextDescriptor theContextKey;
                    foreach (Match context in contexts)
                    {
                        ContextDescriptor theContext = new ContextDescriptor(FontName, theTarget.Target, context.ToString());
                        if (AnalyseByFont.Checked)
                        {
                            theContextKey = new ContextDescriptor(FontName, theTarget.Target, context.ToString());
                        }
                        else
                        {
                            theContextKey = new ContextDescriptor(theTarget.Target, context.ToString());
                        }
                        IncrementContextCount(theContextDictionary, theContextKey);
                    }

                }
            }
            return CharacterCount;
        }
        private void CheckNormalisation(string theString)
        {
            if (theString.IsNormalized())
            {
                return;  // We need do no more
            }
            string theNormalisedString = theString.Normalize(NormalizationForm.FormC);  // Full canonical normalisation
            theString = GetCodes(theString, USV);
            // Look to see if we have found it already
            bool Found = false;
            foreach (DataGridViewRow theViewRow in listNormalisedErrors.Rows)
            {
                if (theViewRow.Cells[0].Value.ToString() == theString)
                {
                    Found = true;
                    break;
                }
            }
            if (!Found)
            {
                theNormalisedString = GetCodes(theNormalisedString, USV);
                string[] theRow = new string[] { theString, theNormalisedString };
                listNormalisedErrors.Rows.Add(theRow);
                tabControl1.SelectedTab = tabControl1.TabPages[2];
                Application.DoEvents();
            }

        }

        private string GetCodes(string theString, int theCode)
        {
            /*
             * Return the character code for a character
             */
            string tmpString = "";
            foreach (var theChar in theString)
            {
                /*
                 * We loop through each characer in the string in case we get a composed character.
                 * This aspect hasn't been tested, so I don't know if it will work, but worth a try.
                 */
                int temp = Convert.ToUInt16(theChar);
                string tempstring = "";

                switch (theCode)
                {
                    case dec:
                        tempstring = temp.ToString();
                        break;
                    case hex:
                        tempstring = temp.ToString("X");  // Upper case hex
                        break;
                    case USV:
                        tempstring = String.Format("U+{0:X4}", temp);  // USV
                        break;
                }
                tmpString += tempstring + " ";
            }

            return tmpString.Trim(); ;
        }

        private void documentationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(Application.StartupPath, "CharacterCounter.docx");
            try
            {
                System.Diagnostics.Process.Start(HelpPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file " + HelpPath + "\r" + ex.Message, "Error", MessageBoxButtons.OK);
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 About = new AboutBox1();
            About.Show();
        }

        private void LicenseMenuItem_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(Application.StartupPath, "gpl.txt");
            System.Diagnostics.Process.Start("Wordpad.exe", '"' + HelpPath + '"');
        }


        private bool MarkWorking(bool On, bool Working, Dictionary<string, bool> ControlDictionary)
        {
            if (Working && On)
            {
                return true;  // Do no more if we are already working and want to mark as working
            }
            if (!Working && !On)
            {
                return false;  // Do no more if we are not working and want to mark as not working
            }
            // Mark the application state
            //Application.UseWaitCursor = On;
            this.UseWaitCursor = On;
            if (On)
            {
                ControlDictionary.Clear();  // We are about to mark the controls as working and will record their current state
            }
            SwitchControl(this, !On, ControlDictionary);  // now enable or disable text boxes, buttons and check boxes
            if (On)
            {
                btnClose.Text = "Abort";
                btnClose.Enabled = true;  // allow us to abort if we wish
            }
            else
            {
                btnClose.Text = "Close";
            }
            return On;
        }
        private void SwitchControl(Control Parent, bool Enable, Dictionary<string, bool> ControlDictionary)
        {
            foreach (Control theControl in Parent.Controls)
            {
                if ((theControl is TextBox || theControl is Button || theControl is CheckBox) && theControl.Name != "btnPause")
                {
                    if (!Enable)
                    {
                        // We are about to disable it so we record its current state
                        theControlDictionary.Add(theControl.Name, theControl.Enabled);
                        theControl.Enabled = Enable;
                    }
                    else
                    {
                        // Get its previous state
                        theControl.Enabled = theControlDictionary[theControl.Name];
                    }

                }
                SwitchControl(theControl, Enable, ControlDictionary); // Handle daughter controls
            }

        }

        private void IncrementCharCount(Dictionary<CharacterDescriptor, int> theDictionary, CharacterDescriptor theKey)
        {
            /*
             * Increment the value of the relevant key, and handle a non-existent value.
             */
            if (theDictionary.Keys.Contains(theKey))
            {
                /*
                 * Increment the character count
                 */
                //DateTime Start = DateTime.Now;
                theDictionary[theKey]++;
                //toolStripStatusLabel2.Text = DateTime.Now.Subtract(Start).TotalSeconds.ToString();
            }
            else
            {
                /*
                 * If haven't met the character/font combination before, the increment operation fails
                 * so we come here and generate a new entry with a count of 1
                 */
                //DateTime Start = DateTime.Now;
                theDictionary.Add(theKey, 1);
                // toolStripStatusLabel2.Text = " Key added in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString();
            }
            return;
        }
        private void IncrementContextCount(Dictionary<ContextDescriptor, int> theDictionary, ContextDescriptor theKey)
        {
            /*
             * Increment the value of the relevant key, and handle a non-existent value.
             */
            if (theDictionary.Keys.Contains(theKey))
            {
                /*
                 * Increment the character count
                 */
                //DateTime Start = DateTime.Now;
                theDictionary[theKey]++;
                //toolStripStatusLabel2.Text = DateTime.Now.Subtract(Start).TotalSeconds.ToString();
            }
            else
            {
                /*
                 * If haven't met the character/font combination before, the increment operation fails
                 * so we come here and generate a new entry with a count of 1
                 */
                //DateTime Start = DateTime.Now;
                theDictionary.Add(theKey, 1);
                // toolStripStatusLabel2.Text = " Key added in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString();
            }
            return;
        }

        private string LoadAggregateStats(Dictionary<CharacterDescriptor, int> theMultiDictionary, Dictionary<CharacterDescriptor, int> theMultiSummaryDictionary,
            Dictionary<CharacterDescriptor, int> theIndividualDictionary,
            Dictionary<ContextDescriptor, int> theMultiContextDictionary, Dictionary<ContextDescriptor, int> theMultiContextSummaryDictionary,
            Dictionary<ContextDescriptor, int> theIndividualContextDictionary,
            string FileList, string InputFile)
        {
            // Here is where we load the aggregate statistics dictionary.
            string tmpString = FileList;
            int Counter = Convert.ToInt16(FileCounter.Text);
            foreach (KeyValuePair<CharacterDescriptor, int> kvp in theIndividualDictionary)
            {
                CharacterDescriptor tmpKey = new CharacterDescriptor(kvp.Key);
                CharacterDescriptor tmpKey2 = new CharacterDescriptor(tmpKey);
                /*
                 * We maintain two aggregate dictionaries, one with character counts by filename and font and character, and the other just by font and character.
                 */
                tmpKey.FileName = null;  // Make sure the file name is null for the summary.
                AddToCharDictionary(theMultiSummaryDictionary, tmpKey, kvp.Value); // just font and character

                tmpKey2.FileName = InputFile;
                AddToCharDictionary(theMultiDictionary, tmpKey2, kvp.Value);  // and the filename, too.


            }
            foreach (KeyValuePair<ContextDescriptor, int> kvp in theIndividualContextDictionary)
            {
                ContextDescriptor tmpKey = new ContextDescriptor(kvp.Key);
                ContextDescriptor tmpKey2 = new ContextDescriptor(tmpKey);
                /*
                 * We maintain two aggregate dictionaries, one with character counts by filename and font and character, and the other just by font and character.
                 */
                tmpKey.FileName = null;  // Make sure the file name is null for the summary.
                AddToContextDictionary(theMultiContextSummaryDictionary, tmpKey, kvp.Value); // just font and character

                tmpKey2.FileName = InputFile;
                AddToContextDictionary(theMultiContextDictionary, tmpKey2, kvp.Value);  // and the filename, too.


            }

            if (tmpString != "")
            {
                tmpString += ", ";
            }
            tmpString += InputFile;
            Counter++;
            FileCounter.Text = Counter.ToString();
            return tmpString;
        }
        private void AddToCharDictionary(Dictionary<CharacterDescriptor, int> theDictionary, CharacterDescriptor theKey, int theCount)
        {
            // Add to a CharacterDescriptor dictionary if the key isn't there, otherwise increment a count.
            if (theDictionary.Keys.Contains(theKey))
            {
                theDictionary[theKey] += theCount;
            }
            else
            {
                theDictionary.Add(theKey, theCount);
            }


        }
        private void AddToContextDictionary(Dictionary<ContextDescriptor, int> theDictionary, ContextDescriptor theKey, int theCount)
        {
            // Add to a CharacterDescriptor dictionary if the key isn't there, otherwise increment a count.
            if (theDictionary.Keys.Contains(theKey))
            {
                theDictionary[theKey] += theCount;
            }
            else
            {
                theDictionary.Add(theKey, theCount);
            }


        }

        private void ReportProgress(int CharacterCount, int RangeCharacterCount, Stopwatch theStopwatch)
        {
            if ((CharacterCount % 100) == 0)
            {
                // report progress
                TimeSpan TimeToFinish = TimeSpan.FromTicks((long)((RangeCharacterCount - CharacterCount) * ((float)theStopwatch.ElapsedTicks / CharacterCount)));
                toolStripStatusLabel1.Text = CharacterCount.ToString() + " of about " + RangeCharacterCount.ToString()
                    + " chars. Approx time to finish analysis: " + TimeToFinish.ToString(@"hh\:mm\:ss"); ;
                toolStripProgressBar1.Value = Math.Min(CharacterCount, toolStripProgressBar1.Maximum);
                Application.DoEvents();
                if (!theStopwatch.IsRunning)
                {
                    theStopwatch.Start();
                }

            }
            return;

        }

        private void btnListFonts_Click(object sender, EventArgs e)
        {
            // list the fonts in the documnent
            Control theControl = (Control)sender;
            Working = MarkWorking(true, Working, theControlDictionary);
            Application.DoEvents();
            Stopwatch theStopwatch = new Stopwatch();
            theStopwatch.Start();
            AnalyseByFont.Enabled = false;
            CombDecomposedChars.Enabled = false;
            toolStripProgressBar1.Value = 0;
            XmlNode theRoot = null;
            List<string> theFontTable = new List<string>();
            toolStripStatusLabel1.Text = "Listing fonts in " + Path.GetFileName(InputFileBox.Text) + "...";
            Application.DoEvents();
            /*
             * Open the Word file(s)
             *
             */
            theFontTable.Clear();
            foreach (string theFile in openInputDialogue.FileNames)
            {
                if (GetFileType(theFile) == WordDoc)  // only get the data if the files are Word files.
                {

                    if (DocumentLoader(theFile, theXMLDictionary, ref nsManager))
                    {
                        // only run if successful
                        string theFileName = Path.GetFileName(theFile);
                        theRoot = theXMLDictionary[theFileName].DocumentElement;
                        XmlNodeList theFontList = theRoot.SelectNodes(@"w:fonts/w:font", nsManager);
                        foreach (XmlNode theFont in theFontList)
                        {
                            string theFontName = theFont.Attributes["w:name"].Value;
                            if (!theFontTable.Contains(theFontName))
                            {
                                theFontTable.Add(theFont.Attributes["w:name"].Value);  // Add the font if we don't have it already
                            }
                        }
                        theFontTable.Sort();
                        FontList.Items.Clear();  // Clear so we don't load it more than once.
                        foreach (string theFont in theFontTable)
                        {
                            FontList.Items.Add(theFont);
                        }
                        toolStripStatusLabel1.Text = "Finished loading fonts from " + theFile;
                        Application.DoEvents();
                    }
                }
            }
            Working = MarkWorking(false, Working, theControlDictionary);
            string theFontListFile = "";
            AnalyseByFont.Enabled = true;
            CombDecomposedChars.Enabled = true;
            toolStripStatusLabel1.Text = "Finished loading fonts in " +
                ((float)theStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            if (Individual)
            {
                theFontListFile = FontListFileBox.Text;
            }
            else
            {
                theFontListFile = BulkFontListFileBox.Text;
            }
            if (theControl.Name == "btnSaveFontList")
            {
                WriteFontList(theFontListFile, "List of fonts", FontList);
            }
            Working = MarkWorking(false, Working, theControlDictionary);
            btnSaveFontList.Enabled = theFontListFile != "";
            theStopwatch.Stop();
            theStopwatch = null;

        }

        private void InputFileBox_TextChanged(object sender, EventArgs e)
        {
            // A different file so we need to reload it and clear lots of things.
            TextBox theBox = (TextBox)sender;
            if (theBox.Text != "" && !File.Exists(theBox.Text))
            {
                MessageBox.Show("File does not exist", "Error", MessageBoxButtons.OK);
                theBox.Select();
                return;
            }
            if (InputFileName != null)
            {
                theXMLDictionary.Remove(InputFileName);
                theTextDictionary.Remove(InputFileName);
            }
            InputFileName = Path.GetFileName(InputFileBox.Text);  // Remember the file name for later.
            // Suggest names for the output files
            OutputFileBox.Text = Path.Combine(OutputDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + ".xlsx";
            ErrorListBox.Text = Path.Combine(ErrorDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + " (Suggested Chars).xlsx";
            InputDir = Path.GetDirectoryName(openInputDialogue.FileName);
            Registry.SetValue(keyName, "InputDir", InputDir);
            btnAnalyse.Enabled = true;
            saveExcelDialogue.FileName = OutputFileBox.Text;
            FileType = GetFileType(InputFileBox.Text, false);
            if (FileType == WordDoc)
            {
                // Only enable these if we are analysing Word documents.
                StyleListFileBox.Text = Path.Combine(StyleDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + " (Styles).xlsx";
                FontListFileBox.Text = Path.Combine(FontDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + " (Fonts).xlsx";
                XMLFileBox.Text = Path.Combine(XMLDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + ".xml";
            }

            listStyles.Rows.Clear(); // clear the styles
            FontList.Items.Clear(); // and the fonts
            theStyleDictionary.Clear();
            listNormalisedErrors.Rows.Clear();
            theStyleDictionary.Clear();
            FileType = GetFileType(theBox.Text);  // redetermine the file type.
            openInputDialogue.FileName = theBox.Text;
            btnAnalyse.Enabled = true;  // Let you analyse the new file.


        }

        private void btnGetStyles_Click(object sender, EventArgs e)
        {
            Control theControl = (Control)sender;
            theControl.Enabled = false;
            XmlNode theRoot = null;
            listStyles.Rows.Clear();  // Empty the style list
            Dictionary<Style, bool> theStyleList = new Dictionary<Style, bool>(50, new StyleComparer());
            //List<DataGridViewRow> theStyleList = new List<DataGridViewRow>(10);
            foreach (string theFile in openInputDialogue.FileNames)
            {
                if (GetFileType(theFile) == WordDoc)  //Only analyse Word documents
                {
                    if (DocumentLoader(theFile, theXMLDictionary, ref nsManager))
                    {
                        string theFileName = Path.GetFileName(theFile); // We just use the file name, not the full path
                        theRoot = theXMLDictionary[theFileName].DocumentElement;
                        GetStylesInUse(theRoot, nsManager, theStyleDictionary);
                        XmlNode theStylesNode = theRoot.SelectSingleNode("w:styles", nsManager);
                        foreach (string theStyleID in theStyleDictionary.Keys)
                        {
                            // Look up the font name rather than the ID
                            string tmpStyle = theStyleID;
                            XmlNode theStyleNode = theStylesNode.SelectSingleNode("w:style[@w:styleId = \"" + theStyleID + "\"]", nsManager);
                            if (theStyleNode != null)
                            {
                                XmlNode theNameNode = theStyleNode.SelectSingleNode("w:name", nsManager);
                                if (theStyleNode.Attributes != null)
                                {
                                    tmpStyle = theNameNode.Attributes["w:val"].Value;
                                }
                            }
                            // Create a new row, and add it if we haven't already got it.
                            DataGridViewRow theRow = new DataGridViewRow();
                            Style theStyle = new Style(tmpStyle, theStyleDictionary[theStyleID]);
                            if (!theStyleList.Keys.Contains(theStyle))
                            {
                                theStyleList.Add(theStyle, true);
                            }
                        }
                    }
                }

            }
            foreach (Style theStyle in theStyleList.Keys)
            {
                DataGridViewRow theRow = new DataGridViewRow();
                string[] theStringArray = { theStyle.Name, theStyle.Font };

                theRow.CreateCells(listStyles, theStringArray);

                //if (!theStyleList.Contains(theRow))
                //{
                //    theStyleList.Add(theRow);
                //    listStyles.Rows.Add(theRow);
                //}
                listStyles.Rows.Add(theRow);
                if (listStyles.Rows.Count > 0)
                {
                    listStyles.Sort(listStyles.Columns[0], ListSortDirection.Ascending);  // Sort the list
                }
            }
            string theStyleListFile = "";
            if (Individual)
            {
                theStyleListFile = StyleListFileBox.Text;
            }
            else
            {
                theStyleListFile = BulkStyleListBox.Text;
            }
            if (theControl.Name == "btnSaveStyles")
            {
                // Write the list to Excel
                WriteDataGridView(theStyleListFile, listStyles);
                Registry.SetValue(keyName, "StyleDir", StyleDir);

            }
            theControl.Enabled = true;
            btnSaveStyles.Enabled = theStyleListFile != "";
            Application.DoEvents();

        }
        private XmlDocument LoadWordDocument(string WordFile)
        {
            // Load the Word document into XML.
            Stopwatch theStopWatch = new Stopwatch();
            theStopWatch.Start();
            XmlDocument theXMLDocument = new XmlDocument();
            try
            {
                DialogResult theResult = DialogResult.Retry;
                while (theResult == DialogResult.Retry)
                {
                    try
                    {
                        wrdApp.Documents.Open(WordFile, missing, true);
                        theDocument = wrdApp.ActiveDocument;
                        theResult = DialogResult.OK;
                    }
                    catch (Exception ex)
                    {
                        COMException ComEx = (COMException)ex;
                        if (ComEx.ErrorCode == -2147023174) // RPC Server Unavailable
                        {
                            wrdApp = new Word();
                            theResult = DialogResult.Retry;
                        }
                        else
                        {
                            theResult = MessageBox.Show(ComEx.Message + "\r" + ComEx.StackTrace, "Word failed to open!", MessageBoxButtons.AbortRetryIgnore);
                            if (theResult == DialogResult.Abort)
                            {
                                CloseApps(); // Shut down
                                return null;
                            }
                            if (theResult == DialogResult.Ignore)
                            {
                                return null;
                            }
                        }
                    }
                }
                theDocument.Select();
                theXMLDocument.LoadXml(wrdApp.Selection.get_XML(false));
                theDocument.Close();  // We no longer need it.
                theDocument = null;
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Error opening document", MessageBoxButtons.OK);
                toolStripStatusLabel1.Text = "Failed to open" + WordFile + " after " +
                    ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
                theStopWatch.Stop();
                theStopWatch = null;
                return null;
            }
            toolStripStatusLabel1.Text = "Loaded document after " + ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            theStopWatch.Stop();
            theStopWatch = null;
            return theXMLDocument;
        }
        private bool DocumentLoader(string WordFile, Dictionary<string, XmlDocument> theXMLDictionary, ref XmlNamespaceManager nsManager)
        {
            /*
             *  If we've already loaded the document, we don't need to do it again
             *
             */
            string theFileName = Path.GetFileName(WordFile);  // We just use the file name as a key, not the whole path.
            if (theXMLDictionary.Keys.Contains(theFileName))
            {
                return true;
            }
            else
            {
                /*
                 * Open the Word file
                */
                XmlDocument theXMLDocument = null;
                DialogResult TheResult = DialogResult.Retry;
                while (TheResult == DialogResult.Retry)
                {
                    try
                    {
                        theXMLDocument = LoadWordDocument(WordFile);
                        TheResult = DialogResult.OK;
                    }
                    catch (Exception ex)
                    {
                        TheResult = MessageBox.Show("Failed to open " + WordFile + "\r" + ex.Message + "\r" + ex.StackTrace, "File open failure", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    }
                }
                if (theXMLDocument != null)
                {
                    nsManager = new XmlNamespaceManager(theXMLDocument.NameTable);
                    nsManager.AddNamespace("wx", wordmlxNamespace);
                    nsManager.AddNamespace("w", wordmlNamespace);
                    // If successful add the root to the dictionary
                    theXMLDictionary.Add(theFileName, theXMLDocument);
                    return true;
                }
                else
                {
                    // We failed
                    return false;
                }
            }
        }
        private bool DocumentLoader(string TextFile, Dictionary<string, string> theTextDictionary)
        {
            /*
             *  If we've already loaded the document, we don't need to do it again
             *
             */
            string theFileName = Path.GetFileName(TextFile);  // We just use the file name as a key, not the whole path.
            if (theTextDictionary.Keys.Contains(theFileName))
            {
                return true;
            }
            else
            {
                /*
                 * Open the text file
                */
                string theFile = null;
                DialogResult TheResult = DialogResult.Retry; // Assume success
                while (TheResult == DialogResult.Retry)
                {
                    try
                    {
                        theFile = LoadTextDocument(TextFile);
                        TheResult = DialogResult.OK;
                    }
                    catch (Exception ex)
                    {
                        TheResult = MessageBox.Show("Failed to open " + TextFile + "\r" + ex.Message + "\r" + ex.StackTrace, "File open failure", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    }
                }

                if (theFile != null)
                {
                    // If successful add the root to the dictionary
                    theTextDictionary.Add(theFileName, theFile);
                    return true;
                }
                else
                {
                    // We failed
                    return false;
                }
            }
        }

        private string LoadTextDocument(string TextFile)
        {
            // Load a text document into a string.
            Stopwatch theStopWatch = new Stopwatch();
            string theTextDocument = null;
            theStopWatch.Start();
            DialogResult Retry = DialogResult.Retry;
            while (Retry == DialogResult.Retry)
            {
                try
                {
                    theTextDocument = File.ReadAllText(TextFile, theEncoding);
                    Retry = DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retry = MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Word failed to open!", MessageBoxButtons.RetryCancel);
                    if (Retry == DialogResult.Cancel)
                    {
                        toolStripStatusLabel1.Text = "Document load cancelled after " + ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
                        theStopWatch.Stop();
                        theStopWatch = null;
                        return null;
                    }
                }
            }
            toolStripStatusLabel1.Text = "Loaded document after " + ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            theStopWatch.Stop();
            theStopWatch = null;
            return theTextDocument;
        }

        private void btnSaveXML_Click(object sender, EventArgs e)
        {
            DialogResult Retrying = System.Windows.Forms.DialogResult.Retry;
            bool DocumentLoaded = DocumentLoader(InputFileBox.Text, theXMLDictionary, ref nsManager);


            while (Retrying == System.Windows.Forms.DialogResult.Retry)
            {
                try
                {
                    theXMLDictionary[InputFileName].Save(XMLFileBox.Text);
                    Retrying = System.Windows.Forms.DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retrying = MessageBox.Show(Ex.Message, "Failed to save XML file", MessageBoxButtons.RetryCancel);
                    if (Retrying == System.Windows.Forms.DialogResult.Cancel)
                    {
                        toolStripStatusLabel1.Text = "Failed to save XML file " + Path.GetFileName(XMLFileBox.Text);
                        return;  // We do no more
                    }
                }
            }
            Registry.SetValue(keyName, "XMLDir", XMLDir); // Save the output directory
            toolStripStatusLabel1.Text = Path.GetFileName(XMLFileBox.Text) + " saved";
            System.Media.SystemSounds.Beep.Play();
            return;


        }

        private bool LoadDecomposedGlyphs(Dictionary<string, string> theGlyphDictionary, ExcelApp theApp)
        {
            DialogResult Retrying = DialogResult.Retry;
            int theRow = 2;
            ExcelRoot.Workbook theWorkbook = null;
            theGlyphDictionary.Clear();  // Make sure it is empty
            while (Retrying == DialogResult.Retry)
            {
                try
                {
                    theWorkbook = theApp.Workbooks.Open(DecompGlyphBox.Text, missing, true);
                    Retrying = DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retrying = MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Error opening glyph file", MessageBoxButtons.RetryCancel);
                    if (Retrying == DialogResult.Cancel)
                    {
                        return false;
                    }
                }
            }
            //
            //  We have opened the Excel file, so read it
            //
            ExcelRoot.Range theRange = theWorkbook.ActiveSheet.Cells[theRow, 1];
            while (theRange.Value != null)
            {
                string FontName = theRange.Font.Name;
                string theChar = theRange.Value;  // Escape out certain significant Regex characters.

                if (theGlyphDictionary.Keys.Contains(FontName))
                {
                    theGlyphDictionary[FontName] += theChar;
                }
                else
                {
                    theGlyphDictionary[FontName] = "(.[" + theChar;
                }
                theRow++;
                theRange = theWorkbook.ActiveSheet.Cells[theRow, 1];
            }
            theWorkbook.Close(false); // Close the workbook, we don't need it again.
            for (int KeyCounter = 0; KeyCounter < theGlyphDictionary.Keys.Count; KeyCounter++)
            {
                // we now close off the regular expression
                string theKeyName = theGlyphDictionary.Keys.ElementAt(KeyCounter);
                theGlyphDictionary[theKeyName] += "])";
            }
            // We end up with a regular expression of the form (.[<combchar>|<combchar>!..])
            // I.e. we match any character followed by a combining character.
            Registry.SetValue(keyName, "AggregateDir", AggregateDir);
            return true;
        }
        private bool LoadTargets(List<TargetDescriptor> TargetList, ExcelApp theApp)
        {
            DialogResult Retrying = DialogResult.Retry;
            int theRow = 2;
            ExcelRoot.Workbook theWorkbook = null;
            TargetList.Clear(); // Make sure it is empty
            bool result = true;
            toolStripStatusLabel1.Text = "Loading context targets from " + ContextCharacterFileBox.Text + "...";
            while (Retrying == DialogResult.Retry)
            {
                try
                {
                    theWorkbook = theApp.Workbooks.Open(ContextCharacterFileBox.Text, missing, true);
                    Retrying = DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retrying = MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Error opening glyph file", MessageBoxButtons.RetryCancel);
                    if (Retrying == DialogResult.Cancel)
                    {
                        return false;
                    }
                }
            }
            //
            //  We have opened the Excel file, so read it
            //
            string theTarget = theWorkbook.ActiveSheet.Cells[theRow, 1].Value;
            while (theTarget != null && result)
            {
                try
                {
                   TargetDescriptor newTarget = new TargetDescriptor(theTarget, theWorkbook.ActiveSheet.Cells[theRow, 2].Value,
                        theWorkbook.ActiveSheet.Cells[theRow, 3].Value);
                    if (newTarget.Valid)
                    {
                        TargetList.Add(newTarget);
                    }
                    else
                    {
                        result = false;  // We have failed
                        break;
                    }
                    string Regex = newTarget.GetRegEx();
                    theRow++;
                    theTarget = theWorkbook.ActiveSheet.Cells[theRow, 1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    result = false;
                }
            }
            theWorkbook.Close(false); // Close the workbook, we don't need it again.
            toolStripStatusLabel1.Text = "Finished loading context targets from " + ContextCharacterFileBox.Text;
            return result;
        }

        private void DecompGlyphBox_TextChanged(object sender, EventArgs e)
        {
            // The file name of the glyph file has changed so we need to reload
            // If the file doesn't exist, then we won't even try to load it
            TextBox theTextBox = (TextBox)sender;
            theGlyphDictionary.Clear();
            GlyphsLoaded = !File.Exists(theTextBox.Text);
            return;
        }

        private void btnSaveErrorList_Click(object sender, EventArgs e)
        {
            string theErrorFile;
            if (Individual)
            {
                theErrorFile = ErrorListBox.Text;
            }
            else
            {
                theErrorFile = BulkErrorListBox.Text;
            }
            WriteDataGridView(theErrorFile, listNormalisedErrors);
            Registry.SetValue(keyName, "ErrorDir", ErrorDir);

        }

        private void btnGetFont_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                FontBox.Text = fontDialog1.Font.Name;
                Registry.SetValue(keyName, "Font", FontBox.Text); //remember it
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Load the encodings
            foreach (EncodingInfo theEncodingInfo in Encoding.GetEncodings())
            {
                Encoding theEncoding = theEncodingInfo.GetEncoding();
                theEncodingDictionary[theEncodingInfo.DisplayName] = theEncoding; // save it in the dictionary
            }
            this.SetEncoding(Registry.GetValue(keyName, "Encoding", "Western European (Windows)").ToString());  // the default is ANSI Code Page 1252


        }

        private void btnGetEncoding_Click(object sender, EventArgs e)
        {
            EncodingForm theEncodingForm = new EncodingForm();
            DialogResult theResult = theEncodingForm.ShowDialog(this);
        }
        public void SetEncoding(string theEncodingName)
        {
            theEncoding = theEncodingDictionary[theEncodingName];
            EncodingTextBox.Text = theEncodingName;
            Registry.SetValue(keyName, "Encoding", theEncodingName); // Remember the encoding
        }
        public string GetEncoding()
        {
            if (theEncoding == null)
            {
                return "";
            }
            else
            {
                return theEncoding.EncodingName;
            }
        }

        private void AggregateStatsBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            if (theBox.Text == "")
            {
                AggregateStats.Enabled = false;
                AggregateStats.Checked = false;
                btnSaveAggregateStats.Enabled = false;
            }
            else
            {
                AggregateStats.Enabled = true;
                AggregateStats.Checked = true;

            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Working)
            {
                // We are aborting
                AggregateStats.Checked = false;  // So we don't try to save the file
                CloseApps(sender, e);
                return;  // and do no more
            }
            if (AggregateStats.Checked && !AggregateSaved)
            {
                AggregateSaved = WriteOutput(theAggregateDictionary, theAggregateContextDictionary, "", AggregateDir, AggregateStatsBox.Text, AggregateFileList, 
                    theAggregateSummaryDictionary, theAggregateContextSummaryDictionary);
            }

        }

        private void btnSaveAggregateStats_Click(object sender, EventArgs e)
        {
            AggregateSaved = WriteOutput(theAggregateDictionary, theAggregateContextDictionary, "", AggregateDir, AggregateStatsBox.Text, AggregateFileList, 
                theAggregateSummaryDictionary, theAggregateContextSummaryDictionary);
            btnSaveAggregateStats.Enabled = !AggregateSaved && AggregateStatsBox.Text != "" && FileCounter.Text != "0";
            System.Media.SystemSounds.Beep.Play();  // and beep
        }

        private void CombDecomposedChars_CheckStateChanged(object sender, EventArgs e)
        {
            btnDecompGlyph.Enabled = CombDecomposedChars.Checked;
        }

        private void CombiningCharacters_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(Application.StartupPath, "CombiningCharacters.xlsm");
            try
            {
                System.Diagnostics.Process.Start(HelpPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening file " + HelpPath + "\r" + ex.Message, "Error", MessageBoxButtons.OK);
            }


        }

        private void Individual_Entered(object sender, EventArgs e)
        {
            // Clear things if we have switched from bulk to individual
            if (Individual == false)
            {
                Individual = true;
                ClearLists();
            }
        }

        private void Bulk_Entered(object sender, EventArgs e)
        {
            if (Individual == true)
            {
                Individual = false;
                InputFolderBox.Text = InputDir;
                OutputFolderBox.Text = OutputDir;
                btnSaveFontList.Enabled = BulkFontListFileBox.Text != "";
                btnSaveStyles.Enabled = BulkStyleListBox.Text != "";
                btnSaveErrorList.Enabled = BulkErrorListBox.Text != "";
                btnGetFont.Enabled = true;
                btnGetEncoding.Enabled = true;
                ClearLists();
            }
        }
        private void ClearLists()
        {
            //  Clear lists when switching back and forth between individual and bulk processing
            openInputDialogue.FileName = ""; // Forget any files we selected.
            AggregateStatsBox.Text = "";  // and any aggregate files we specify
            FileCounter.Text = 0.ToString();
            FontList.Items.Clear();
            listStyles.Rows.Clear();
            listNormalisedErrors.Rows.Clear();
            theAggregateDictionary.Clear();
            theAggregateSummaryDictionary.Clear();
            btnSaveAggregateStats.Enabled = false;
            btnAnalyse.Enabled = false;

        }
        private void btnInputFolder_Click(object sender, EventArgs e)
        {
            FolderDialogue.RootFolder = Environment.SpecialFolder.Desktop;
            FolderDialogue.ShowNewFolderButton = false;
            FolderDialogue.SelectedPath = GetDirectory("InputDir");
            if (FolderDialogue.ShowDialog() == DialogResult.OK)
            {
                InputFolderBox.Text = FolderDialogue.SelectedPath;
            }

        }

        private void InputFolderBox_TextChanged(object sender, EventArgs e)
        {
            // A different folder so we need to clear everything it and clear lots of things.

            theXMLDictionary.Clear();
            theTextDictionary.Clear();
            listStyles.Rows.Clear(); // clear the styles
            FontList.Items.Clear(); // and the fonts
            theStyleDictionary.Clear();
            listNormalisedErrors.Rows.Clear();
            theStyleDictionary.Clear();
            if (InputFolderBox.Text != "")
            {
                InputDir = InputFolderBox.Text;
                Registry.SetValue(keyName, "InputDir", InputDir);
                btnSelectFiles.Enabled = true;
            }

        }

        private void btnSelectFiles_Click(object sender, EventArgs e)
        {
            openInputDialogue.Multiselect = true;
            openInputDialogue.InitialDirectory = GetDirectory("InputDir");
            if (openInputDialogue.ShowDialog() == DialogResult.OK && openInputDialogue.FileNames.Count() > 0)
            {
                theAggregateDictionary.Clear();  // Make sure it is empty before we analyse in bulk
                theAggregateSummaryDictionary.Clear(); // and the summary
                AggregateFileList = "";  // and clear the list of files, too.
                FileCounter.Text = "0";

                if (OutputFolderBox.Text != "")
                {
                    btnAnalyse.Enabled = true;
                    FontList.Items.Clear();
                    listStyles.Rows.Clear();
                    WriteIndividualFile.Checked = true;
                    WriteIndividualFile.Enabled = true;
                    btnListFonts.Enabled = openInputDialogue.FileNames.Count() > 0;
                    btnGetStyles.Enabled = openInputDialogue.FileNames.Count() > 0;

                    btnSaveFontList.Enabled = BulkFontListFileBox.Text != "" && btnListFonts.Enabled;
                    btnSaveStyles.Enabled = BulkStyleListBox.Text != "" && btnGetStyles.Enabled;
                    btnSaveErrorList.Enabled = BulkErrorListBox.Text != "" && openInputDialogue.FileNames.Count() > 0;
                }
                else
                {
                    WriteIndividualFile.Enabled = false;
                    WriteIndividualFile.Checked = false;
                }
            }


        }

        private void btnCharStatFolder_Click(object sender, EventArgs e)
        {
            FolderDialogue.Description = "Select the directory to receive the individual files";
            FolderDialogue.SelectedPath = GetDirectory("OutputDir");
            FolderDialogue.ShowNewFolderButton = true;
            if (FolderDialogue.ShowDialog() == DialogResult.OK)
            {
                OutputFolderBox.Text = FolderDialogue.SelectedPath;
                OutputDir = OutputFolderBox.Text;
                Registry.SetValue(keyName, "OutputDir", OutputDir);
            }
        }
        private void OutputFileSuffixBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void OutputDirBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            OutputDir = theBox.Text;
            Registry.SetValue(keyName, "OutputDir", OutputDir);
            if (theBox.Text != "" && openInputDialogue.FileNames.Count() > 0)
            {
                btnAnalyse.Enabled = true;
            }

        }

        private void FontListFileBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            btnSaveFontList.Enabled = theBox.Text != "";
        }

        private void StyleListFileBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            btnSaveStyles.Enabled = theBox.Text != "";

        }

        private void BulkErrorListbox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            btnSaveErrorList.Enabled = theBox.Text != "";

        }

        private void ContextCharacterFileBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            if (theBox.Text != "")
            {
                if (File.Exists(theBox.Text))
                {
                    btnCheckContextFile.Enabled = true;
                    AnalyseContext = true;
                }
                else
                {
                    MessageBox.Show(theBox.Text + " not found", "File not found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnCheckContextFile.Enabled = false;
                    AnalyseContext = false;
                }

            }
            else
            {
                AnalyseContext = false;
            }

        }

        private void btnCheckContextFile_Click(object sender, EventArgs e)
        {
            // Check the validity fo the list of targets.
            AnalyseContext = LoadTargets(TargetList, excelApp);
        }
    }
    class CharacterDescriptor
    {
        public string FileName;
        public string Font;
        public string Text;

        // Constructor
        public CharacterDescriptor(string FileName, string Font, string Text)
        {
            this.FileName = FileName;
            this.Font = Font;
            this.Text = Text;
        }
        public CharacterDescriptor(string Font, string Text)
        {
            this.FileName = null;
            this.Font = Font;
            this.Text = Text;
        }

        public CharacterDescriptor(string Text)
        {
            this.FileName = null;
            this.Text = Text;
            this.Font = null;
        }
        public CharacterDescriptor(CharacterDescriptor theCharacterDescriptor)
        {
            this.FileName = theCharacterDescriptor.FileName;
            this.Font = theCharacterDescriptor.Font;
            this.Text = theCharacterDescriptor.Text;
        }
        public CharacterDescriptor()
        {
            this.FileName = null;
            this.Font = null;
            this.Text = null;
        }
    }
    class ContextDescriptor
    {
        public string FileName;
        public string Font;
        public string Target;
        public string Context;

        // Constructor
        public ContextDescriptor(string FileName, string FontName, string Target, string Context)
        {
            this.FileName = FileName;
            this.Font = FontName;
            this.Target = Target;
            this.Context = Context;
        }
        public ContextDescriptor(string FontName, string Target, string Context)
        {
            this.FileName = null;
            this.Font = FontName;
            this.Target = Target;
            this.Context = Context;
        }

        public ContextDescriptor(string Target, string Context)
        {
            this.FileName = null;
            this.Font = null;
            this.Target = Target;
            this.Context = Context;
        }
        public ContextDescriptor(ContextDescriptor theContextDescriptor)
        {
            this.FileName = theContextDescriptor.FileName;
            this.Font = theContextDescriptor.Font;
            this.Target = theContextDescriptor.Target;
            this.Context = theContextDescriptor.Context;
        }
        public ContextDescriptor()
        {
            this.FileName = null;
            this.Font = null;
            this.Target = null;
            this.Context = null;
        }
    }
    class CharacterEqualityComparer : EqualityComparer<CharacterDescriptor>
    {
        override public bool Equals(CharacterDescriptor key1, CharacterDescriptor key2)
        {
            bool isEqual = (key1.FileName == key2.FileName) & (key1.Font == key2.Font) & (key1.Text == key2.Text);
            return isEqual;
        }
        override public int GetHashCode(CharacterDescriptor key)
        {
            int HashCode = 0;
            if (key.Font == "")
            {
                HashCode = (key.FileName + "\r" + key.Text).GetHashCode();
            }
            else
            {
                HashCode = (key.FileName + "\r" + key.Text + "\r" + key.Font).GetHashCode();
            }
            return HashCode;
        }
    }
    class ContextEqualityComparer : EqualityComparer<ContextDescriptor>
    {
        override public bool Equals(ContextDescriptor key1, ContextDescriptor key2)
        {
            bool isEqual = (key1.FileName == key2.FileName) & (key1.Font == key2.Font) & (key1.Target == key2.Target) & (key1.Context == key2.Context);
            return isEqual;
        }
        override public int GetHashCode(ContextDescriptor key)
        {
            int HashCode = 0;
            if (key.Font == "")
            {
                HashCode = (key.FileName + "\r" + key.Target + "\r" + key.Context).GetHashCode();
            }
            else
            {
                HashCode = (key.FileName + "\r" + key.Target + "\r" + key.Context + "\r" + key.Font).GetHashCode();
            }
            return HashCode;
        }
    }
    class Style
    {
        public string Name;
        public string Font;

        // Constructor
        public Style(string Name, string Font)
        {
            this.Font = Font;
            this.Name = Name;
        }
        public Style(Style theStyle)
        {
            this.Font = theStyle.Font;
            this.Name = theStyle.Name;
        }

    }
    class StyleComparer : EqualityComparer<Style>
    {
        override public bool Equals(Style key1, Style key2)
        {
            bool isEqual = (key1.Name == key2.Name) & (key1.Font == key2.Font);
            return isEqual;
        }
        override public int GetHashCode(Style key)
        {
            int HashCode = 0;
            if (key.Font == "")
            {
                HashCode = key.Name.GetHashCode();
            }
            else
            {
                HashCode = (key.Name + "\r" + key.Font).GetHashCode();
            }
            return HashCode;
        }
    }
    class TargetDescriptor
    {
        public string Target;
        public bool Valid;
        public int CharactersBefore;
        public int CharactersAfter;

        public TargetDescriptor()
        {
            this.Target = null;
            this.CharactersBefore = 0;
            this.CharactersAfter = 0;
            this.Valid = false;
        }
        public TargetDescriptor(string Target, object CharactersBefore, object CharactersAfter)
        {
            this.Valid = true; // Assume success
            this.Target = this.GetTarget(Target, ref this.Valid);
            this.CharactersAfter = GetInteger(CharactersAfter, ref this.Valid);
            this.CharactersBefore = GetInteger(CharactersBefore, ref this.Valid);
            this.Valid = (this.Target != null);
        }
        public string GetTarget(string Targets, ref bool valid)
        {
            const string pattern = "U\\+(([0-9]|[A-F]){4})"; // Match U+xxxx where xxxx is a four digit hex number
            string hexnumber = "";

            string[] TargetArray = Regex.Split(Targets.ToUpper().Trim(), " "); // We also make sure it is in upper case

            char[] theCharacter = new char[1]; // a single valued array
            string result = "";
            foreach (string Target in TargetArray)
            {
                if (Regex.IsMatch(Target, pattern))
                {
                    hexnumber = Regex.Replace(Target, pattern, "$1");
                    theCharacter[0] = (char)int.Parse(hexnumber, NumberStyles.HexNumber);
                    result += new string(theCharacter);
                    valid = true && valid;
                }
                else
                {
                    MessageBox.Show(Target + " is not a valid Unicode value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    valid = false;
                }
            }
            return result;
        }
        private int GetInteger( object TheValue, ref bool valid)
        {
            int result = -1;
            valid = false; // assume failure
            if (TheValue == null)
            {
                MessageBox.Show("Null value found - check spreadsheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return result;
            }
            try
            {
                if (TheValue.GetType() == typeof(double))
                {
                    result = (int)(double)TheValue;
                    valid = true;
                }
                else
                {
                    MessageBox.Show(TheValue + " is not a valid number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    valid = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                valid = false;
            }

            return result;
        }
        public string GetRegEx(string theTarget, int CharactersBefore, int CharactersAfter)
        {
            // Build the regular expression to search for context
            string result = String.Format(".{{0,{0}}}{1}.{{0,{2}}}", CharactersBefore, GetHex(theTarget, "\\u"), CharactersAfter);
            return result;

        }
        public string GetRegEx(TargetDescriptor theTarget)
        {
            return GetRegEx(theTarget.Target, theTarget.CharactersBefore, theTarget.CharactersAfter);
        }
        public string GetRegEx()
        {
            return GetRegEx(this);
        }
        public string GetHex(string theString, string prefix="", string suffix = "")
        {
            // Get the hexadecimal strings for the characters in a string
            // and return the resultant string
            char[] theCharacters = theString.ToCharArray();
            int theCounter = 0;
            string output = "";
            foreach (char theCharacter in theCharacters)
            {
                int value = Convert.ToInt32(theCharacter);
                output += prefix + value.ToString("X4");
                if (theCounter < theCharacters.Length)
                {
                    output += suffix;
                }
            }
            return output;
        }

    }
}






