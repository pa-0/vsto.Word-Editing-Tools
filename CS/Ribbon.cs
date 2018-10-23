using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
//using System.Text;
//using System.Windows.Forms;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using edu.stanford.nlp.tagger.maxent;
using edu.stanford.nlp.ling;
using java.util;
using EditTools.Scripts;

namespace EditTools
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public edu.stanford.nlp.tagger.maxent.MaxentTagger tagger;
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region | Task Panes |

        /// <summary>
        /// Settings TaskPane
        /// </summary>
        public TaskPane.Settings mySettings;

        /// <summary>
        /// Comments TaskPane
        /// </summary>
        public TaskPane.Comments myComments;

        /// <summary>
        /// Words TaskPane
        /// </summary>
        public TaskPane.Words myWords;

        /// <summary>
        /// Settings Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;

        /// <summary>
        /// Comments Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneComments;

        /// <summary>
        /// Words Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneWords;

        #endregion

        #region | Ribbon Events |

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("EditTools.Ribbon.xml");
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            //Properties.Settings.Default.Options_ProofLanguageID = (Microsoft.Office.Interop.Word.WdLanguageID)Word.Language;
        }

        /// <summary> 
        /// Assigns text to a label on the ribbon from the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for a label. </returns> 
        public string GetLabelText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tabEditingTools":
                        //if (Application.ProductVersion.Substring(0, 2) == "15") //for Word 2013
                        //{
                        //    return AssemblyInfo.Title.ToUpper();
                        //}
                        //else
                        //{
                        return AssemblyInfo.Title;
                    //}
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        DateTime createDate = Properties.Settings.Default.App_ReleaseDate;
                        return createDate.ToString("dd-MMM-yyyy hh:mm tt");
                    case "txtMinPhraseLen":
                        return Properties.Settings.Default.Options_PhraseLengthMin.ToString();
                    case "txtMaxPhraseLen":
                        return Properties.Settings.Default.Options_PhraseLengthMax.ToString();
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary> 
        /// Assigns an image to a button on the ribbon in the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        public System.Drawing.Bitmap GetButtonImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnCopyrightLogo":
                        return Properties.Resources.logo;
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;
            }
        }

        /// <summary>
        /// Assigns the value to an application setting
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                //Ribbon.AppVariables.ControlLabel = GetLabelText(control);
                switch (control.Id)
                {

                    case "btnApplyLanguage":
                        ApplyLanguage();
                        break;
                    case "btnApplyComments":
                        ApplyComments();
                        break;
                    case "btnCommentList":
                        OpenComments();
                        break;
                    case "btnSingularData":
                        SingularData();
                        break;
                    case "btnProperNouns":
                        ProperNouns();
                        break;
                    case "btnWords":
                        WordFrequencyList();
                        break;
                    case "btnWordsList":
                        OpenWords();
                        break;
                    case "btnWordFrequencyList":
                        WordFrequencyList();
                        break;
                    case "btnPhrases":
                        PhraseList();
                        break;
                    case "btnAcceptChanges":
                        AcceptChanges();
                        break;
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        public string GetContent(Office.IRibbonControl control)
        {
            string dynamicMenu = string.Empty;
            dynamicMenu = @" < menu xmlns = ""http://schemas.microsoft.com/office/2009/07/customui"" >";
            dynamicMenu += @" < button id = ""button1"" label = ""Button 1"" />";
            dynamicMenu += @" < button id = ""button2"" label = ""Button 2"" />";
            dynamicMenu += @" < button id = ""button3"" label = ""Button 3"" />";
            dynamicMenu += " </ menu >";
            return dynamicMenu;

        }

        #endregion

        #region | Ribbon Buttons |

        /// <summary> 
        /// Opens the comments taskpane
        /// </summary>
        /// <remarks></remarks>
        public void OpenComments()
        {
            try
            {
                if (myTaskPaneComments != null)
                {
                    if (myTaskPaneComments.Visible == true) //it doesn't like this line if already visible and you open a document
                    {
                        myTaskPaneComments.Visible = false;
                    }
                    else
                    {
                        myTaskPaneComments.Visible = true;
                    }
                }
                else
                {
                    myComments = new TaskPane.Comments();
                    myTaskPaneComments = Globals.ThisAddIn.CustomTaskPanes.Add(myComments, "Comments");
                    myTaskPaneComments.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneComments.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneComments.Width = 675;
                    myTaskPaneComments.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens the words taskpane
        /// </summary>
        /// <remarks></remarks>
        public void OpenWords()
        {
            try
            {
                if (myTaskPaneWords != null)
                {
                    if (myTaskPaneWords.Visible == true) //it doesn't like this line if already visible and you open a document
                    {
                        myTaskPaneWords.Visible = false;
                    }
                    else
                    {
                        myTaskPaneWords.Visible = true;
                    }
                }
                else
                {
                    myWords = new TaskPane.Words();
                    myTaskPaneWords = Globals.ThisAddIn.CustomTaskPanes.Add(myWords, "Words");
                    myTaskPaneWords.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneWords.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneWords.Width = 325;
                    myTaskPaneWords.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens the settings taskpane
        /// </summary>
        /// <remarks></remarks>
        public void OpenSettings()
        {
            try
            {
                if (myTaskPaneSettings != null)
                {
                    if (myTaskPaneSettings.Visible == true)
                    {
                        myTaskPaneSettings.Visible = false;
                    }
                    else
                    {
                        myTaskPaneSettings.Visible = true;
                    }
                }
                else
                {
                    mySettings = new TaskPane.Settings();
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings");
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneSettings.Width = 675;
                    myTaskPaneSettings.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenReadMe()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenNewIssue()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathNewIssue);

        }

        public void ApplyLanguage()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Properties.Settings.Default.Save();
            foreach (Word.Range rng in TextHelpers.GetText(doc))
            {
                rng.LanguageID = Properties.Settings.Default.Options_ProofLanguageID;
                rng.NoProofing = 0;
            }
            if (Properties.Settings.Default.Options_DisplayLanguageDialog) { Globals.ThisAddIn.Application.CommandBars.ExecuteMso("SetLanguage"); }
        }

        static public void ApplyComments(string comment = "")
        {
            Properties.Settings.Default.Save();
            StringCollection comments = Properties.Settings.Default.Options_StandardComments;
            for (int i = 0; i < comments.Count - 1; i++)
            {
                string commentString = comments[i];
                string commentText = comments[i + 1];
                if (comment == "" || comment == commentString)
                {
                    Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                    Word.Range rng = doc.Content;
                    rng.Find.Forward = true;
                    rng.Find.Text = commentString;
                    rng.Find.Execute();
                    while (rng.Find.Found)
                    {
                        rng.Select();
                        Word.Selection selection = Globals.ThisAddIn.Application.Selection;
                        if (selection.Comments.Count == 0)  // don't create the comment if it already exists
                        {
                            selection.Comments.Add(selection.Range, commentText);
                            rng.Find.Execute();
                        }
                    }
                }
            }

        }

        public void SingularData()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs pgraphs;
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            bool fromSelection = false;
            if (sel != null && sel.Range != null && sel.Characters.Count > 1)
            {
                pgraphs = sel.Paragraphs;
                fromSelection = true;
            }
            else
            {
                pgraphs = doc.Paragraphs;
            }
            //Debug.WriteLine("From selection: " + fromSelection.ToString());

            ProgressDialog d = new ProgressDialog();
            d.pbMax = pgraphs.Count;
            d.pbVal = 0;
            d.Show();

            foreach (Word.Paragraph pgraph in pgraphs)
            {
                d.pbVal++;
                Word.Range rng = pgraph.Range;
                foreach (Word.Range sentence in rng.Sentences)
                {
                    // POS
                    var tsentence = MaxentTagger.tokenizeText(new java.io.StringReader(sentence.Text)).toArray();
                    var taggedSentence = tagger.tagSentence((ArrayList)tsentence[0]);
                    var taglist = taggedSentence.toArray();
                    Boolean singular = false;
                    Boolean hasdata = false;

                    //First find obviously singular "data"
                    foreach (TaggedWord entry in taglist)
                    {
                        if (entry.word().ToLower() == "data")
                        {
                            hasdata = true;
                            if ((entry.tag() == "NN") || (entry.tag() == "NNP"))
                            {
                                singular = true;
                                break;
                            }
                        }
                    }

                    //Now look for plural tags with singular verbs
                    if ((hasdata) && (!singular))
                    {
                        foreach (TaggedWord entry in taglist)
                        {
                            if (entry.tag() == "VBZ")
                            {
                                singular = true;
                                break;
                            }
                        }
                    }

                    //Highlight problematic sentences
                    if ((hasdata) && (singular))
                    {
                        sentence.HighlightColorIndex = Word.WdColorIndex.wdGray50;
                    }
                }

                //var sentences = MaxentTagger.tokenizeText(new java.io.StringReader(rng.Text)).toArray();
                //foreach (ArrayList sentence in sentences)
                //{
                //    String origsent = String.Join(" ", sentence.toArray());
                //    Debug.WriteLine(origsent);
                //    var taggedSentence = tagger.tagSentence(sentence);
                //    var taglist = taggedSentence.toArray();
                //    foreach (TaggedWord entry in taglist)
                //    {
                //        if (entry.word().ToLower() == "data")
                //        {
                //            if ( (entry.tag() == "NN") || (entry.tag() == "NNP") )
                //            {
                //                Debug.WriteLine("Found singular 'data' in the following sentence: " + origsent);
                //                TextHelpers.highlightText(pgraph.Range, "", Word.WdColorIndex.wdGray50);
                //            }
                //        }
                //    }
                //}

                //// Typed Dependencies
                //foreach (Word.Range sentence in rng.Sentences)
                //{
                //    var tokenizerFactory = PTBTokenizer.factory(new CoreLabelTokenFactory(), "");
                //    var sent2Reader = new java.io.StringReader(sentence.Text);
                //    var rawWords = tokenizerFactory.getTokenizer(sent2Reader).tokenize();
                //    sent2Reader.close();
                //    var tree = lp.apply(rawWords);

                //    var tlp = new PennTreebankLanguagePack();
                //    var gsf = tlp.grammaticalStructureFactory();
                //    var gs = gsf.newGrammaticalStructure(tree);
                //    var tdl = gs.typedDependenciesCCprocessed();
                //    foreach (var dep in tdl.toArray())
                //    {
                //        Debug.WriteLine(dep);
                //    }
                //    Debug.WriteLine("=-=-=-=-=-");
                //}
            }
            d.Hide();
            //MessageBox.Show("Possible uses of 'data' as a singular noun have been highlighted in grey.");
        }

        public void ProperNouns()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            HashSet<string> wordlist = new HashSet<string>();

            foreach (Word.Range rng in TextHelpers.GetText(doc))
            {
                string txt = rng.Text;
                Word.Style style = rng.get_Style();
                if (style != null)
                {
                    Regex re_heading = new Regex(@"(?i)(heading|title|date|toc)");
                    Match m = re_heading.Match(style.NameLocal);
                    if (m.Success)
                    {
                        continue;
                    }
                }

                HashSet<string> propers = new HashSet<string>();
                propers = TextHelpers.ProperNouns(txt);
                wordlist.UnionWith(propers);
            }

            //Produce the groupings
            HashSet<string> capped = TextHelpers.KeepCaps(wordlist);

            //DoubleMetaphone
            Dictionary<ushort, List<string>> mpgroups = new Dictionary<ushort, List<string>>();
            //Dictionary<string, List<string>> mpgroups = new Dictionary<string, List<string>>();
            ShortDoubleMetaphone sdm = new ShortDoubleMetaphone();
            //HashSet<string> tested = new HashSet<string>();

            foreach (string word in capped)
            {
                /*
                if (tested.Contains(word))
                {
                    continue;
                }
                else
                {
                    tested.Add(word);
                }
                */
                sdm.computeKeys(word);
                ushort pri = sdm.PrimaryShortKey;
                ushort alt = sdm.AlternateShortKey;
                if (mpgroups.ContainsKey(pri))
                {
                    mpgroups[pri].Add(word);
                }
                else
                {
                    List<string> node = new List<string>();
                    node.Add(word);
                    mpgroups[pri] = node;
                }
                if (mpgroups.ContainsKey(alt))
                {
                    mpgroups[alt].Add(word);
                }
                else
                {
                    List<string> node = new List<string>();
                    node.Add(word);
                    mpgroups[alt] = node;
                }
            }

            //Edit Distance
            List<string> dtested = new List<string>();
            int mindist;
            int.TryParse(Properties.Settings.Default.Options_ProperNounDistanceMin.ToString(), out mindist);
            if (mindist == 0)
            {
                mindist = 2;
            }
            Dictionary<string, List<string>> distgroups = new Dictionary<string, List<string>>();

            foreach (string word1 in capped)
            {
                if (dtested.Contains(word1))
                {
                    continue;
                }
                else
                {
                    dtested.Add(word1);
                }

                if (word1.Length <= mindist)
                {
                    continue;
                }

                foreach (string word2 in capped)
                {
                    if (word2.Length <= mindist)
                    {
                        continue;
                    }
                    int dist = TextHelpers.EditDistance(word1, word2);
                    //int percent = (int)Math.Round((dist / word1.Length) * 100.0);
                    if ((dist > 0) && (dist <= mindist))
                    //if ((dist > 0) && (percent <= distpercent))
                    {
                        dtested.Add(word2);
                        if (distgroups.ContainsKey(word1))
                        {
                            distgroups[word1].Add(word2);
                        }
                        else
                        {
                            List<string> node = new List<string>();
                            node.Add(word2);
                            distgroups[word1] = node;
                        }
                    }
                }
            }

            //Create new document
            Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add();
            Word.View view = Globals.ThisAddIn.Application.ActiveWindow.View;
            view.DisplayPageBoundaries = false;
            Word.Paragraph pgraph;

            //Intro text
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Heading 1"]);
            pgraph.Range.Text = "Proper Noun Checker\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Normal"]);
            pgraph.Range.Text = "This tool only looks at words that start with a capital letter. It then uses phonetic comparison and edit distance to find other proper nouns that are similar. Words in all caps (acronyms) are not included.\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.Text = "The system tries to ignore words at the beginning of sentences and in headers. This means some errors may go unseen, so use multiple tools!\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.Text = "Most of what you see here are false positives! That's unavoidable. But it still catches certain otherwise-hard-to-find misspellings.\n";

            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            Word.Section sec = newdoc.Sections[2];
            sec.PageSetup.TextColumns.SetCount(2);
            sec.PageSetup.TextColumns.LineBetween = -1;

            //Distance
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Heading 2"]);
            //pgraph.KeepWithNext = 0;
            pgraph.Range.Text = "Edit Distance (" + mindist + ")\n";

            foreach (string key in distgroups.Keys)
            {
                List<string> group = distgroups[key];
                pgraph = newdoc.Content.Paragraphs.Add();
                pgraph.set_Style(newdoc.Styles["Normal"]);
                pgraph.Range.Text = key + ", " + string.Join(", ", group) + "\n";

            }

            pgraph.Range.InsertBreak(Word.WdBreakType.wdPageBreak);
            //pgraph = newdoc.Content.Paragraphs.Add();
            //pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            //Word.InlineShape line = pgraph.Range.InlineShapes.AddHorizontalLineStandard();
            //line.Height = 2;
            //line.Fill.Solid();
            //line.HorizontalLineFormat.NoShade = true;
            //line.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            //line.HorizontalLineFormat.PercentWidth = 90;
            //line.HorizontalLineFormat.Alignment = WdHorizontalLineAlignment.wdHorizontalLineAlignCenter;
            //sec = newdoc.Sections[3];
            //sec.PageSetup.TextColumns.SetCount(2);
            //sec.PageSetup.TextColumns.LineBetween = -1;

            //Metaphone
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Heading 2"]);
            //pgraph.KeepWithNext = 0;
            pgraph.Range.Text = "Phonetic Comparisons\n";

            foreach (ushort key in mpgroups.Keys)
            {
                if (key == 65535)
                {
                    continue;
                }
                List<string> group = mpgroups[key];
                if (group.Count > 1)
                {
                    pgraph = newdoc.Content.Paragraphs.Add();
                    pgraph.set_Style(newdoc.Styles["Normal"]);
                    pgraph.Range.Text = string.Join(", ", group) + "\n";
                }
            }

            //pgraph = newdoc.Content.Paragraphs.Add();
            //pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            newdoc.GrammarChecked = true;
        }

        public void WordList()
        {
            HashSet<string> wordlist = new HashSet<string>();
            wordlist = GetWordList();

            //Create new document
            Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add();
            Word.Paragraph pgraph;

            //Intro text
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Heading 1"]);
            pgraph.Range.Text = "Word List\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Normal"]);
            pgraph.Range.Text = "This is a proofreading tool. It takes every word in the document, strips the punctuation, removes words that consist only of numbers, and then presents them all in alphabetical order. This is a great way to find typos and inconsistencies.\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.Text = "Capitalization is retained as is. That means that words that appear at the beginning of a sentence will appear capitalized.\n";

            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            Word.Section sec = newdoc.Sections[2];
            sec.PageSetup.TextColumns.SetCount(3);

            string[] words = wordlist.ToArray();
            Array.Sort(words);
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.Text = string.Join("\n", words) + "\n";

            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            newdoc.GrammarChecked = true;
        }

        static public HashSet<string> GetWordList()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            HashSet<string> wordlist = new HashSet<string>();

            foreach (Word.Range rng in TextHelpers.GetText(doc))
            {
                string txt = rng.Text;

                //strip punctuation
                txt = TextHelpers.StripPunctuation(txt);

                //get word list
                HashSet<string> newwords = TextHelpers.ToWords(txt);
                wordlist.UnionWith(newwords);
            }

            //strip words that are all numbers
            wordlist = TextHelpers.RemoveNumbers(wordlist);
            return wordlist;

        }

        static public Dictionary<string, uint> GetWordFrequencyList()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Dictionary<string, uint> wordlist = new Dictionary<string, uint>();
            Regex re_allnums = new Regex(@"^\d+$");

            IEnumerable<Word.Range> textranges = TextHelpers.GetText(doc);
            //d.pbMax = textranges.Count();
            //d.pbVal = 0;
            foreach (Word.Range rng in textranges)
            {
                //d.pbVal++;
                //Application.StatusBar = Left("Importing Data... | " & Format(App.EndTime - App.StartTime, "hh:mm:ss") & " | (" & Ribbon.fileNbr & " of " & App.FileTotal & ") " & Format(Ribbon.fileNbr / App.FileTotal, "0.0%") & " | " & filePath, 255)
                //Word.Application.StatusBar = "";
                //Word.Application.StatusBar = "test in status bar";
                string txt = rng.Text;

                //strip punctuation
                txt = TextHelpers.StripPunctuation(txt);

                string[] substrs = Regex.Split(txt, @"\s+");
                foreach (string word in substrs)
                {
                    Match m = re_allnums.Match(word);
                    if (!m.Success)
                    {
                        if (word.Trim() != "")
                        {
                            if (wordlist.ContainsKey(word))
                            {
                                wordlist[word]++;
                            }
                            else
                            {
                                wordlist.Add(word, 1);
                            }
                        }
                    }

                }
            }
            return wordlist;
        }

        public void WordFrequencyList()
        {
            //ProgressDialog d = new ProgressDialog();
            //d.Show();

            Stopwatch watch = new Stopwatch();
            watch.Start();

            Dictionary<string, uint> wordlist = new Dictionary<string, uint>();
            wordlist = GetWordFrequencyList();

            //Debug.WriteLine("Counts tabulated. Time elapsed: " + watch.Elapsed.ToString());
            watch.Restart();

            //Create new document
            Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add();
            Word.Paragraph pgraph;

            //Intro text
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Heading 1"]);
            pgraph.Range.Text = "Word Frequency List\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Normal"]);
            pgraph.Range.Text = "Capitalization is retained as is. That means that words that appear at the beginning of a sentence will appear capitalized. Don't forget that you can sort the table!\n";
            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.set_Style(newdoc.Styles["Normal"]);
            pgraph.Range.Text = "Total words found (case sensitive): " + wordlist.Count.ToString() + "\n";

            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            Word.Section sec = newdoc.Sections[2];
            sec.PageSetup.TextColumns.SetCount(3);

            var words = wordlist.ToList();
            words.Sort((pair1, pair2) => pair2.Value.CompareTo(pair1.Value));
            newdoc.Tables.Add(pgraph.Range, words.Count, 2);
            //newdoc.Tables.Add(pgraph.Range, 1, 2);
            newdoc.Tables[1].AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
            newdoc.Tables[1].AllowAutoFit = true;
            //d.pbMax = words.Count;
            //d.pbVal = 0;
            int row = 1;
            foreach (var pair in words)
            {
                //d.pbVal++;
                //newdoc.Tables[1].Rows.Add();
                Word.Cell cell = newdoc.Tables[1].Cell(row, 1);
                cell.Range.Text = pair.Key;
                cell = newdoc.Tables[1].Cell(row, 2);
                cell.Range.Text = pair.Value.ToString();
                row++;
            }

            pgraph = newdoc.Content.Paragraphs.Add();
            pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
            newdoc.GrammarChecked = true;
            //Debug.WriteLine("All done. Time elapsed: " + watch.Elapsed.ToString());
            watch.Stop();
            //d.Hide();
        }

        public void PhraseList()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            uint newminlen;
            uint newmaxlen;
            UInt32.TryParse(Properties.Settings.Default.Options_PhraseLengthMin.ToString(), out newminlen);
            UInt32.TryParse(Properties.Settings.Default.Options_PhraseLengthMax.ToString(), out newmaxlen);
            if ((newminlen != 0) && (newmaxlen != 0) && (newminlen <= newmaxlen))
            {
                Properties.Settings.Default.Options_PhraseLengthMin = newminlen;
                Properties.Settings.Default.Options_PhraseLengthMax = newmaxlen;
                Properties.Settings.Default.Save();

                Dictionary<string, uint> phrases = new Dictionary<string, uint>();
                //Iterate through all text
                foreach (Word.Range rng in TextHelpers.GetText(doc))
                {
                    //Break into sentences
                    foreach (Word.Range sentence in rng.Sentences)
                    {
                        //Strip punctuation
                        string nopunc = TextHelpers.StripPunctuation(sentence.Text);
                        nopunc = nopunc.Replace("  ", " ");
                        //Break into words
                        string[] words = nopunc.Split(' ');
                        //Extract phrases
                        for (uint i = newminlen; i <= newmaxlen; i++)
                        {
                            for (int start = 0; start < words.Length - i; start++)
                            {
                                List<string> phraselst = new List<string>();
                                for (int idx = 0; idx < i; idx++)
                                {
                                    phraselst.Add(words[start + idx]);
                                }
                                string phrase = string.Join(" ", phraselst).ToLower();
                                //Add to data structre
                                if (phrases.ContainsKey(phrase))
                                {
                                    phrases[phrase]++;
                                }
                                else
                                {
                                    phrases[phrase] = 1;
                                }
                            }
                        }
                    }
                }

                //Display results

                //Create new document
                Word.Document newdoc = Globals.ThisAddIn.Application.Documents.Add();
                Word.Paragraph pgraph;

                //Intro text
                pgraph = newdoc.Content.Paragraphs.Add();
                pgraph.set_Style(newdoc.Styles["Heading 1"]);
                pgraph.Range.Text = "Phrase Frequency List\n";
                pgraph = newdoc.Content.Paragraphs.Add();
                pgraph.set_Style(newdoc.Styles["Normal"]);
                pgraph.Range.Text = "Punctuation (other than apostrophes) has been removed. All words have been lowercased for comparison.\n";

                pgraph = newdoc.Content.Paragraphs.Add();
                pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
                Word.Section sec = newdoc.Sections[2];
                sec.PageSetup.TextColumns.SetCount(2);

                var phraselist = phrases.Where(x => x.Value > 1).ToList();
                phraselist.Sort((pair1, pair2) => pair2.Value.CompareTo(pair1.Value));
                foreach (var pair in phraselist)
                {
                    pgraph = newdoc.Content.Paragraphs.Add();
                    pgraph.set_Style(newdoc.Styles["Normal"]);
                    pgraph.Range.Text = pair.Key + "\t" + pair.Value.ToString() + "\n";
                }

                pgraph = newdoc.Content.Paragraphs.Add();
                pgraph.Range.InsertBreak(Word.WdBreakType.wdSectionBreakContinuous);
                newdoc.GrammarChecked = true;
            }
            else
            {
                //MessageBox.Show("The phrase length fields must contain numbers greater than zero, and the minimum length must be less than or equal to the maximum length.");
            }
        }

        public void AcceptChanges()
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Window win = Globals.ThisAddIn.Application.ActiveWindow;
            Word.View view = win.View;

            view.ShowComments = false;
            view.ShowInkAnnotations = false;
            view.ShowInsertionsAndDeletions = false;

            doc.AcceptAllRevisionsShown();

            view.ShowComments = true;
            view.ShowInkAnnotations = true;
            view.ShowInsertionsAndDeletions = true;
            //MessageBox.Show("Formatting changes have been accepted.");
        }

        #endregion

    }
}
