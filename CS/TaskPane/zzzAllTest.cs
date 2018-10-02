using System;
using System.Linq;
using System.Xml;
using System.Windows.Forms;
using System.Reflection;
using System.Collections.Generic;
using System.Collections.Specialized;
using Word = Microsoft.Office.Interop.Word;

namespace EditTools.TaskPane
{
    /// <summary>
    /// Comments TaskPane
    /// </summary>
    public partial class zzzAllTest : UserControl
    {
        /// <summary>
        /// Initialize the controls in the object
        /// </summary>
        public zzzAllTest()
        {
            InitializeComponent();
            LoadTaskPane();
        }

        private void LoadTaskPane()
        {
            try
            {
                dgvComments.Columns.Add("col_Name", "Characters");
                dgvComments.Columns.Add("col_Value", "Comment Text");
                dgvComments.Columns["col_Value"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgvComments.Columns["col_Value"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dgvComments.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

                StringCollection strings = Properties.Settings.Default.Options_StandardComments;
                Dictionary<string, string> dict = new Dictionary<string, string>();

                for (int i = 0; i < strings.Count - 1; i += 2)
                {
                    if ((strings[i] != null) && (strings[i + 1] != null))
                    {
                        dict.Add(strings[i], strings[i + 1]);
                    }
                }

                foreach (string key in dict.Keys)
                {
                    string[] row = new string[] { key, dict[key] };
                    dgvComments.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                //ErrorHandler.DisplayMessage(ex);

            }
        }

        private void tsbSave_Click(object sender, EventArgs e)
        {
            try
            {
                StringCollection strings = new StringCollection();

                foreach (DataGridViewRow row in dgvComments.Rows)
                {
                    if (((string)row.Cells[0].Value != null) && ((string)row.Cells[1].Value != null))
                    {
                        strings.Add((string)row.Cells[0].Value);
                        strings.Add((string)row.Cells[1].Value);
                    }
                }
                Properties.Settings.Default.Options_StandardComments.Clear();
                Properties.Settings.Default.Options_StandardComments = strings;

                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                //ErrorHandler.DisplayMessage(ex);

            }
        }

        private void tsbImport_Click(object sender, EventArgs e)
        {
            ImportDialog id = new ImportDialog();
            id.ShowDialog();
        }

        private void tsbExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.AddExtension = true;
            sfd.DefaultExt = "xml";
            sfd.CheckPathExists = true;
            sfd.Filter = "XML files (*.xml)|*.xml";
            sfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            sfd.OverwritePrompt = true;
            sfd.Title = "Export Standard Comments";
            sfd.FileName = "";
            //sfd.RestoreDirectory = false;

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (XmlWriter writer = XmlWriter.Create(sfd.FileName))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("boilerplate");
                    StringCollection sc = Properties.Settings.Default.Options_StandardComments;
                    for (int i = 0; i < sc.Count - 1; i += 2)
                    {
                        writer.WriteStartElement("entry");
                        writer.WriteAttributeString("key", sc[i]);
                        writer.WriteString(sc[i + 1]);
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                MessageBox.Show("Settings exported to " + sfd.FileName + ".");
            }
            else
            {
                MessageBox.Show("Export cancelled.");
            }
        }

        private void tsbDecreaseDistance_Click(object sender, EventArgs e)
        {
            //Properties.Settings.Default.Options_ProperNounDistanceMin -= 1;
            //tstDistance.Text = Properties.Settings.Default.Options_ProperNounDistanceMin.ToString();
        }

        private void tsbIncreaseDistance_Click(object sender, EventArgs e)
        {
            //Properties.Settings.Default.Options_ProperNounDistanceMin += 1;
            //tstDistance.Text = Properties.Settings.Default.Options_ProperNounDistanceMin.ToString();
        }

        private void tsbApplyAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvComments.Rows)
            {
                string commentString = row.Cells["col_Name"].Value.ToString();
                string commentText = row.Cells["col_Value"].Value.ToString();
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

        private void tsbRemoveAll_Click(object sender, EventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            doc.DeleteAllComments();
        }

        private void dgvComments_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            string id = dgvComments[0, e.RowIndex].Value.ToString();
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Find find = doc.Content.Find;
            find.ClearHitHighlight();
            find.HitHighlight(FindText: id, MatchCase: false, HighlightColor: Word.WdColor.wdColorYellow);
        }

        private void dgvComments_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Find find = doc.Content.Find;
            find.ClearHitHighlight();
        }
    }
}
