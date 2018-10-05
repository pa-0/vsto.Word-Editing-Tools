using System;
using System.Data;
using System.Drawing;
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
    public partial class Phrases : UserControl
    {
        /// <summary>
        /// Initialize the controls in the object
        /// </summary>
        public Phrases()
        {
            InitializeComponent();
            LoadTaskPane();
        }

        private void LoadTaskPane()
        {
            try
            {
                DataTable dt = new DataTable("Words");
                dt.Columns.Add("Value", System.Type.GetType("System.String"));
                dt.Columns.Add("#", System.Type.GetType("System.Int32"));

                Dictionary<string, uint> wordlist = new Dictionary<string, uint>();
                wordlist = Ribbon.GetWordFrequencyList();

                var words = wordlist.ToList();
                words.Sort((pair1, pair2) => pair2.Value.CompareTo(pair1.Value));

                foreach (var pair in words)
                {
                    dt.Rows.Add(new object[] { pair.Key, pair.Value });
                }

                dgvComments.DataSource = dt;
                dgvComments.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dgvComments.ReadOnly = true;
                dgvComments.AllowUserToAddRows = false;
                dgvComments.MultiSelect = false;

            }
            catch (Exception ex)
            {
                //ErrorHandler.DisplayMessage(ex);

            }
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

        private void dgvComments_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            string id = dgvComments[0, e.RowIndex].Value.ToString();
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Find find = doc.Content.Find;
            find.ClearHitHighlight();
            find.HitHighlight(FindText: id, MatchCase: true, HighlightColor: Word.WdColor.wdColorYellow, MatchWholeWord: true);
        }

        private void dgvComments_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Find find = doc.Content.Find;
            find.ClearHitHighlight();

        }
    }
}
