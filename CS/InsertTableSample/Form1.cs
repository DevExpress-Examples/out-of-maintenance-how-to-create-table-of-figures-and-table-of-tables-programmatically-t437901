using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office;

namespace InsertTableSample
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void nbiTableOfFigures_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            PrepareDocumentForFigures();
            InsertTableOfEntries("Image");
        }

        private void nbiTableOfTables_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            PrepareDocumentForTables();
            InsertTableOfEntries("Table");
        }

        #region #TOCInsertion
        private void InsertTableOfEntries(String key)
        {
            Document document = richEditControl1.Document;
            document.BeginUpdate();
            Field field = document.Fields.Create(document.Range.Start, string.Format("TOC \\h \\c \"{0}\"", key));
            field.Update();
            document.Fields.Update();
            document.EndUpdate();
        }
        #endregion #TOCInsertion

        #region #InitialDocumentGeneration
        private void PrepareDocumentForFigures()
        {
            richEditControl1.CreateNewDocument();
            Document document = richEditControl1.Document;
            document.BeginUpdate();

            document.AppendText(Characters.PageBreak.ToString());
            document.AppendText("Images:\r\n");

            for (int i = 0; i < imageCollection1.Images.Count; i++)
            {
                // Insert the caption
                document.AppendText("Image ");
                // Insert the SEQ field
                Field field = document.Fields.Create(document.Range.End, "SEQ  Image \\* ARABIC");
                document.Images.Append(imageCollection1.Images[i].Clone() as Image);
                document.Paragraphs.Append();
            }
            //Update the inserted field
            document.Fields.Update();
            document.EndUpdate();
        }
        #endregion #InitialDocumentGeneration

        #region #PrepareDocumentForTables
        private void PrepareDocumentForTables()
        {
            richEditControl1.CreateNewDocument();
            Document document = richEditControl1.Document;
            document.BeginUpdate();
            document.AppendText(Characters.PageBreak.ToString());
            document.AppendText("Tables:\r\n");

            for (int i = 0; i < 3; i++)
            {
                if (i > 0)
                document.AppendText(Characters.PageBreak.ToString());
                document.AppendText("Table ");
                Field field = document.Fields.Create(document.Range.End, "SEQ Table \\* ARABIC");
                CreateTable(document);
            }
                        
            document.Fields.Update();
            document.EndUpdate();
        }

        private Table CreateTable(Document document)
        {
            Random random = new Random();
            Table table = document.Tables.Create(document.Range.End, random.Next(10) + 1, random.Next(5) + 1, AutoFitBehaviorType.AutoFitToWindow);
            table.ForEachCell((cell, rowIndex, cellIndex) =>
            {
                document.InsertText(cell.Range.Start, string.Format("Row {0}, Column {1}", rowIndex, cellIndex));
            });

            return table;
        }
        #endregion #PrepareDocumentForTables
    }
}
