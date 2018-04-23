Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.Office

Namespace InsertTableSample
    Partial Public Class Form1
        Inherits DevExpress.XtraEditors.XtraForm

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub nbiTableOfFigures_LinkClicked(ByVal sender As Object, ByVal e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles nbiTableOfFigures.LinkClicked
            PrepareDocumentForFigures()
            InsertTableOfEntries("Image")
        End Sub

        Private Sub nbiTableOfTables_LinkClicked(ByVal sender As Object, ByVal e As DevExpress.XtraNavBar.NavBarLinkEventArgs) Handles nbiTableOfTables.LinkClicked
            PrepareDocumentForTables()
            InsertTableOfEntries("Table")
        End Sub

        #Region "#TOCInsertion"
        Private Sub InsertTableOfEntries(ByVal key As String)
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()
            Dim field As Field = document.Fields.Create(document.Range.Start, String.Format("TOC \h \c ""{0}""", key))
            field.Update()
            document.Fields.Update()
            document.EndUpdate()
        End Sub
        #End Region ' #TOCInsertion

        #Region "#InitialDocumentGeneration"
        Private Sub PrepareDocumentForFigures()
            richEditControl1.CreateNewDocument()
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()

            document.AppendText(Characters.PageBreak.ToString())
            document.AppendText("Images:" & ControlChars.CrLf)

            For i As Integer = 0 To imageCollection1.Images.Count - 1
                ' Insert the caption
                document.AppendText("Image ")
                ' Insert the SEQ field
                Dim field As Field = document.Fields.Create(document.Range.End, "SEQ  Image \* ARABIC")
                document.Images.Append(TryCast(imageCollection1.Images(i).Clone(), Image))
                document.Paragraphs.Append()
            Next i
            'Update the inserted field
            document.Fields.Update()
            document.EndUpdate()
        End Sub
        #End Region ' #InitialDocumentGeneration

        #Region "#PrepareDocumentForTables"
        Private Sub PrepareDocumentForTables()
            richEditControl1.CreateNewDocument()
            Dim document As Document = richEditControl1.Document
            document.BeginUpdate()
            document.AppendText(Characters.PageBreak.ToString())
            document.AppendText("Tables:" & ControlChars.CrLf)

            For i As Integer = 0 To 2
                If i > 0 Then
                document.AppendText(Characters.PageBreak.ToString())
                End If
                document.AppendText("Table ")
                Dim field As Field = document.Fields.Create(document.Range.End, "SEQ Table \* ARABIC")
                CreateTable(document)
            Next i

            document.Fields.Update()
            document.EndUpdate()
        End Sub

        Private Function CreateTable(ByVal document As Document) As Table
            Dim random As New Random()
            Dim table As Table = document.Tables.Create(document.Range.End, random.Next(10) + 1, random.Next(5) + 1, AutoFitBehaviorType.AutoFitToWindow)
            table.ForEachCell(Sub(cell, rowIndex, cellIndex) document.InsertText(cell.Range.Start, String.Format("Row {0}, Column {1}", rowIndex, cellIndex)))

            Return table
        End Function
        #End Region ' #PrepareDocumentForTables
    End Class
End Namespace
