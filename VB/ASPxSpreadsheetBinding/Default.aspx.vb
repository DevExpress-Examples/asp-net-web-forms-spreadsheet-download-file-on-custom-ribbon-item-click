Imports System
Imports System.Data
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports DevExpress.Spreadsheet
Imports DevExpress.Web.Office
Imports DevExpress.Web
Imports System.IO

Namespace ASPxSpreadsheetBinding

    Public Partial Class [Default]
        Inherits Page

        Private SessionKey As String = "EditedDocuemntID"

        Protected Property EditedDocuemntID As String
            Get
                Return If(CStr(Session(SessionKey)), String.Empty)
            End Get

            Set(ByVal value As String)
                Session(SessionKey) = value
            End Set
        End Property

        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            If Not IsPostBack Then
                SetupSpreadsheetRibbon()
                If Not String.IsNullOrEmpty(EditedDocuemntID) Then CloseSpreadsheetDocument()
                OpenDocumentFromDatabase()
            End If

            ExportExcelDocument()
        End Sub

        Private Sub ExportExcelDocument()
            If Equals(Request("__EVENTTARGET"), "DownloadExcel") Then
                Using stream = New MemoryStream()
                    Spreadsheet.Document.SaveDocument(stream, DocumentFormat.Xlsx)
                    Response.Clear()
                    Response.ContentType = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    Response.AddHeader("Content-Disposition", "attachment; filename=document.xlsx")
                    Response.AddHeader("Content-Length", stream.Length.ToString())
                    Response.BinaryWrite(stream.ToArray())
                    Response.Flush()
                    Response.Close()
                    Response.End()
                End Using
            End If
        End Sub

        Private Sub SetupSpreadsheetRibbon()
            Spreadsheet.CreateDefaultRibbonTabs(True)
            Spreadsheet.RibbonTabs.Find(CType(Function(p) Equals(p.Text, "File"), Predicate(Of RibbonTab))).Visible = False
            Dim item As RibbonButtonItem = New RibbonButtonItem("Download", "Download")
            item.Size = RibbonItemSize.Large
            item.LargeImage.IconID = ASPxThemes.IconID.ActionsDownload32x32
            Spreadsheet.RibbonTabs(1).Groups(0).Items.Insert(0, item)
        End Sub

#Region "open/close documents from a database"
        Private Sub OpenDocumentFromDatabase()
            EditedDocuemntID = Guid.NewGuid().ToString()
            Dim view As DataView = CType(SqlDataSource1.Select(DataSourceSelectArguments.Empty), DataView)
            Spreadsheet.Open(EditedDocuemntID, DocumentFormat.Xlsx, Function() CType(view.Table.Rows(0)("DocBytes"), Byte()))
        End Sub

        Private Sub CloseSpreadsheetDocument()
            Call DocumentManager.CloseDocument(DocumentManager.FindDocument(EditedDocuemntID).DocumentId)
            EditedDocuemntID = String.Empty
        End Sub
#End Region
    End Class
End Namespace
