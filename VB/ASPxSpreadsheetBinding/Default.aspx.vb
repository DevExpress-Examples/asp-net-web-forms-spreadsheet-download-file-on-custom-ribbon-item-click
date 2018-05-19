Imports System
Imports System.Data
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports DevExpress.Spreadsheet
Imports DevExpress.Web.Office
Imports DevExpress.Web
Imports System.IO

Namespace ASPxSpreadsheetBinding
    Partial Public Class [Default]
        Inherits System.Web.UI.Page

        Private SessionKey As String = "EditedDocuemntID"
        Protected Property EditedDocuemntID() As String
            Get
                Return If(DirectCast(Session(SessionKey), String), String.Empty)
            End Get
            Set(ByVal value As String)
                Session(SessionKey) = value
            End Set
        End Property
        Protected Sub Page_Init(ByVal sender As Object, ByVal e As EventArgs)
            If Not IsPostBack Then
                SetupSpreadsheetRibbon()
                If Not String.IsNullOrEmpty(EditedDocuemntID) Then
                    CloseSpreadsheetDocument()
                End If
                OpenDocumentFromDatabase()
            End If
            ExportExcelDocument()
        End Sub

        Private Sub ExportExcelDocument()
            If Request("__EVENTTARGET") = "DownloadExcel" Then
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
            Spreadsheet.RibbonTabs.Find(Function(p) p.Text = "File").Visible = False
            Dim item As New RibbonButtonItem("Download", "Download")
            item.Size = RibbonItemSize.Large
            item.LargeImage.IconID = DevExpress.Web.ASPxThemes.IconID.ActionsDownload32x32
            Spreadsheet.RibbonTabs(1).Groups(0).Items.Insert(0, item)
        End Sub

        #Region "open/close documents from a database"
        Private Sub OpenDocumentFromDatabase()
            EditedDocuemntID = Guid.NewGuid().ToString()
            Dim view As DataView = DirectCast(SqlDataSource1.Select(DataSourceSelectArguments.Empty), DataView)
            Spreadsheet.Open(EditedDocuemntID, DocumentFormat.Xlsx, Function()
                Return CType(view.Table.Rows(0)("DocBytes"), Byte())
            End Function)
        End Sub

        Private Sub CloseSpreadsheetDocument()
            DocumentManager.CloseDocument(DocumentManager.FindDocument(EditedDocuemntID).DocumentId)
            EditedDocuemntID = String.Empty
        End Sub
        #End Region

    End Class
End Namespace