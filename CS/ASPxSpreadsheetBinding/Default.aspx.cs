using System;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using DevExpress.Spreadsheet;
using DevExpress.Web.Office;
using DevExpress.Web;
using System.IO;

namespace ASPxSpreadsheetBinding {
    public partial class Default : System.Web.UI.Page {
        string SessionKey = "EditedDocuemntID";
        protected string EditedDocuemntID {
            get { return (string)Session[SessionKey] ?? string.Empty; }
            set { Session[SessionKey] = value; }
        }
        protected void Page_Init(object sender, EventArgs e) {
            if (!IsPostBack) {
                SetupSpreadsheetRibbon();
                if (!string.IsNullOrEmpty(EditedDocuemntID))
                    CloseSpreadsheetDocument();
                OpenDocumentFromDatabase();
            }
            ExportExcelDocument();
        }

        void ExportExcelDocument() {
            if (Request["__EVENTTARGET"] == "DownloadExcel") {
                using (var stream = new MemoryStream()) {
                    Spreadsheet.Document.SaveDocument(stream, DocumentFormat.Xlsx);
                    Response.Clear();
                    Response.ContentType = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("Content-Disposition", "attachment; filename=document.xlsx");
                    Response.AddHeader("Content-Length", stream.Length.ToString());
                    Response.BinaryWrite(stream.ToArray());
                    Response.Flush();
                    Response.Close();
                    Response.End();
                }
            }
        }
        void SetupSpreadsheetRibbon() {
            Spreadsheet.CreateDefaultRibbonTabs(true);
            Spreadsheet.RibbonTabs.Find(p => p.Text == "File").Visible = false;
            RibbonButtonItem item = new RibbonButtonItem("Download", "Download");
            item.Size = RibbonItemSize.Large;
            item.LargeImage.IconID = DevExpress.Web.ASPxThemes.IconID.ActionsDownload32x32;
            Spreadsheet.RibbonTabs[1].Groups[0].Items.Insert(0, item);
        }

        #region open/close documents from a database
        void OpenDocumentFromDatabase() {
            EditedDocuemntID = Guid.NewGuid().ToString();
            DataView view = (DataView)SqlDataSource1.Select(DataSourceSelectArguments.Empty);
            Spreadsheet.Open(
                EditedDocuemntID,
                DocumentFormat.Xlsx,
                () => { return (byte[])view.Table.Rows[0]["DocBytes"]; }
            );
        }

        void CloseSpreadsheetDocument() {
            DocumentManager.CloseDocument(DocumentManager.FindDocument(EditedDocuemntID).DocumentId);
            EditedDocuemntID = string.Empty;
        }
        #endregion

    }
}