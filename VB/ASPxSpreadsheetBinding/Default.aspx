<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="ASPxSpreadsheetBinding.Default" %>
<%@ Register Assembly="DevExpress.Web.v17.2, Version=17.2.7.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.ASPxSpreadsheet.v17.2, Version=17.2.7.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxSpreadsheet" TagPrefix="dx" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script>
        function onCustomCommandExecuted(s,e) {
            if (e.commandName == "Download") {
                __doPostBack("DownloadExcel");
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <dx:ASPxSpreadsheet ID="Spreadsheet" runat="server" >
            <ClientSideEvents CustomCommandExecuted="onCustomCommandExecuted" />
        </dx:ASPxSpreadsheet>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DocumentsConnectionString %>" 
            SelectCommand="SELECT * FROM [Docs]">
        </asp:SqlDataSource>
    </form>
</body>
</html>