<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/134059990/18.1.10%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T623622)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Spreadsheet for ASP.NET Web Forms - How to download a document on a custom ribbon item click
<!-- run online -->
**[[Run Online]](https://codecentral.devexpress.com/134059990/)**
<!-- run online end -->

This example illustrates how to download a file on a custom ribbon item click.

## Overview

The item is added in code behind to the existing ribbon tabs. The tabs are generated by the [ASPxSpreadsheet.CreateDefaultRibbonTabs](https://docs.devexpress.com/AspNet/DevExpress.Web.ASPxSpreadsheet.ASPxSpreadsheet.CreateDefaultRibbonTabs(System.Boolean)) method:

```cs
  Spreadsheet.CreateDefaultRibbonTabs(true);
  RibbonButtonItem item = new RibbonButtonItem("Download", "Download");
  Spreadsheet.RibbonTabs[1].Groups[0].Items.Insert(0, item);
```

Since it is not possible to download a document on a callback request, a ribbon item should send a postback. For this, use the [__doPostBack](https://www.codeproject.com/Articles/667531/doPostBack-function) function in the client-side [ASPxClientSpreadsheet.CustomCommandExecuted](https://docs.devexpress.com/AspNet/js-ASPxClientSpreadsheet.CustomCommandExecuted) event handler:

```js
function onCustomCommandExecuted(s,e) {
  if (e.commandName == "Download") 
      __doPostBack("DownloadExcel");
  }
```

ASPxSpreadsheet in this example loads data from a database, so the default file dialogs are disabled.

Note that a cell doesn't get the value an end-user entered until the user clicks Enter or tabs out of this cell. As a result the issue with saving an entered value may occur if the value was not submitted. To force submission of the last entered value, call the [ASPxClientSpreadsheet.ApplyCellEdit](https://docs.devexpress.com/AspNet/js-ASPxClientSpreadsheet.ApplyCellEdit) method.

## Documentation

* [Why it is impossible to download files on callbacks](https://docs.devexpress.com/AspNet/403753/troubleshooting/server-side-issues/impossible-download-files-on-callback)
* [Spreadsheet Document Management](https://docs.devexpress.com/AspNet/116406/components/spreadsheet/document-management)
