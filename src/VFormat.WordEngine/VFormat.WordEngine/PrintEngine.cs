using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using VFormat.WordEngine.Interop;

namespace VFormat.WordEngine;

public static class PrintEngine
{
    public static List<string> GetInstalledPrinters(out string defaultPrinter)
    {
        var list = new List<string>();
        defaultPrinter = "";
        try
        {
            var settings = new PrinterSettings();
            defaultPrinter = settings.PrinterName ?? "";
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                list.Add(printer);
            }
        }
        catch
        {
        }

        if (string.IsNullOrEmpty(defaultPrinter) && list.Count > 0)
        {
            defaultPrinter = list[0];
        }
        return list;
    }

    public static void PrintOddPages(Document doc, int copies = 1, bool reverse = false, string printer = null)
    {
        PrintPages(doc, WdPrintOutPages.wdPrintOddPagesOnly, "", copies, reverse, printer);
    }

    public static void PrintEvenPages(Document doc, int copies = 1, bool reverse = false, string printer = null)
    {
        PrintPages(doc, WdPrintOutPages.wdPrintEvenPagesOnly, "", copies, reverse, printer);
    }

    public static void PrintPages(Document doc, WdPrintOutPages pageType, string pages = "", int copies = 1, bool reverse = false, string printer = null)
    {
        if (doc == null)
        {
            throw new ArgumentNullException(nameof(doc), "Tài liệu Word không hợp lệ.");
        }

        string originalPrinter = "";
        try
        {
            if (!string.IsNullOrEmpty(printer) && doc.Application != null)
            {
                originalPrinter = doc.Application.ActivePrinter;
                if (!string.Equals(originalPrinter, printer, StringComparison.OrdinalIgnoreCase))
                {
                    doc.Application.ActivePrinter = printer;
                }
            }
        }
        catch
        {
        }

        object background = false;
        object append = false;
        object range = string.IsNullOrWhiteSpace(pages) ? (object)WdPrintOutRange.wdPrintAllDocument : (object)WdPrintOutRange.wdPrintRangeOfPages;
        object outputFileName = Missing.Value;
        object from = Missing.Value;
        object to = Missing.Value;
        object item = WdPrintOutItem.wdPrintDocumentContent;
        object copiesObj = Math.Max(1, copies);
        object pagesObj = string.IsNullOrWhiteSpace(pages) ? Missing.Value : (object)pages;
        object pageTypeObj = pageType;
        object printToFile = false;
        object collate = true;
        object activePrinterMacGX = Missing.Value;
        object manualDuplexPrint = false;
        object printZoomColumn = Missing.Value;
        object printZoomRow = Missing.Value;
        object printZoomPaperWidth = Missing.Value;
        object printZoomPaperHeight = Missing.Value;

        // Nếu người dùng chọn đảo ngược thứ tự in (hỗ trợ úp giấy in mặt sau tiện lợi)
        if (reverse && string.IsNullOrWhiteSpace(pages))
        {
            try
            {
                int totalPages = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
                var pageList = new List<int>();
                for (int i = 1; i <= totalPages; i++)
                {
                    if (pageType == WdPrintOutPages.wdPrintOddPagesOnly && i % 2 != 0)
                    {
                        pageList.Add(i);
                    }
                    else if (pageType == WdPrintOutPages.wdPrintEvenPagesOnly && i % 2 == 0)
                    {
                        pageList.Add(i);
                    }
                    else if (pageType == WdPrintOutPages.wdPrintAllPages)
                    {
                        pageList.Add(i);
                    }
                }

                pageList.Reverse();
                if (pageList.Count > 0)
                {
                    pagesObj = string.Join(",", pageList);
                    range = WdPrintOutRange.wdPrintRangeOfPages;
                    pageTypeObj = WdPrintOutPages.wdPrintAllPages;
                }
            }
            catch
            {
            }
        }

        doc.PrintOut(
            ref background, ref append, ref range, ref outputFileName,
            ref from, ref to, ref item, ref copiesObj, ref pagesObj,
            ref pageTypeObj, ref printToFile, ref collate, ref activePrinterMacGX,
            ref manualDuplexPrint, ref printZoomColumn, ref printZoomRow,
            ref printZoomPaperWidth, ref printZoomPaperHeight
        );

        try
        {
            if (!string.IsNullOrEmpty(originalPrinter) && doc.Application != null)
            {
                doc.Application.ActivePrinter = originalPrinter;
            }
        }
        catch
        {
        }
    }
}
