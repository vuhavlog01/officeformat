using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Interop.Word;
using VFormat.WordEngine;

namespace VFormat.UI.Common;

public partial class PrintDuplexDialog : System.Windows.Window
{
    private readonly Document _doc;

    public PrintDuplexDialog(Document doc)
    {
        InitializeComponent();
        _doc = doc;
        Loaded += PrintDuplexDialog_Loaded;
    }

    private void PrintDuplexDialog_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            var printers = PrintEngine.GetInstalledPrinters(out string defaultPrinter);
            cboPrinters.ItemsSource = printers;
            if (!string.IsNullOrEmpty(defaultPrinter) && printers.Contains(defaultPrinter))
            {
                cboPrinters.SelectedItem = defaultPrinter;
            }
            else if (printers.Count > 0)
            {
                cboPrinters.SelectedIndex = 0;
            }
        }
        catch
        {
        }
    }

    private void BtnOddFast_Click(object sender, RoutedEventArgs e)
    {
        DoPrint(WdPrintOutPages.wdPrintOddPagesOnly, "", false);
    }

    private void BtnEvenFast_Click(object sender, RoutedEventArgs e)
    {
        DoPrint(WdPrintOutPages.wdPrintEvenPagesOnly, "", false);
    }

    private void BtnPrint_Click(object sender, RoutedEventArgs e)
    {
        WdPrintOutPages pageType = WdPrintOutPages.wdPrintAllPages;
        string pages = "";

        if (rbOdd.IsChecked == true)
        {
            pageType = WdPrintOutPages.wdPrintOddPagesOnly;
        }
        else if (rbEven.IsChecked == true)
        {
            pageType = WdPrintOutPages.wdPrintEvenPagesOnly;
        }
        else if (rbCustom.IsChecked == true)
        {
            pages = (txtPages.Text ?? "").Trim();
            if (string.IsNullOrEmpty(pages))
            {
                System.Windows.MessageBox.Show("Vui lòng nhập dải trang cần in (ví dụ: 1-10 hoặc 1,3,5).", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        bool reverse = chkReverse.IsChecked == true;
        DoPrint(pageType, pages, reverse);
    }

    private void DoPrint(WdPrintOutPages pageType, string pages, bool reverse)
    {
        if (_doc == null)
        {
            System.Windows.MessageBox.Show("Không tìm thấy tài liệu Word đang mở.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!int.TryParse(txtCopies.Text, out int copies) || copies < 1)
        {
            copies = 1;
        }

        string printer = cboPrinters.SelectedItem as string;

        try
        {
            PrintEngine.PrintPages(_doc, pageType, pages, copies, reverse, printer);
            string modeName = pageType == WdPrintOutPages.wdPrintOddPagesOnly ? "trang lẻ" : (pageType == WdPrintOutPages.wdPrintEvenPagesOnly ? "trang chẵn" : "các trang đã chọn");
            System.Windows.MessageBox.Show($"Đã gửi lệnh in {modeName} đến máy in thành công!", "In ấn hoàn tất", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Lỗi trong quá trình in: {ex.Message}", "Lỗi in ấn", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
