using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using VFormat.Core.Common;
using VFormat.WordEngine;

namespace VFormat.UI.Common;

public class BankInfo
{
    public string Name { get; set; }
    public string Bin { get; set; }
    public string Code { get; set; }

    public override string ToString() => $"{Name} ({Code})";
}

public partial class QrCodeDialog : System.Windows.Window
{
    private readonly Document _doc;
    private string _currentQrPayload = "";

    private readonly List<BankInfo> _banks = new List<BankInfo>
    {
        new BankInfo { Name = "Vietcombank (VCB)", Bin = "970436", Code = "VCB" },
        new BankInfo { Name = "MB Bank (Quân đội)", Bin = "970422", Code = "MB" },
        new BankInfo { Name = "Techcombank", Bin = "970407", Code = "TCB" },
        new BankInfo { Name = "BIDV (Đầu tư & Phát triển)", Bin = "970418", Code = "BIDV" },
        new BankInfo { Name = "VietinBank (Công thương)", Bin = "970415", Code = "CTG" },
        new BankInfo { Name = "Agribank (Nông nghiệp)", Bin = "970405", Code = "VBA" },
        new BankInfo { Name = "VPBank", Bin = "970432", Code = "VPB" },
        new BankInfo { Name = "TPBank (Tiên Phong)", Bin = "970423", Code = "TPB" },
        new BankInfo { Name = "ACB (Á Châu)", Bin = "970416", Code = "ACB" },
        new BankInfo { Name = "Sacombank", Bin = "970403", Code = "STB" },
        new BankInfo { Name = "HDBank", Bin = "970437", Code = "HDB" },
        new BankInfo { Name = "SHB (Sài Gòn - Hà Nội)", Bin = "970443", Code = "SHB" },
        new BankInfo { Name = "VIB (Quốc tế)", Bin = "970441", Code = "VIB" },
        new BankInfo { Name = "MSB (Hàng Hải)", Bin = "970426", Code = "MSB" },
        new BankInfo { Name = "OCB (Phương Đông)", Bin = "970448", Code = "OCB" },
        new BankInfo { Name = "SeABank (Đông Nam Á)", Bin = "970440", Code = "SEAB" },
        new BankInfo { Name = "LPBank (Lộc Phát)", Bin = "970449", Code = "LPB" }
    };

    public int SelectedTabIndex
    {
        get => tabQrType.SelectedIndex;
        set => tabQrType.SelectedIndex = value;
    }

    public QrCodeDialog(Document doc, string selectedText = null)
    {
        InitializeComponent();
        _doc = doc;

        cboBank.ItemsSource = _banks;
        cboBank.SelectedIndex = 0;

        if (!string.IsNullOrEmpty(selectedText))
        {
            txtCustomText.Text = selectedText;
            tabQrType.SelectedIndex = 1; // Chọn tab Custom
        }
        else
        {
            // Tự động quét thông tin văn bản
            if (_doc != null)
            {
                var meta = QrCodeEngine.ExtractMetadata(_doc);
                txtDocAgency.Text = meta.Agency;
                txtDocNumber.Text = meta.DocumentNumber;
                txtDocDate.Text = meta.DateText;
                txtDocExcerpt.Text = meta.Excerpt;
                txtDocSigner.Text = meta.Signer;
                txtDocUrl.Text = meta.LookupUrl;
            }
        }

        UpdateQrPreview();
    }

    private void TabQrType_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateQrPreview();
    }

    private void DocInputs_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (tabQrType.SelectedIndex == 0)
        {
            UpdateQrPreview();
        }
    }

    private void TxtCustomText_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (tabQrType.SelectedIndex == 1)
        {
            UpdateQrPreview();
        }
    }

    private void VietQrInputs_Changed(object sender, EventArgs e)
    {
        if (tabQrType.SelectedIndex == 2)
        {
            UpdateQrPreview();
        }
    }

    private void UpdateQrPreview()
    {
        string payload = "";

        switch (tabQrType.SelectedIndex)
        {
            case 0: // Văn bản
                payload = QrCodeHelper.BuildDocumentMetadataString(
                    txtDocAgency.Text, txtDocNumber.Text, txtDocDate.Text,
                    txtDocExcerpt.Text, txtDocSigner.Text, txtDocUrl.Text
                );
                break;

            case 1: // Custom / Selection
                payload = (txtCustomText.Text ?? "").Trim();
                break;

            case 2: // VietQR
                var selectedBank = cboBank.SelectedItem as BankInfo;
                string bankBin = selectedBank != null ? selectedBank.Bin : "970436";
                decimal? amount = null;
                if (decimal.TryParse((txtAmount.Text ?? "").Replace(".", "").Replace(",", ""), out decimal a) && a > 0)
                {
                    amount = a;
                }
                payload = QrCodeHelper.BuildVietQrEmvCoPayload(
                    bankBin, txtAccountNo.Text, amount, txtMemo.Text, txtAccountName.Text
                );
                break;
        }

        if (string.IsNullOrWhiteSpace(payload))
        {
            payload = "https://dichvucong.gov.vn";
        }

        _currentQrPayload = payload;

        try
        {
            byte[] pngBytes = QrCodeHelper.GenerateQrPngBytes(payload, 8);
            using (var ms = new MemoryStream(pngBytes))
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.StreamSource = ms;
                bitmap.EndInit();
                bitmap.Freeze();
                imgQrPreview.Source = bitmap;
            }
        }
        catch
        {
        }
    }

    private void BtnInsert_Click(object sender, RoutedEventArgs e)
    {
        if (_doc == null)
        {
            System.Windows.MessageBox.Show("Không tìm thấy tài liệu Word đang mở.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(_currentQrPayload))
        {
            System.Windows.MessageBox.Show("Nội dung mã QR đang trống.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var loc = QrInsertLocation.Cursor;
        if (rbInsertTopRight.IsChecked == true) loc = QrInsertLocation.HeaderTopRight;
        else if (rbInsertEnd.IsChecked == true) loc = QrInsertLocation.DocumentEnd;

        float sizeCm = 2.5f;
        if (rbSizeSmall.IsChecked == true) sizeCm = 2.0f;
        else if (rbSizeLarge.IsChecked == true) sizeCm = 3.5f;

        var mode = rbFormatNative.IsChecked == true ? QrFormatMode.WordNativeField : QrFormatMode.ImageFile;

        try
        {
            QrCodeEngine.InsertQrCode(_doc, _currentQrPayload, loc, mode, sizeCm);
            System.Windows.MessageBox.Show("Đã chèn mã QR vào tài liệu Word thành công!", "Tạo mã QR", MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Lỗi khi chèn mã QR: {ex.Message}", "Lỗi tạo QR", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnSaveImage_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentQrPayload))
        {
            System.Windows.MessageBox.Show("Nội dung mã QR đang trống.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var sfd = new SaveFileDialog
        {
            Filter = "Ảnh PNG (*.png)|*.png",
            FileName = "QRCode.png",
            Title = "Lưu mã QR thành tệp ảnh"
        };

        if (sfd.ShowDialog() == true)
        {
            try
            {
                QrCodeHelper.SaveQrPng(_currentQrPayload, sfd.FileName, 12);
                System.Windows.MessageBox.Show("Đã lưu ảnh mã QR thành công!", "Lưu ảnh", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Lỗi khi lưu ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
