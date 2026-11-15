using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using VFormat.Core.Pdf;

namespace VFormat.UI.Common
{
    public partial class PdfPageOrganizerDialog : Window
    {
        private class PageCardModel
        {
            public int OriginalIndex { get; set; } // 0-based
            public int RotationAngle { get; set; } // 0, 90, 180, 270
            public Border UIElement { get; set; }
        }

        private readonly string _pdfPath;
        private readonly int _totalPages;
        private readonly List<PageCardModel> _pages = new List<PageCardModel>();
        private int _selectedIndex = -1;

        public string ExportedPath { get; private set; } = string.Empty;

        public PdfPageOrganizerDialog(string pdfPath)
        {
            InitializeComponent();
            _pdfPath = pdfPath;
            txtFileName.Text = $" - {Path.GetFileName(pdfPath)}";

            _totalPages = PdfFlipEngine.GetPageCount(pdfPath);
            txtSummaryBadge.Text = $"{_totalPages} trang";

            // Khởi tạo danh sách trang ban đầu
            for (int i = 0; i < _totalPages; i++)
            {
                _pages.Add(new PageCardModel
                {
                    OriginalIndex = i,
                    RotationAngle = 0
                });
            }

            RenderGrid();
            SyncOrderText();
        }

        private void RenderGrid()
        {
            pnlPages.Children.Clear();

            var primaryBrush = (Brush)FindResource("BrushPrimary");
            var primaryLight = (Brush)FindResource("BrushPrimaryLight");
            var borderBrush = (Brush)FindResource("BrushBorderLight");

            for (int i = 0; i < _pages.Count; i++)
            {
                int pos = i;
                var page = _pages[i];

                bool isSelected = (pos == _selectedIndex);

                var card = new Border
                {
                    Width = 120,
                    Height = 160,
                    Margin = new Thickness(8),
                    CornerRadius = new CornerRadius(6),
                    Background = Brushes.White,
                    BorderThickness = new Thickness(isSelected ? 2.5 : 1),
                    BorderBrush = isSelected ? primaryBrush : borderBrush,
                    Cursor = Cursors.Hand
                };

                var grid = new Grid();
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(28) });
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });

                // Header badge: Thứ tự xuất hiện mới
                var topBadge = new Border
                {
                    Background = isSelected ? primaryBrush : primaryLight,
                    CornerRadius = new CornerRadius(5, 5, 0, 0),
                    Padding = new Thickness(4, 2, 4, 2)
                };
                var txtTop = new TextBlock
                {
                    Text = $"#{pos + 1}",
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    Foreground = isSelected ? Brushes.White : primaryBrush,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                topBadge.Child = txtTop;
                Grid.SetRow(topBadge, 0);
                grid.Children.Add(topBadge);

                // Body: Mô phỏng trang PDF
                var bodyBorder = new Border
                {
                    Margin = new Thickness(10, 6, 10, 4),
                    Background = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4)
                };

                var bodyGrid = new Grid();
                var pageIcon = new TextBlock
                {
                    Text = "📄",
                    FontSize = 28,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                bodyGrid.Children.Add(pageIcon);

                if (page.RotationAngle != 0)
                {
                    var rotBadge = new Border
                    {
                        Background = new SolidColorBrush(Color.FromArgb(200, 245, 158, 11)),
                        CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(3, 1, 3, 1),
                        HorizontalAlignment = HorizontalAlignment.Right,
                        VerticalAlignment = VerticalAlignment.Top,
                        Margin = new Thickness(2)
                    };
                    rotBadge.Child = new TextBlock
                    {
                        Text = $"⟳ {page.RotationAngle}°",
                        FontSize = 9,
                        FontWeight = FontWeights.Bold,
                        Foreground = Brushes.White
                    };
                    bodyGrid.Children.Add(rotBadge);
                }

                bodyBorder.Child = bodyGrid;
                Grid.SetRow(bodyBorder, 1);
                grid.Children.Add(bodyBorder);

                // Bottom badge: Trang gốc
                var txtBottom = new TextBlock
                {
                    Text = $"Gốc: Trang {page.OriginalIndex + 1}",
                    FontSize = 10,
                    Foreground = (Brush)FindResource("BrushTextSecondary"),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetRow(txtBottom, 2);
                grid.Children.Add(txtBottom);

                card.Child = grid;
                card.MouseLeftButtonDown += (s, e) =>
                {
                    _selectedIndex = pos;
                    RenderGrid();
                };

                page.UIElement = card;
                pnlPages.Children.Add(card);
            }

            txtSummaryBadge.Text = $"{_pages.Count} trang";
        }

        private void SyncOrderText()
        {
            var orderList = _pages.Select(p => (p.OriginalIndex + 1).ToString()).ToList();
            string fullStr = string.Join(", ", orderList);
            txtCurrentOrder.Text = fullStr;
            txtOrderSpec.Text = fullStr;
        }

        private void BtnMoveLeft_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedIndex > 0 && _selectedIndex < _pages.Count)
            {
                var item = _pages[_selectedIndex];
                _pages.RemoveAt(_selectedIndex);
                _selectedIndex--;
                _pages.Insert(_selectedIndex, item);
                RenderGrid();
                SyncOrderText();
            }
        }

        private void BtnMoveRight_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedIndex >= 0 && _selectedIndex < _pages.Count - 1)
            {
                var item = _pages[_selectedIndex];
                _pages.RemoveAt(_selectedIndex);
                _selectedIndex++;
                _pages.Insert(_selectedIndex, item);
                RenderGrid();
                SyncOrderText();
            }
        }

        private void BtnRotate_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedIndex >= 0 && _selectedIndex < _pages.Count)
            {
                _pages[_selectedIndex].RotationAngle = (_pages[_selectedIndex].RotationAngle + 90) % 360;
                RenderGrid();
            }
            else
            {
                ModernMessageDialog.ShowInfo("Vui lòng nhấn chọn một trang để xoay 90°.", "THÔNG BÁO");
            }
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedIndex >= 0 && _selectedIndex < _pages.Count)
            {
                _pages.RemoveAt(_selectedIndex);
                if (_selectedIndex >= _pages.Count)
                {
                    _selectedIndex = _pages.Count - 1;
                }
                RenderGrid();
                SyncOrderText();
            }
        }

        private void BtnReverseAll_Click(object sender, RoutedEventArgs e)
        {
            _pages.Reverse();
            if (_selectedIndex >= 0 && _selectedIndex < _pages.Count)
            {
                _selectedIndex = _pages.Count - 1 - _selectedIndex;
            }
            RenderGrid();
            SyncOrderText();
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            _pages.Clear();
            for (int i = 0; i < _totalPages; i++)
            {
                _pages.Add(new PageCardModel
                {
                    OriginalIndex = i,
                    RotationAngle = 0
                });
            }
            _selectedIndex = -1;
            RenderGrid();
            SyncOrderText();
        }

        private void BtnApplyOrderSpec_Click(object sender, RoutedEventArgs e)
        {
            string spec = txtOrderSpec.Text.Trim();
            if (string.IsNullOrEmpty(spec))
            {
                ModernMessageDialog.ShowWarning("Vui lòng nhập chuỗi thứ tự trang (VD: 1, 3, 2 hoặc 5-1).", "THÔNG BÁO");
                return;
            }

            try
            {
                var indices = PdfFlipEngine.ParsePageOrder(spec, _totalPages);
                var newPages = new List<PageCardModel>();

                foreach (var idx in indices)
                {
                    var existing = _pages.FirstOrDefault(p => p.OriginalIndex == idx);
                    newPages.Add(new PageCardModel
                    {
                        OriginalIndex = idx,
                        RotationAngle = existing?.RotationAngle ?? 0
                    });
                }

                _pages.Clear();
                _pages.AddRange(newPages);
                _selectedIndex = -1;
                RenderGrid();
                SyncOrderText();
            }
            catch (Exception ex)
            {
                ModernMessageDialog.ShowError($"Lỗi phân tích thứ tự trang:\n{ex.Message}", "LỖI ĐỊNH DẠNG");
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_pages.Count == 0)
            {
                ModernMessageDialog.ShowWarning("Không có trang nào để xuất.", "THÔNG BÁO");
                return;
            }

            string dir = Path.GetDirectoryName(_pdfPath) ?? "";
            string name = Path.GetFileNameWithoutExtension(_pdfPath);
            string defName = $"{name}_sap_xep.pdf";

            var sfd = new SaveFileDialog
            {
                Filter = "Tệp PDF (*.pdf)|*.pdf",
                FileName = defName,
                InitialDirectory = dir,
                Title = "Lưu tệp PDF đã sắp xếp trang"
            };

            if (sfd.ShowDialog() == true)
            {
                string savePath = sfd.FileName;
                var pageOrder = _pages.Select(p => p.OriginalIndex).ToList();
                var rotations = new Dictionary<int, int>();
                foreach (var p in _pages)
                {
                    if (p.RotationAngle != 0)
                    {
                        rotations[p.OriginalIndex] = p.RotationAngle;
                    }
                }

                var res = PdfFlipEngine.ProcessPdfNative(
                    _pdfPath,
                    outDir: Path.GetDirectoryName(savePath),
                    suffix: "",
                    inplace: false,
                    overwrite: true,
                    reverseOrder: false,
                    pageOrder: pageOrder,
                    manualRotations: rotations);

                if (res.Success)
                {
                    // Di chuyển đến file đúng nếu tên khác
                    if (!string.Equals(res.OutputPath, savePath, StringComparison.OrdinalIgnoreCase) && File.Exists(res.OutputPath))
                    {
                        try
                        {
                            if (File.Exists(savePath)) File.Delete(savePath);
                            File.Move(res.OutputPath, savePath);
                        }
                        catch { }
                    }

                    ExportedPath = savePath;
                    ModernMessageDialog.ShowSuccess($"Xuất tệp PDF thành công!\n\nĐường dẫn: {savePath}\nTổng số trang: {_pages.Count}", "HOÀN TẤT");
                    DialogResult = true;
                    Close();
                }
                else
                {
                    ModernMessageDialog.ShowError($"Lỗi khi xuất tệp PDF:\n{res.ErrorMessage}", "LỖI");
                }
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
