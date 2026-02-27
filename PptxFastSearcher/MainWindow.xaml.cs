using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using PptxFastSearcher.Core;
using PptxFastSearcher.Models;

using MessageBox = System.Windows.MessageBox;
using Application = System.Windows.Application;
using Button = System.Windows.Controls.Button;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace PptxFastSearcher
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<SearchResult> CurrentResults { get; set; }

        // Danh sách lưu lịch sử (Tối đa 6)
        public List<SearchHistorySession> SearchHistory { get; set; } = new List<SearchHistorySession>();

        // Công cụ để Hủy Task đang chạy
        private CancellationTokenSource _cancellationTokenSource;

        public MainWindow()
        {
            InitializeComponent();
            CurrentResults = new ObservableCollection<SearchResult>();
            // Cấp dữ liệu cho Component tái sử dụng
            mainResultControl.DataContext = CurrentResults;
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Chọn thư mục chứa các file PPTX";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    txtFolderPath.Text = dialog.SelectedPath;
                }
            }
        }

        // BẮT SỰ KIỆN PHÍM ENTER
        private void txtKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // Chỉ cho phép Enter nếu nút Tìm Kiếm đang sáng (nghĩa là không có tiến trình nào đang chạy)
                if (btnSearch.IsEnabled)
                {
                    btnSearch_Click(sender, e);
                }
            }
        }

        // NÚT HỦY TÌM KIẾM
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            if (_cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel(); // Phát tín hiệu Hủy
                btnCancel.IsEnabled = false;
                txtStatus.Text = "Đang hủy...";
            }
        }

        private async void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = txtFolderPath.Text;
            string keyword = txtKeyword.Text;

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return;
            if (string.IsNullOrWhiteSpace(keyword))
                return;

            bool isExact = rbExact.IsChecked == true;
            bool isFuzzy = rbFuzzy.IsChecked == true;
            bool isLoose = rbLoose.IsChecked == true;

            CurrentResults.Clear();
            btnSearch.IsEnabled = false;

            // Hiển thị nút Hủy và Thanh tiến trình
            btnCancel.Visibility = Visibility.Visible;
            btnCancel.IsEnabled = true;
            pbSearchProgress.Visibility = Visibility.Visible;
            txtStatus.Text = "Đang quét file...";

            // Khởi tạo CancellationToken mới cho phiên tìm kiếm này
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            string cleanKeyword = SearchEngine.NormalizeSpaces(keyword);

            try
            {
                await Task.Run(() =>
                {
                    string[] pptxFiles = Directory.GetFiles(folderPath, "*.pptx", SearchOption.AllDirectories);
                    int totalFiles = pptxFiles.Length;
                    int processedFiles = 0;

                    foreach (string file in pptxFiles)
                    {
                        // KIỂM TRA XEM NGƯỜI DÙNG CÓ BẤM HỦY KHÔNG
                        if (token.IsCancellationRequested)
                        {
                            break; // Thoát khỏi vòng lặp quét file ngay lập tức
                        }

                        var slidesText = PptxReader.ExtractTextFromPptx(file);

                        for (int i = 0; i < slidesText.Count; i++)
                        {
                            string slideContent = slidesText[i];
                            if (string.IsNullOrWhiteSpace(slideContent)) continue;

                            string cleanContent = SearchEngine.NormalizeSpaces(slideContent);
                            bool isMatch = false;

                            if (isExact)
                            {
                                isMatch = cleanContent.Contains(cleanKeyword);
                            }
                            else
                            {
                                string processedKey = SearchEngine.RemovePunctuation(SearchEngine.RemoveDiacritics(cleanKeyword)).ToLowerInvariant();
                                string processedContent = SearchEngine.RemovePunctuation(SearchEngine.RemoveDiacritics(cleanContent)).ToLowerInvariant();

                                if (isFuzzy) isMatch = processedContent.Contains(processedKey);
                                else if (isLoose)
                                {
                                    string[] words = processedKey.Split(' ');
                                    isMatch = true;
                                    foreach (string word in words)
                                        if (!processedContent.Contains(word)) { isMatch = false; break; }
                                }
                            }

                            if (isMatch)
                            {
                                Application.Current.Dispatcher.Invoke(() =>
                                {
                                    CurrentResults.Add(new SearchResult
                                    {
                                        FileName = Path.GetFileName(file),
                                        FilePath = file,
                                        SlideNumber = $"Slide {i + 1}",
                                        MatchedText = cleanContent.Length > 200 ? cleanContent.Substring(0, 200) + "..." : cleanContent
                                    });
                                });
                            }
                        }

                        processedFiles++;
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            pbSearchProgress.Value = (processedFiles * 100) / totalFiles;
                            txtStatus.Text = $"Đã quét: {processedFiles}/{totalFiles} file";
                        });
                    }
                }, token);

                // Nếu chạy xong mà không bị Hủy -> Lưu vào Lịch sử
                if (!token.IsCancellationRequested)
                {
                    txtStatus.Text = $"Hoàn tất! Tìm thấy {CurrentResults.Count} kết quả.";

                    // LƯU LỊCH SỬ
                    SearchHistory.Insert(0, new SearchHistorySession
                    {
                        Keyword = keyword,
                        // Cần tạo một bản sao của list kết quả để lưu trữ, tránh bị Clear ở lần tìm sau
                        Results = new ObservableCollection<SearchResult>(CurrentResults.ToList())
                    });

                    // Xóa bớt nếu vượt quá 6 cái gần nhất
                    if (SearchHistory.Count > 6) SearchHistory.RemoveAt(6);
                    btnHistory.Content = $"Lịch sử ({SearchHistory.Count})";
                }
                else
                {
                    txtStatus.Text = "Đã hủy quá trình tìm kiếm.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra: {ex.Message}");
            }
            finally
            {
                // Khôi phục UI
                btnSearch.IsEnabled = true;
                pbSearchProgress.Visibility = Visibility.Collapsed;
                btnCancel.Visibility = Visibility.Collapsed;
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }

        // MỞ CỬA SỔ LỊCH SỬ
        private void btnHistory_Click(object sender, RoutedEventArgs e)
        {
            if (SearchHistory.Count == 0)
            {
                MessageBox.Show("Chưa có lịch sử tìm kiếm nào được lưu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Mở cửa sổ mới và truyền danh sách lịch sử sang
            HistoryWindow historyWin = new HistoryWindow(SearchHistory);
            historyWin.ShowDialog();
        }
    }
}