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
            int performanceMode = cmbPerformance.SelectedIndex;

            CurrentResults.Clear();
            btnSearch.IsEnabled = false;

            // Hiển thị nút Hủy và Thanh tiến trình
            btnCancel.Visibility = Visibility.Visible;
            btnCancel.IsEnabled = true;
            pbSearchProgress.Visibility = Visibility.Visible;
            txtStatus.Text = "Đang quét file...";

            // Bật thanh tiến trình Taskbar (Màu xanh lá)
            TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Normal;
            TaskbarItemInfo.ProgressValue = 0;

            // Khởi tạo CancellationToken mới cho phiên tìm kiếm này
            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;

            string cleanKeyword = SearchEngine.NormalizeSpaces(keyword);

            try
            {
                await Task.Run(() =>
                {
                    string[] pptxFiles = Directory.GetFiles(folderPath, "*.pptx", SearchOption.AllDirectories)
                              .OrderByDescending(f => File.GetLastWriteTime(f))
                              .ToArray();
                    int totalFiles = pptxFiles.Length;
                    int processedFiles = 0;

                    // TÍNH TOÁN SỐ LUỒNG (NHÂN CPU) DỰA TRÊN LỰA CHỌN
                    int maxThreads = 1; // Mặc định luôn là 1 nhân (An toàn tuyệt đối)

                    if (performanceMode == 1) // Chế độ Cao (50%)
                    {
                        maxThreads = Math.Max(1, Environment.ProcessorCount / 2);
                    }
                    else if (performanceMode == 2) // Chế độ Turbo (100%)
                    {
                        maxThreads = Environment.ProcessorCount;
                    }

                    // CẤU HÌNH ĐA LUỒNG
                    ParallelOptions options = new ParallelOptions
                    {
                        MaxDegreeOfParallelism = maxThreads,
                        CancellationToken = token
                    };

                    try
                    {
                        // SỬ DỤNG PARALLEL ĐỂ QUÉT SONG SONG CÁC FILE
                        Parallel.ForEach(pptxFiles, options, file =>
                        {
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
                                    DateTime fileTime = File.GetLastWriteTime(file);
                                    string rootFolderName = new DirectoryInfo(folderPath).Name;
                                    string relativePath = file.Substring(folderPath.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                                    string displayPath = Path.Combine(rootFolderName, relativePath);

                                    var newResult = new SearchResult
                                    {
                                        FileName = Path.GetFileName(file),
                                        FilePath = file,
                                        SlideNumber = $"Slide {i + 1}",
                                        DisplayPath = displayPath,
                                        MatchedText = cleanContent.Length > 200 ? cleanContent.Substring(0, 200) + "..." : cleanContent,
                                        LastWriteTime = fileTime // Gắn thời gian vào kết quả
                                    };

                                    // 2. Đưa kết quả lên UI một cách an toàn
                                    Application.Current.Dispatcher.InvokeAsync(() =>
                                    {
                                        // TÌM VỊ TRÍ CHÈN ĐỂ DUY TRÌ THỨ TỰ TỪ MỚI NHẤT ĐẾN CŨ NHẤT
                                        int insertPos = 0;
                                        // Dấu >= giúp các Slide trong CÙNG 1 file giữ đúng thứ tự 1, 2, 3...
                                        while (insertPos < CurrentResults.Count && CurrentResults[insertPos].LastWriteTime >= newResult.LastWriteTime)
                                        {
                                            insertPos++;
                                        }

                                        // Chèn vào đúng vị trí thay vì tống hết xuống cuối danh sách
                                        CurrentResults.Insert(insertPos, newResult);
                                    });
                                }
                            }

                            // Đếm số file đã quét an toàn trong môi trường đa luồng
                            Interlocked.Increment(ref processedFiles);

                            // Cập nhật UI và Taskbar mượt mà
                            Application.Current.Dispatcher.InvokeAsync(() =>
                            {
                                pbSearchProgress.Value = (processedFiles * 100) / totalFiles;
                                txtStatus.Text = $"Đã quét: {processedFiles}/{totalFiles} file";
                                TaskbarItemInfo.ProgressValue = (double)processedFiles / totalFiles;
                            });
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        // Bắt lỗi khi người dùng bấm Hủy để vòng lặp Parallel dừng lại êm đẹp
                    }
                }, token);

                // Nếu chạy xong mà không bị Hủy -> Lưu vào Lịch sử
                if (!token.IsCancellationRequested)
                {
                    txtStatus.Text = $"Hoàn tất! Tìm thấy {CurrentResults.Count} kết quả.";

                    // Tắt thanh Taskbar khi hoàn tất
                    TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;

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
                    TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.None;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra: {ex.Message}");

                // Đổi Taskbar thành màu ĐỎ báo lỗi
                TaskbarItemInfo.ProgressState = System.Windows.Shell.TaskbarItemProgressState.Error;
                TaskbarItemInfo.ProgressValue = 1.0;
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