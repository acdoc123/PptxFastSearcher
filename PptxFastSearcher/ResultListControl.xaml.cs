using System;
using System.Diagnostics;
using System.Windows;
using PptxFastSearcher.Models;

using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace PptxFastSearcher
{
    public partial class ResultListControl : UserControl
    {
        public ResultListControl()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;

            // Lấy toàn bộ đối tượng SearchResult từ Tag
            if (button != null && button.Tag is SearchResult result)
            {
                try
                {
                    // Tách lấy con số từ chuỗi "Slide 3"
                    int slideIndex = int.Parse(result.SlideNumber.Replace("Slide ", "").Trim());

                    // --- SỬ DỤNG DYNAMIC LATE-BINDING TỐI ƯU ---
                    // Tự động mò tìm PowerPoint trên máy mà không cần quan tâm phiên bản
                    Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                    if (pptType == null)
                    {
                        throw new Exception("Không tìm thấy ứng dụng PowerPoint trên máy tính này.");
                    }

                    // Khởi tạo PowerPoint ẩn và gọi lệnh hiển thị
                    dynamic pptApp = Activator.CreateInstance(pptType);
                    pptApp.Visible = true;

                    // Mở file PPTX
                    dynamic presentation = pptApp.Presentations.Open(result.FilePath);

                    // Lệnh nhảy đến đúng Slide và bôi đen nó
                    presentation.Slides[slideIndex].Select();
                }
                catch (Exception)
                {
                    // DỰ PHÒNG: Nếu code động ở trên lỗi (do máy người dùng bị lỗi Office),
                    // tự động chuyển về cách mở file cơ bản nhất.
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = result.FilePath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Không thể mở file. Lỗi: {ex.Message}");
                    }
                }
            }
        }
    }
}