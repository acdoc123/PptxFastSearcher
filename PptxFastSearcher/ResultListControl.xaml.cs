using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using Button = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;
using Application = System.Windows.Application;

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
            // Lấy FilePath từ Tag của nút được bấm
            var button = sender as Button;
            if (button != null && button.Tag != null)
            {
                string filePath = button.Tag.ToString();
                try
                {
                    // Mở file bằng ứng dụng mặc định (PowerPoint)
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
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