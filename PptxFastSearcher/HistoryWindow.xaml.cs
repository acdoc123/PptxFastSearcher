using System.Collections.Generic;
using System.Windows;
using PptxFastSearcher.Models;

namespace PptxFastSearcher
{
    public partial class HistoryWindow : Window
    {
        public HistoryWindow(List<SearchHistorySession> historyData)
        {
            InitializeComponent();
            // Gắn danh sách lịch sử vào TabControl
            tabControlHistory.ItemsSource = historyData;
        }
    }
}