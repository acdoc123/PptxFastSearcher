using System;

namespace PptxFastSearcher.Models
{
    public class SearchResult
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string SlideNumber { get; set; } // Ví dụ: "Slide 3"
        public string MatchedText { get; set; } // Đoạn văn bản chứa từ khóa
        public DateTime LastWriteTime { get; set; }
    }
}