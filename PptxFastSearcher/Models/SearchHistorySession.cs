using System.Collections.ObjectModel;

namespace PptxFastSearcher.Models
{
    public class SearchHistorySession
    {
        public string Keyword { get; set; }
        public string ShortKeyword => Keyword.Length > 20 ? Keyword.Substring(0, 20) + "..." : Keyword;
        public ObservableCollection<SearchResult> Results { get; set; }
    }
}