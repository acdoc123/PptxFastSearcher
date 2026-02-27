using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace PptxFastSearcher.Core
{
    public static class SearchEngine
    {
        // Hàm này MẶC ĐỊNH luôn chạy để dọn dẹp khoảng cách dư
        public static string NormalizeSpaces(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            // Gom nhiều dấu cách thành 1 và xóa khoảng trắng ở 2 đầu
            return Regex.Replace(input, @"\s+", " ").Trim();
        }

        // Hàm xóa dấu câu (.,!?;"...)
        public static string RemovePunctuation(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            // \p{P} là Regex đại diện cho tất cả các loại dấu câu
            return Regex.Replace(input, @"\p{P}", "");
        }

        // Hàm xóa dấu tiếng Việt
        public static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }
            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
    }
}