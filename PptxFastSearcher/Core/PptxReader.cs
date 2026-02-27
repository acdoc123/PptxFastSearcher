using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxFastSearcher
{
    public class PptxReader
    {
        // Hàm này trả về danh sách các đoạn text, mỗi phần tử là nội dung của 1 Slide
        public static List<string> ExtractTextFromPptx(string filePath)
        {
            var slideTexts = new List<string>();

            try
            {
                // Mở file PPTX ở chế độ chỉ đọc (Read-only) để tăng tốc và tránh lỗi khóa file
                using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false))
                {
                    PresentationPart presentationPart = presentationDocument.PresentationPart;

                    if (presentationPart != null && presentationPart.SlideParts != null)
                    {
                        foreach (SlidePart slidePart in presentationPart.SlideParts)
                        {
                            // Tìm tất cả các Node chứa Text trong Slide hiện tại
                            var texts = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text);

                            // Ghép tất cả chữ trong 1 slide thành 1 chuỗi dài cách nhau bởi khoảng trắng
                            string fullSlideText = string.Join(" ", texts);

                            if (!string.IsNullOrWhiteSpace(fullSlideText))
                            {
                                slideTexts.Add(fullSlideText);
                            }
                            else
                            {
                                slideTexts.Add(string.Empty); // Slide trống
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Nếu file bị lỗi, đang được app khác mở, hoặc bị đặt mật khẩu -> Bỏ qua an toàn
                // Bạn có thể ghi log lỗi ở đây nếu cần
            }

            return slideTexts;
        }
    }
}