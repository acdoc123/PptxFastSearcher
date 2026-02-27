using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxFastSearcher.Core // Nhớ kiểm tra lại xem có khớp namespace của bạn không nhé
{
    public class PptxReader
    {
        // Hàm này trả về danh sách các đoạn text, mỗi phần tử là nội dung của 1 Slide (THEO ĐÚNG THỨ TỰ)
        public static List<string> ExtractTextFromPptx(string filePath)
        {
            var slideTexts = new List<string>();

            try
            {
                // Mở file PPTX ở chế độ chỉ đọc
                using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false))
                {
                    PresentationPart presentationPart = presentationDocument.PresentationPart;

                    // Kiểm tra xem presentationPart và danh sách SlideId có tồn tại không
                    if (presentationPart != null && presentationPart.Presentation != null && presentationPart.Presentation.SlideIdList != null)
                    {
                        // Lấy danh sách các SlideId THEO ĐÚNG THỨ TỰ TRÌNH BÀY (Visual Order)
                        var slideIds = presentationPart.Presentation.SlideIdList.Elements<SlideId>();

                        foreach (SlideId slideId in slideIds)
                        {
                            // Lấy SlidePart dựa trên RelationshipId của SlideId hiện tại
                            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);

                            if (slidePart != null && slidePart.Slide != null)
                            {
                                // Tìm tất cả các Node chứa Text trong Slide
                                var texts = slidePart.Slide.Descendants<A.Text>().Select(t => t.Text);

                                string fullSlideText = string.Join(" ", texts);

                                if (!string.IsNullOrWhiteSpace(fullSlideText))
                                {
                                    slideTexts.Add(fullSlideText);
                                }
                                else
                                {
                                    slideTexts.Add(string.Empty); // Giữ chỗ cho Slide trống để không bị lệch Index
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Bỏ qua nếu file bị lỗi hoặc đang bị khóa
            }

            return slideTexts;
        }
    }
}