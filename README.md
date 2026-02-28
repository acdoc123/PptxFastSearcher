🚀 **Tóm tắt chức năng chính:**
Tìm kiếm siêu tốc: Quét hàng ngàn file .pptx trong nháy mắt mà không cần mở ứng dụng PowerPoint chạy ngầm.

Bộ lọc thông minh (3 chế độ): Tự động dọn dẹp khoảng trắng dư thừa. Hỗ trợ tìm chính xác, tìm gần đúng (bỏ qua viết hoa, dấu tiếng Việt, dấu câu) và tìm mở rộng (các từ nằm rải rác).

Ưu tiên file mới: Tự động đẩy các file mới được chỉnh sửa gần đây nhất lên quét trước để bạn thấy kết quả ngay lập tức.

Chỉ đường tận nơi: Bấm nút "Mở file", ứng dụng sẽ gọi PowerPoint lên và cuộn thẳng đến đúng Slide chứa từ khóa.

Tiện ích đi kèm: Có nút "Hủy" để dừng quét giữa chừng và tính năng lưu trữ Lịch sử tìm kiếm cực kỳ trực quan bằng các Tab.

⚠️** Một vài lưu ý nhỏ:**
Quá trình tìm kiếm hoàn toàn độc lập, nhưng để tính năng nhảy đến đúng Slide hoạt động hoàn hảo, máy tính chạy app bắt buộc phải cài đặt sẵn phần mềm Microsoft PowerPoint. (Nếu máy không có, nó vẫn mở file lên bình thường nhưng sẽ nằm ở Slide đầu tiên).
Dung lượng file .exe có thể hơi lớn nếu bạn chọn đóng gói bao gồm cả môi trường .NET, nhưng bù lại máy nào cắm vào cũng chạy được ngay.


**Hướng dẫn chạy code:**
**Bước 1:** Tải dự án file zip hoặc clone về....
**Bước 2:** Mở dự án
- Mở thư mục vừa tải về, tìm và nhấp đúp vào file solution: PptxFastSearcher.sln.
- Visual Studio 2022 sẽ tự động khởi động và tải dự án.

**Bước 3: Khôi phục thư viện (Restore NuGet Packages)**
- Dự án sử dụng một số thư viện bên ngoài (OpenXML, MaterialDesignThemes, FuzzySharp).
- Thông thường, Visual Studio sẽ tự động tải các thư viện này khi bạn mở project.
- Nếu nó không tự tải, bạn hãy chuột phải vào thẻ Solution ở bảng Solution Explorer (góc phải màn hình) -> Chọn Restore NuGet Packages.

**Bước 4: Build và Chạy**
- Nhấn phím F5 (hoặc bấm nút Start hình tam giác màu xanh lá cây ở thanh menu trên cùng).
- Visual Studio sẽ tiến hành biên dịch (build) và ứng dụng sẽ hiển thị lên ngay lập tức! 🎉

📦 **Cách đóng gói ứng dụng (Publish)**
Nếu bạn muốn tự tay đóng gói mã nguồn thành một file .exe duy nhất để gửi cho người khác dùng luôn mà không cần cài Visual Studio:
- Chuột phải vào project PptxFastSearcher -> Chọn Publish...
- Chọn Folder làm đích đến.
- Ở phần Show all settings, hãy cấu hình:
- Target framework: net8.0-windows
- Deployment mode: Self-contained (Tích hợp sẵn .NET để máy khác không cần cài thêm).
- Target runtime: win-x64/86
- File publish options: Tích chọn Produce single file.

- Bấm Publish và lấy file .exe trong thư mục đích.
