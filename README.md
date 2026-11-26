# go-app

Dọn tàn dư của ứng dụng Windows đã gỡ dở dang: vẫn hiện trong "Installed apps", vẫn tự chạy,
hoặc làm hỏng liên kết file và logo. Đúc kết từ lần dọn WPS Office, TeraBox, Riff (11/2026).

## Cách dùng

1. Gỡ ứng dụng trong **Settings → Apps** trước. Nếu gỡ được thì thường không cần script này.
2. Mở thư mục này trong File Explorer, chạy trong PowerShell của bạn:

   ```powershell
   .\go-app.cmd -Pattern "tenapp" -DryRun
   ```

3. **Đọc kỹ danh sách.** Đúng hết thì chạy lại, bỏ `-DryRun`, rồi gõ `y`.

Bấm đúp `go-app.cmd` cũng được, script sẽ hỏi tên app.

`-Pattern` là regex, không phân biệt hoa thường:

| Trường hợp | Mẫu |
|---|---|
| Một tên | `terabox` |
| Nhiều tên của cùng một app | `"terabox\|dubox"` |
| WPS Office | `"kingsoft\|\\wps office\\"` |
| Tên ngắn, dễ trùng (RIFF còn là một định dạng file) | `"\\Programs\\riff\\"` |

Tên nhà phát hành và tên thư mục cài đặt thường là mẫu tốt nhất (xem trong danh sách `-DryRun`).

## Script làm gì

- `HKCU\Software\Classes`: xóa kiểu file riêng của app; với khóa dùng chung (vd `Word.Document.12`)
  chỉ xóa giá trị app ghi đè; gỡ các đuôi file còn trỏ tới app.
- Mục gỡ cài đặt, mục tự chạy, `HKCU\Software\<app>`, kết nối tiện ích trình duyệt, ô Start.
- Thư mục của app trong `AppData` → Thùng rác.
- Sao lưu registry vào `backups\` trước khi xóa; khởi động lại Explorer và tạo lại bộ đệm icon.

## Không làm

- Không sửa `HKLM` (chỉ liệt kê), không đụng Program Files, Documents, Downloads.
- Không xóa `UserChoice`: Windows chặn (driver UCPD). Chọn lại ứng dụng mặc định trong Settings.
- Từ chối chạy khi app còn trình gỡ cài đặt, hoặc mẫu quá rộng.

## Bài học

- **Đừng chạy từ terminal của app đóng gói (vd Claude Desktop).** Nó thấy bản ảo của `HKCU` và
  AppData: lệnh báo xóa thành công nhưng máy thật không đổi. Script tự kiểm tra việc này qua WMI.
- File đã mất nhưng logo trắng: tìm `DefaultIcon` trong `HKCU\Software\Classes` còn trỏ tới app cũ.
- Tìm theo tên nhà phát hành thôi có thể sót: WPS dùng cả `Kingsoft`, `KWPS.*`, `ET.*`, `WPP.*`.

## Khôi phục

Bấm đúp file `.reg` trong `backups\<thời gian>\`. File thư mục lấy lại từ Thùng rác.

## Kiểm tra script

```powershell
powershell -NoProfile -File .\test.ps1
```
