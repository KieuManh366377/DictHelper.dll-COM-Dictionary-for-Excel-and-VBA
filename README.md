# 📘 DictHelper.dll – COM Dictionary for Excel and VBA Automation

**DictHelper.dll** là một thư viện COM do **Kieu Manh** phát triển bằng **Embarcadero C++ Builder**, kế thừa trực tiếp từ lớp `TDictionary<TKey, TValue>` trong **System.Generics.Collections** của Delphi. Thư viện này cung cấp khả năng lưu trữ và thao tác dữ liệu dạng **key–value** một cách trực quan, mạnh mẽ và tương thích hoàn toàn với VBA, Excel, VBScript và các ứng dụng COM-based khác.

📧 Tác giả: Kieu Manh  
📮 Email: kieumanh366377@gmail.com

---

## ⚙️ Nền tảng kỹ thuật

DictHelper được xây dựng trên nền tảng **Delphi Runtime Library (RTL)**, tận dụng sức mạnh của `System.Generics.Collections` – một phần không thể thiếu của Delphi hiện đại. Việc kế thừa từ `TDictionary` giúp đảm bảo:

- Hiệu năng cao và quản lý bộ nhớ tối ưu  
- Tương thích tốt với kiểu dữ liệu `Variant`  
- Hỗ trợ đầy đủ các thao tác cơ bản và nâng cao của Dictionary  

💡 Xin trân trọng cảm ơn các kỹ sư Delphi đã xây dựng nên nền tảng Generics Collections – một công trình tuyệt vời giúp DictHelper trở nên mạnh mẽ và đáng tin cậy.

---

## ✨ Tính năng nổi bật

- Tương thích hoàn toàn với VBA và Excel  
- Hỗ trợ cú pháp `For Each` để duyệt dữ liệu  
- Lưu trữ dữ liệu động với kiểu `VARIANT`  
- Dễ dàng tích hợp vào macro, automation, hoặc ứng dụng COM  

---

## 🔧 Các phương thức và thuộc tính hỗ trợ

1. **Add(key, value)** – Thêm một cặp key–value vào dictionary. Nếu key đã tồn tại, giá trị sẽ được cập nhật.  
2. **GetItem(key)** – Truy xuất giá trị tương ứng với key đã cho.  
3. **Remove(key)** – Xóa một key và giá trị tương ứng khỏi dictionary.  
4. **Exists(key)** – Kiểm tra xem key có tồn tại trong dictionary hay không.  
5. **Count** – Trả về tổng số phần tử hiện có trong dictionary.  
6. **Item(key)** – Truy xuất hoặc gán giá trị trực tiếp bằng cú pháp `dict(key) = value`.  
7.  – Hỗ trợ `For Each` trong VBA để duyệt qua tất cả các key.  

---

## 🧪 Ví dụ sử dụng trong VBA

```vb
Sub DemoDictHelper()
    Dim dict As Object
    Set dict = CreateObject("DictHelper.Dictionary")

    dict.Add "Name", "Kieu"
    dict.Add "City", "Lai Thieu"

    If dict.Exists("Name") Then
        MsgBox "Tên: " & dict.GetItem("Name")
    End If

    Debug.Print "Tổng số phần tử: " & dict.Count

    Dim key As Variant
    For Each key In dict
        Debug.Print key & " = " & dict(key)
    Next
End Sub
```
## 🔄 So sánh DictHelper.dll với Scripting.Dictionary

**DictHelper.dll** được thiết kế mô phỏng theo đối tượng `Scripting.Dictionary` của Microsoft, nhằm mang lại trải nghiệm tương tự trong môi trường VBA, Excel và VBScript. Người dùng có thể sử dụng DictHelper với cùng cú pháp, phương thức và thuộc tính như Scripting.Dictionary.

### ✅ Điểm tương đồng

| Tính năng | Scripting.Dictionary | DictHelper.Dictionary |
|----------|----------------------|------------------------|
| Thêm phần tử | `Add key, value` | `Add key, value` |
| Truy xuất giá trị | `dict(key)` hoặc `Item(key)` | `dict(key)` hoặc `Item(key)` |
| Kiểm tra key | `Exists(key)` | `Exists(key)` |
| Xóa phần tử | `Remove(key)` | `Remove(key)` |
| Đếm phần tử | `Count` | `Count` |
| Duyệt key | `For Each key In dict` | `For Each key In dict` |

➡️ **Cú pháp giống nhau 100%**, giúp người dùng chuyển đổi dễ dàng mà không cần viết lại code.

---

### 📌 Điểm khác biệt tiềm năng

- **DictHelper** được viết bằng C++ Builder và kế thừa từ `System.Generics.Collections.TDictionary` của Delphi, nên có thể xử lý kiểu dữ liệu `Variant` tốt hơn trong môi trường COM.

---

### 🧪 Ví dụ chuyển đổi

**Từ Scripting.Dictionary:**

```vb
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "Name", "Kieu"
MsgBox dict("Name")
```

**Sang DictHelper.Dictionary:**

```vb
Set dict = CreateObject("DictHelper.Dictionary")
dict.Add "Name", "Kieu"
MsgBox dict("Name")
```

➡️ Không cần thay đổi cú pháp, chỉ thay đổi tên COM ProgID.

---

---

## 📦 Cài đặt và đăng ký

1. Copy file `DictHelper.dll` vào thư mục hệ thống hoặc thư mục dự án  
2. Đăng ký DLL bằng lệnh sau trong Command Prompt:
   ```
   regsvr32 DictHelper.dll
   ```
3. Sử dụng trong VBA bằng cách gọi:
   ```
   Set dict = CreateObject("DictHelper.Dictionary")
   ```

---

## 🎯 Ứng dụng thực tế

- Lọc dữ liệu trùng lặp trong Excel  
- Thống kê tần suất xuất hiện  
- Tạo bảng ánh xạ key–value động  
- Tích hợp vào quy trình automation hoặc báo cáo  

---

DictHelper không chỉ là một thư viện tiện ích – nó là cầu nối giữa sức mạnh của Delphi và sự linh hoạt của VBA. Nếu bạn là lập trình viên Excel, người dùng VBScript, hoặc đang tìm giải pháp lưu trữ dữ liệu linh hoạt trong COM, **DictHelper.dll** là một lựa chọn đáng tin cậy và dễ triển khai.
