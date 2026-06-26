# Stockchecker - build Windows EXE

## Cách build trên máy Windows

1. Cài Python 3.11 hoặc 3.12 và tick **Add Python to PATH**.
2. Giải nén thư mục này.
3. Mở file `build_windows_exe.bat`.
4. File chạy sẽ nằm ở:

```text
dist\Stockchecker\Stockchecker.exe
```

## Cách chạy

Double click:

```text
Stockchecker.exe
```

App sẽ chạy Streamlit local và mở trên trình duyệt.

## File dữ liệu MB52

Nếu dùng MB52 mặc định, tạo thư mục:

```text
dist\Stockchecker\data\MB52.xlsx
```

Hoặc upload MB52 trực tiếp trong giao diện app.

## Vì sao không có 1 file .exe duy nhất?

Với Streamlit + pandas, dạng thư mục `onedir` ổn định hơn `onefile`, ít lỗi thiếu file static/template hơn, và chạy nhanh hơn.
