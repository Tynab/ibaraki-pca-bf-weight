# TOUHOKU (PCA-BF) WEIGHT

Ứng dụng hỗ trợ đội 西山 thuộc エマールグループ nhập và chuyển dữ liệu nhanh hơn cho mẫu trọng lượng 茨城 (プレキャス - BF) từ đối tác 文化シャッター.

## Màn Hình

<p align='center'>
<img src='pic/0.png'></img>
</p>

## Thành Phần Chính

| Thành phần | Tệp | Mô tả |
|---|---|---|
| `Main` | `Script/Main.vb` | Điểm vào ứng dụng: thiết lập console UTF-8, kiểm tra serial bản quyền và gọi `RunApp`. |
| `Common` | `Script/Common.vb` | Hàm dùng chung: kiểm tra cập nhật, thao tác Excel COM, định dạng console, hiệu ứng biểu mẫu, quản lý tiến trình và file. |
| `Constant` | `Script/Constant.vb` | Hằng số toàn ứng dụng: tên tiến trình Excel, tên file cài đặt và đường dẫn trong AppData. |
| `Service` | `Script/Service.vb` | Điều phối luồng `WtIbarakiPcaBF`, chạy lần lượt toàn bộ nhóm câu hỏi nhập liệu. |
| `Util` | `Script/Util.vb` | Logic nghiệp vụ: mỗi nhóm nhập liệu được ánh xạ tới các ô Excel tương ứng. |
| `FrmUpdate` | `Control/FrmUpdate.vb` | Biểu mẫu cập nhật: tải file MSI, hiển thị tiến độ, chặn Alt+F4 và chạy trình cài đặt khi tải xong. |

## Luồng Chạy

```text
Main()
 ├─ PromptLicenseUntilValid()   ← kiểm tra serial trong tài nguyên
 └─ RunApp()
     ├─ ChkUpd()                ← kiểm tra phiên bản qua mạng; mở FrmUpdate nếu có bản mới
     ├─ KillXl()                ← kết thúc các tiến trình excel.exe đang mở
     ├─ OpenFileDialog          ← người dùng chọn file *.xlsx / *.xls
     └─ ProcessWorkbook()
         ├─ Excel.Application (COM)
         ├─ WtIbarakiPcaBF()    ← ghi toàn bộ nhóm dữ liệu vào sổ tính
         ├─ Workbook.Close(Save)
         └─ Process.Start()     ← mở lại file đã lưu bằng ứng dụng mặc định
```

## Nhóm Nhập Liệu

| # | Câu hỏi | Ô Excel |
|---|---|---|
| 1 | 運賃 (2トン車) | BA109, BA110, BA159 |
| 2 | スラブフック型 D13 | BA36 - BA45 |
| 3 | スラブＬ型 D13 | BA46 - BA55 |
| 4 | スラブ直 D13 | BA56 - BA67 |
| 5 | スラブ補強フック型 D10 | BA68 - BA77 |
| 6 | スラブ補強直 D10 | BA78 - BA87 |
| 7 | 下端 D13, bắt buộc | BA111 - BA121 |
| 8 | 下端 D16 | BA98 - BA107 |
| 9 | 端部 D10 | BA88 - BA97 |
| 10 | スリーブ | BA30, BA125 |
| 11 | コーナー | BA122 - BA124 |
| 12 | 土間用さし | BA139 |
| 13 | Ｕ型 D16 | BA126 |
| 14 | ハンチ H250 | BA127 - BA128 |
| 15 | 深基礎用端部スラブ D10 | BA129 |
| 16 | 電気温水器 | BA31 |
| 17 | 副資材リスト, gồm vật tư phụ và thông tin công trình | BA142 - BA158, AD6, BO3, BJ13, BJ14 |

## Minh Họa Mã Nguồn

Trích nguyên văn từ `Ibaraki Pca (BF) Weight/Script/Util.vb`:

```vb
    ''' <summary>
    ''' Ghi lựa chọn vận phí xe 2 tấn và số lượng mặc định cho D13/D10.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang xử lý.</param>
    ''' <param name="chosen">Giá trị chọn 1/0.</param>
    Friend Sub Fare(xlApp As Application, chosen As Integer)
        If chosen = 1 Then
            DctVal(xlApp, "BA159", chosen)
        End If
        DctVal(xlApp, "BA109", 5) ' D13
        DctVal(xlApp, "BA110", 3) ' D10
    End Sub
```

## Gói Phụ Thuộc

<img src='pic/1.png' align='left' width='3%' height='3%'></img>
<div style='display:flex;'>

- Microsoft.Office.Interop.Excel » 15.0.4795.1001

</div>
