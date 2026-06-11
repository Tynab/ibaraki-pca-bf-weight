Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Điều phối toàn bộ luồng nhập liệu và ghi dữ liệu cho mẫu 茨城 PCa-BF.
''' </summary>
Friend Module Service
    ''' <summary>
    ''' Chạy lần lượt các nhóm câu hỏi và ghi kết quả vào sổ tính Excel đang mở.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel đang chứa sổ tính cần xử lý.</param>
    Friend Sub WtIbarakiPcaBF(xlApp As Application)
        ' Vận phí.
        Fare(xlApp, HdrYnq(vbTab & vbTab & "運賃 (2トン車): "))

        ' Sàn dạng móc.
        SlabHookType(xlApp, HdrYnq(vbTab & vbTab & "スラブフック型 (D13): "))

        ' Sàn dạng L.
        SlabLType(xlApp, HdrYnq(vbTab & vbTab & "スラブＬ型 (D13): "))

        ' Sàn thẳng.
        SlabStr(xlApp, HdrYnq(vbTab & vbTab & "スラブ直 (D13): "))

        ' Sàn gia cường dạng móc.
        SlabReinfHookType(xlApp, HdrYnq(vbTab & vbTab & "スラブ補強フック型 (D10): "))

        ' Sàn gia cường thẳng.
        SlabReinfStr(xlApp, HdrYnq(vbTab & vbTab & "スラブ補強直 (D10): "))

        ' Thép đầu dưới D13.
        HdrWrng(vbTab & vbTab & "下端 (D13)" & vbCrLf)
        LwrEndD13(xlApp)

        ' Thép đầu dưới D16.
        LwrEndD16(xlApp, HdrYnq(vbTab & vbTab & "下端 (D16): "))

        ' Thép đầu biên D10.
        Edge(xlApp, HdrYnq(vbTab & vbTab & "端部 (D10): "))

        ' Ống chờ.
        Sleeve(xlApp, HdrDInp(vbTab & vbTab & "スリーブ: "))

        ' Góc nối.
        HdrWrng(vbTab & vbTab & "コーナー" & vbCrLf)
        JtCor(xlApp)

        ' Thép chờ dùng cho doma.
        PubDVal(xlApp, "BA139", HdrDInp(vbTab & vbTab & "土間用さし: "))

        ' Thép dạng chữ U.
        PubDModVal(xlApp, "126", "（Ｕノ字型）", "900×80×900", 3.1, HdrDInp(vbTab & vbTab & "Ｕ型 (D16): "))

        ' Phần vát H250.
        Haunch(xlApp, HdrYnq(vbTab & vbTab & "ハンチ (H250): "))

        ' Sàn đầu biên cho móng sâu.
        PubDModVal(xlApp, "129", "650×250　　フック付", 0.6, HdrDInp(vbTab & vbTab & "深基礎用端部スラブ (D10): "))

        ' Máy nước nóng điện.
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))

        ' Danh sách phụ kiện/vật tư phụ.
        HdrWrng(vbTab & vbTab & "副資材リスト" & vbCrLf)
        Parts(xlApp)
    End Sub
End Module
