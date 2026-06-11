Imports System.ComponentModel
Imports System.Console
Imports System.ConsoleColor
Imports System.IO
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Cung cấp các hàm dùng chung cho update, nhập liệu console, thao tác Excel và form.
''' </summary>
Friend Module Common
    Private Const UpdateCheckTimeoutMs As Integer = 5000
    Private ReadOnly ModifiedCellColor As Integer = RGB(0, 176, 240)

#Region "Network"
    ''' <summary>
    ''' WebClient có thời gian chờ để tránh treo ứng dụng khi địa chỉ cập nhật phản hồi chậm.
    ''' </summary>
    Private Class TimeoutWebClient
        Inherits WebClient

        Public Property Timeout As Integer = UpdateCheckTimeoutMs

        Protected Overrides Function GetWebRequest(address As Uri) As WebRequest
            Dim request = MyBase.GetWebRequest(address)

            If request IsNot Nothing Then
                request.Timeout = Timeout
            End If

            Return request
        End Function
    End Class

    ''' <summary>
    ''' Kiểm tra kết nối mạng bằng URL cơ bản trong tài nguyên.
    ''' </summary>
    ''' <returns>True nếu mở được kết nối; ngược lại là False.</returns>
    Private Function IsNetAvail() As Boolean
        Try
            Using client As New TimeoutWebClient()
                Using stream = client.OpenRead(My.Resources.link_base)
                    Return stream IsNot Nothing
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Kiểm tra server version để biết có bản cập nhật mới hay không.
    ''' </summary>
    ''' <returns>True nếu server không còn chứa version hiện tại.</returns>
    Private Function IsUpdateAvailable() As Boolean
        If Not IsNetAvail() Then
            Return False
        End If

        Try
            Using client As New TimeoutWebClient()
                Dim serverVersion = client.DownloadString(My.Resources.link_ver)
                Return Not serverVersion.Contains(My.Resources.app_ver)
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Kiểm tra cập nhật và mở form tải trình cài đặt khi có phiên bản mới.
    ''' </summary>
    Private Sub ChkUpd()
        HdrSty("アップデートの確認...")

        If Not IsUpdateAvailable() Then
            Return
        End If

        MessageBox.Show($"「{My.Resources.app_true_name}」新しいバージョンが利用可能！", "更新", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Using frmUpd As New FrmUpdate()
            frmUpd.ShowDialog()
        End Using
    End Sub
#End Region

#Region "Helper"
    ''' <summary>
    ''' Lưu trạng thái license hợp lệ vào user settings.
    ''' </summary>
    Friend Sub UpdVldLic()
        My.Settings.Chk_Key = True
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' Fade in form theo từng bước nhỏ để form xuất hiện mượt hơn.
    ''' </summary>
    ''' <param name="frm">Form cần fade in.</param>
    <Extension()>
    Friend Sub FIFrm(frm As Form)
        If frm Is Nothing Then
            Return
        End If

        While frm.Opacity < 1
            frm.Opacity = Math.Min(1, frm.Opacity + 0.05R)
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Fade out form theo từng bước nhỏ trước khi đóng.
    ''' </summary>
    ''' <param name="frm">Form cần fade out.</param>
    <Extension()>
    Friend Sub FOFrm(frm As Form)
        If frm Is Nothing Then
            Return
        End If

        While frm.Opacity > 0
            frm.Opacity = Math.Max(0, frm.Opacity - 0.05R)
            frm.Update()
            Sleep(10)
        End While
    End Sub

    ''' <summary>
    ''' Giải phóng COM object để hạn chế Excel.exe còn treo sau khi xử lý xong.
    ''' </summary>
    ''' <param name="comObject">Đối tượng COM cần giải phóng.</param>
    Private Sub ReleaseComObject(comObject As Object)
        If comObject IsNot Nothing AndAlso Marshal.IsComObject(comObject) Then
            Marshal.FinalReleaseComObject(comObject)
        End If
    End Sub
#End Region

#Region "Master"
    ''' <summary>
    ''' Kết thúc tất cả process theo tên, bỏ qua process đã đóng hoặc không đủ quyền.
    ''' </summary>
    ''' <param name="name">Tên process không kèm .exe.</param>
    Friend Sub KillPrcs(name As String)
        If String.IsNullOrWhiteSpace(name) Then
            Return
        End If

        For Each item As Process In Process.GetProcessesByName(name)
            Try
                item.Kill()
                item.WaitForExit(3000)
            Catch ex As Exception When TypeOf ex Is InvalidOperationException OrElse TypeOf ex Is Win32Exception
                ' Process có thể đã tự đóng hoặc không cho phép kill; bỏ qua để luồng chính tiếp tục.
            Finally
                item.Dispose()
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Nhắc người dùng đóng Excel rồi kết thúc process Excel trước khi ứng dụng mở sổ tính.
    ''' </summary>
    Private Sub KillXl()
        Clear()
        HdrSty("警告：このアプリケーションを使用する前に、すべての「エクセル」を閉じてください。「エンター」キーを押して続行します...")
        ReadLine()
        KillPrcs(XL_NAME)
    End Sub

    ''' <summary>
    ''' Chạy luồng chính: kiểm tra update, chọn file Excel, ghi dữ liệu và mở lại file đã lưu.
    ''' </summary>
    Friend Sub RunApp()
        ChkUpd()
        KillXl()

        Using ofd As New OpenFileDialog With {
            .Multiselect = False,
            .Title = "「エクセル」ドキュメントを開く",
            .Filter = "「エクセル」ドキュメント|*.xlsx;*.xls"
        }
            If ofd.ShowDialog() = DialogResult.OK Then
                ProcessWorkbook(ofd.FileName)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Mở sổ tính bằng Excel COM, ghi dữ liệu nghiệp vụ, lưu và mở file bằng ứng dụng mặc định.
    ''' </summary>
    ''' <param name="filePath">Đường dẫn file Excel cần xử lý.</param>
    Private Sub ProcessWorkbook(filePath As String)
        Dim xlApp As Excel.Application = Nothing
        Dim workbook As Excel.Workbook = Nothing
        Dim workbookClosed As Boolean = False

        Try
            xlApp = New Excel.Application()
            workbook = xlApp.Workbooks.Open(filePath)

            WtIbarakiPcaBF(xlApp)

            workbook.Close(SaveChanges:=True)
            workbookClosed = True

            Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show($"Excelファイルの処理に失敗しました。{Environment.NewLine}{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If workbook IsNot Nothing AndAlso Not workbookClosed Then
                Try
                    workbook.Close(SaveChanges:=False)
                Catch ex As Exception
                    ' Nếu sổ tính đã bị Excel đóng trước đó thì chỉ cần tiếp tục dọn COM.
                End Try
            End If

            If xlApp IsNot Nothing Then
                Try
                    xlApp.Quit()
                Catch ex As Exception
                    ' Excel có thể đã tự thoát sau lỗi COM; bỏ qua để release RCW bên dưới.
                End Try
            End If

            ReleaseComObject(workbook)
            ReleaseComObject(xlApp)
        End Try
    End Sub
#End Region

#Region "Main"
    ''' <summary>
    ''' Tạo thư mục nếu đường dẫn hợp lệ và thư mục chưa tồn tại.
    ''' </summary>
    ''' <param name="path">Đường dẫn thư mục.</param>
    Friend Sub CrtDirAdv(path As String)
        If Not String.IsNullOrWhiteSpace(path) AndAlso Not Directory.Exists(path) Then
            Directory.CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Xóa file nếu đường dẫn hợp lệ và file đang tồn tại.
    ''' </summary>
    ''' <param name="path">Đường dẫn file.</param>
    Friend Sub DelFileAdv(path As String)
        If Not String.IsNullOrWhiteSpace(path) AndAlso File.Exists(path) Then
            File.Delete(path)
        End If
    End Sub

    ''' <summary>
    ''' Hiển thị câu hỏi 1/0 ở phần header và lặp cho tới khi người dùng nhập đúng.
    ''' </summary>
    ''' <param name="caption">Nội dung câu hỏi.</param>
    ''' <returns>1 nếu chọn có; 0 nếu chọn không.</returns>
    Friend Function HdrYnq(caption As String) As Integer
        Dim value = HdrDWrng(caption)

        Do Until value = 0 OrElse value = 1
            value = HdrDErr(caption)
        Loop

        Return CInt(value)
    End Function

    ''' <summary>
    ''' Ghi trực tiếp giá trị vào ô Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    Friend Sub DctVal(xlApp As Excel.Application, cell As String, value As Object)
        SetCellValue(xlApp, cell, value)
    End Sub

    ''' <summary>
    ''' Ghi giá trị vào ô Excel và tô màu để đánh dấu ô được override.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    Private Sub ModVal(xlApp As Excel.Application, cell As String, value As Object)
        SetCellValue(xlApp, cell, value, highlight:=True)
    End Sub

    ''' <summary>
    ''' Ghi giá trị vào range mà không cần Activate/ActiveCell, giúp giảm phụ thuộc trạng thái UI Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị cần ghi.</param>
    ''' <param name="highlight">True nếu cần tô màu ô.</param>
    Private Sub SetCellValue(xlApp As Excel.Application, cell As String, value As Object, Optional highlight As Boolean = False)
        Dim range As Excel.Range = Nothing

        Try
            range = xlApp.Range(cell)
            range.FormulaR1C1 = value

            If highlight Then
                range.Interior.Color = ModifiedCellColor
            End If
        Finally
            ReleaseComObject(range)
        End Try
    End Sub

    ''' <summary>
    ''' Xóa nội dung của merge area chứa ô Excel được chỉ định.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    Friend Sub ClrVal(xlApp As Excel.Application, cell As String)
        Dim range As Excel.Range = Nothing
        Dim mergeArea As Excel.Range = Nothing

        Try
            range = xlApp.Range(cell)
            mergeArea = range.MergeArea
            mergeArea.ClearContents()
        Finally
            ReleaseComObject(mergeArea)
            ReleaseComObject(range)
        End Try
    End Sub

    ''' <summary>
    ''' Nhập chuỗi từ console rồi ghi vào ô Excel.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="caption">Nội dung lời nhắc.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    Friend Sub PubSVal(xlApp As Excel.Application, caption As String, cell As String)
        DctVal(xlApp, cell, DtlSInp(caption))
    End Sub

    ''' <summary>
    ''' Ghi số vào ô Excel khi giá trị lớn hơn 0.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="cell">Địa chỉ ô.</param>
    ''' <param name="value">Giá trị số.</param>
    Friend Sub PubDVal(xlApp As Excel.Application, cell As String, value As Double)
        If value > 0 Then
            DctVal(xlApp, cell, value)
        End If
    End Sub

    ''' <summary>
    ''' Ghi tên, trọng lượng và số lượng thép vào các cột chuẩn AH/CM/BA.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="row">Dòng Excel.</param>
    ''' <param name="name">Tên thép.</param>
    ''' <param name="weight">Trọng lượng.</param>
    ''' <param name="number">Số lượng.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Ghi tiêu đề, tên, trọng lượng và số lượng thép vào các cột chuẩn X/AH/CM/BA.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="row">Dòng Excel.</param>
    ''' <param name="title">Tiêu đề hạng mục.</param>
    ''' <param name="name">Tên thép.</param>
    ''' <param name="weight">Trọng lượng.</param>
    ''' <param name="number">Số lượng.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, title As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Ghi đường kính, tiêu đề, tên, trọng lượng và số lượng thép vào các cột chuẩn S/X/AH/CM/BA.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="row">Dòng Excel.</param>
    ''' <param name="d">Đường kính thép.</param>
    ''' <param name="title">Tiêu đề hạng mục.</param>
    ''' <param name="name">Tên thép.</param>
    ''' <param name="weight">Trọng lượng.</param>
    ''' <param name="number">Số lượng.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, d As String, title As String, name As String, weight As Double, number As Double)
        If number > 0 Then
            DctVal(xlApp, $"S{row}", d)
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            DctVal(xlApp, $"BA{row}", number)
        End If
    End Sub

    ''' <summary>
    ''' Ghi đường kính, tiêu đề, tên, trọng lượng, đơn giá và số lượng thép.
    ''' </summary>
    ''' <param name="xlApp">Ứng dụng Excel.</param>
    ''' <param name="row">Dòng Excel.</param>
    ''' <param name="d">Đường kính thép.</param>
    ''' <param name="title">Tiêu đề hạng mục.</param>
    ''' <param name="name">Tên thép.</param>
    ''' <param name="weight">Trọng lượng.</param>
    ''' <param name="pn">Đơn giá và số lượng.</param>
    Friend Sub PubDModVal(xlApp As Excel.Application, row As String, d As String, title As String, name As String, weight As Double, pn As (price As Double, number As Double))
        If pn.number > 0 Then
            DctVal(xlApp, $"S{row}", d)
            DctVal(xlApp, $"X{row}", title)
            DctVal(xlApp, $"AH{row}", name)
            ModVal(xlApp, $"CM{row}", weight)
            ModVal(xlApp, $"CQ{row}", pn.price)
            DctVal(xlApp, $"BA{row}", pn.number)
        End If
    End Sub
#End Region

#Region "Timer"
    ''' <summary>
    ''' Start timer nếu timer chưa chạy.
    ''' </summary>
    ''' <param name="tmr">Timer cần start.</param>
    <Extension()>
    Friend Sub StrtAdv(tmr As Timer)
        If tmr IsNot Nothing AndAlso Not tmr.Enabled Then
            tmr.Start()
        End If
    End Sub

    ''' <summary>
    ''' Stop timer nếu timer đang chạy.
    ''' </summary>
    ''' <param name="tmr">Timer cần stop.</param>
    <Extension()>
    Friend Sub StopAdv(tmr As Timer)
        If tmr IsNot Nothing AndAlso tmr.Enabled Then
            tmr.Stop()
        End If
    End Sub
#End Region

#Region "Actor"
    ''' <summary>
    ''' Ghi header bằng màu vàng đậm.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Private Sub HdrSty(caption As String)
        ForegroundColor = DarkYellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi phần giới thiệu bằng màu xanh dương.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Private Sub IntroSty(caption As String)
        ForegroundColor = Blue
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi tiêu đề bằng màu xanh lá.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Private Sub TitSty(caption As String)
        ForegroundColor = Green
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi lời nhắc nhập liệu bằng màu xanh lơ.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Private Sub InpSty(caption As String)
        ForegroundColor = Cyan
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi mô tả phụ bằng màu magenta.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Private Sub DescSty(caption As String)
        ForegroundColor = Magenta
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi cảnh báo bằng màu vàng.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Private Sub WrngSty(caption As String)
        ForegroundColor = Yellow
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi lỗi bằng màu đỏ.
    ''' </summary>
    ''' <param name="caption">Nội dung cần ghi.</param>
    Friend Sub ErrSty(caption As String)
        ForegroundColor = Red
        Write(caption)
    End Sub

    ''' <summary>
    ''' Ghi prefix nhập liệu rồi trả màu chữ về trắng.
    ''' </summary>
    ''' <param name="caption">Nội dung prefix.</param>
    Private Sub PrefInp(caption As String)
        InpSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Ghi prefix chọn lựa rồi trả màu chữ về trắng.
    ''' </summary>
    ''' <param name="caption">Nội dung prefix.</param>
    Private Sub PrefSel(caption As String)
        WrngSty(caption)
        ForegroundColor = White
    End Sub

    ''' <summary>
    ''' Ghi prefix cảnh báo: nhãn vàng và phần nhập màu đỏ.
    ''' </summary>
    ''' <param name="caption">Nội dung prefix.</param>
    Private Sub PrefWrng(caption As String)
        WrngSty(caption)
        ForegroundColor = Red
    End Sub

    ''' <summary>
    ''' Ghi mô tả phụ trước dấu hai chấm của lời nhắc.
    ''' </summary>
    ''' <param name="description">Mô tả phụ.</param>
    Private Sub SfxDesc(description As String)
        DescSty(description)
        PrefInp(": ")
    End Sub

    ''' <summary>
    ''' Vẽ lại màn hình intro trước mỗi nhóm câu hỏi.
    ''' </summary>
    Private Sub Intro()
        Clear()
        IntroSty(My.Resources.gr_name & vbCrLf)
        IntroSty(My.Resources.cc_text & vbCrLf)
        TitSty(vbCrLf & My.Resources.app_true_name & vbCrLf & vbCrLf)
    End Sub

    ''' <summary>
    ''' Hiển thị intro rồi hỏi một giá trị số.
    ''' </summary>
    ''' <param name="caption">Nội dung lời nhắc.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDInp(caption As String) As Double
        Intro()
        Return DtlDInp(caption)
    End Function

    ''' <summary>
    ''' Hiển thị intro kèm thông báo cảnh báo.
    ''' </summary>
    ''' <param name="caption">Nội dung cảnh báo.</param>
    Friend Sub HdrWrng(caption As String)
        Intro()
        WrngSty(caption)
    End Sub

    ''' <summary>
    ''' Hiển thị intro rồi hỏi giá trị số ở trạng thái chọn lựa.
    ''' </summary>
    ''' <param name="caption">Nội dung lời nhắc.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDWrng(caption As String) As Double
        Intro()
        PrefSel(caption)
        Return ReadDoubleFromConsole()
    End Function

    ''' <summary>
    ''' Hiển thị intro rồi hỏi lại giá trị số ở trạng thái lỗi.
    ''' </summary>
    ''' <param name="caption">Nội dung lời nhắc.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function HdrDErr(caption As String) As Double
        Intro()
        PrefWrng(caption)
        Return ReadDoubleFromConsole()
    End Function

    ''' <summary>
    ''' Hỏi một giá trị số trong phần chi tiết hiện tại.
    ''' </summary>
    ''' <param name="caption">Nội dung lời nhắc.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function DtlDInp(caption As String) As Double
        PrefInp(caption)
        Return ReadDoubleFromConsole()
    End Function

    ''' <summary>
    ''' Hỏi một giá trị chuỗi trong phần chi tiết hiện tại.
    ''' </summary>
    ''' <param name="caption">Nội dung lời nhắc.</param>
    ''' <returns>Chuỗi người dùng nhập; chuỗi rỗng nếu console trả Nothing.</returns>
    Friend Function DtlSInp(caption As String) As String
        PrefInp(caption)
        Return If(ReadLine(), String.Empty)
    End Function

    ''' <summary>
    ''' Hỏi một giá trị số kèm mô tả phụ trong phần chi tiết hiện tại.
    ''' </summary>
    ''' <param name="caption">Nội dung lời nhắc chính.</param>
    ''' <param name="description">Mô tả phụ.</param>
    ''' <returns>Giá trị số người dùng nhập.</returns>
    Friend Function DtlDInpDesc(caption As String, description As String) As Double
        InpSty(caption)
        SfxDesc(description)
        Return ReadDoubleFromConsole()
    End Function

    ''' <summary>
    ''' Đọc dữ liệu nhập từ console theo cách cũ của ứng dụng: nội dung không phải số sẽ thành 0.
    ''' </summary>
    ''' <returns>Giá trị Double sau khi parse bằng Val.</returns>
    Private Function ReadDoubleFromConsole() As Double
        Return Val(ReadLine())
    End Function
#End Region
End Module
