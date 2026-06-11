Imports System.ComponentModel
Imports System.IO
Imports System.Math
Imports System.Net
Imports System.Windows.Forms
Imports System.Windows.Forms.Keys

''' <summary>
''' Form tải bản cập nhật: hiển thị tiến độ, tải trình cài đặt và chạy file khi hoàn tất.
''' </summary>
Public Class FrmUpdate
#Region "Fields"
    ''' <summary>
    ''' WebClient có thời gian chờ để tránh form cập nhật treo quá lâu khi server không phản hồi.
    ''' </summary>
    Private Class UpdateWebClient
        Inherits WebClient

        Public Property Timeout As Integer = 30000

        Protected Overrides Function GetWebRequest(address As Uri) As WebRequest
            Dim request = MyBase.GetWebRequest(address)

            If request IsNot Nothing Then
                request.Timeout = Timeout
            End If

            Return request
        End Function
    End Class

    Private ReadOnly _wc As New UpdateWebClient()
    Private _downloadCompletedSuccessfully As Boolean
#End Region

#Region "Overridden"
    ''' <summary>
    ''' Ẩn form khỏi Alt+Tab bằng cách thêm tool-window style.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H80
            Return cp
        End Get
    End Property

    ''' <summary>
    ''' Chặn Alt+F4 để người dùng không đóng form giữa lúc đang cập nhật.
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Return keyData = (Alt Or F4) OrElse MyBase.ProcessCmdKey(msg, keyData)
    End Function
#End Region

#Region "Events"
    ''' <summary>
    ''' Reset trạng thái hiển thị khi form load.
    ''' </summary>
    Private Sub FrmUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ResetProgress()
    End Sub

    ''' <summary>
    ''' Chuẩn bị thư mục và bắt đầu tải trình cài đặt khi form đã hiển thị.
    ''' </summary>
    Private Sub FrmUpdate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        FIFrm()

        Try
            CrtDirAdv(FRNT_PATH)
            DelFileAdv(FILE_SETUP_ADR)

            AddHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
            AddHandler _wc.DownloadFileCompleted, AddressOf Upd_DownloadFileCompleted

            Dim setupUrl = _wc.DownloadString(My.Resources.link_app).Trim()

            If String.IsNullOrWhiteSpace(setupUrl) Then
                Throw New InvalidOperationException("Update URL is empty.")
            End If

            _wc.DownloadFileAsync(New Uri(setupUrl), FILE_SETUP_ADR)
        Catch ex As Exception
            HandleDownloadFailure(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Cập nhật dung lượng, phần trăm và thanh tiến độ khi WebClient báo tiến độ.
    ''' </summary>
    Private Sub Upd_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)
        Dim totalText = If(e.TotalBytesToReceive > 0, ToMbText(e.TotalBytesToReceive), "?")

        lblCapacity.Text = $"{ToMbText(e.BytesReceived)} MB / {totalText} MB"
        lblPercent.Text = $"{e.ProgressPercentage}%"
        pnlProgressBar.Width = CInt(Ceiling(e.ProgressPercentage * pnlMain.ClientSize.Width / 100D))
    End Sub

    ''' <summary>
    ''' Đóng form khi file đã tải xong, hoặc hiển thị lỗi nếu tải thất bại.
    ''' </summary>
    Private Sub Upd_DownloadFileCompleted(sender As Object, e As AsyncCompletedEventArgs)
        If e.Cancelled Then
            HandleDownloadFailure(New OperationCanceledException("Download was cancelled."))
            Return
        End If

        If e.Error IsNot Nothing Then
            HandleDownloadFailure(e.Error)
            Return
        End If

        If Not File.Exists(FILE_SETUP_ADR) Then
            HandleDownloadFailure(New FileNotFoundException("Downloaded installer was not found.", FILE_SETUP_ADR))
            Return
        End If

        _downloadCompletedSuccessfully = True
        lblPercent.Text = "100%"
        pnlProgressBar.Width = pnlMain.ClientSize.Width
        Close()
    End Sub

    ''' <summary>
    ''' Bộ hẹn giờ dự phòng: nếu trạng thái hoàn tất đã được đặt thì đóng form.
    ''' </summary>
    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles tmrMain.Tick
        If _downloadCompletedSuccessfully Then
            tmrMain.StopAdv()
            Close()
        End If
    End Sub

    ''' <summary>
    ''' Hủy lượt tải đang chạy nếu form bị đóng trước khi hoàn tất, sau đó làm mờ form.
    ''' </summary>
    Private Sub FrmUpdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Not _downloadCompletedSuccessfully AndAlso _wc.IsBusy Then
            _wc.CancelAsync()
        End If

        FOFrm()
    End Sub

    ''' <summary>
    ''' Dọn sự kiện/WebClient và chỉ chạy trình cài đặt khi lượt tải hoàn tất thành công.
    ''' </summary>
    Private Sub FrmUpdate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        RemoveHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        RemoveHandler _wc.DownloadFileCompleted, AddressOf Upd_DownloadFileCompleted
        tmrMain.StopAdv()
        _wc.Dispose()

        If Not _downloadCompletedSuccessfully Then
            Return
        End If

        Try
            Process.Start(New ProcessStartInfo(FILE_SETUP_ADR) With {.UseShellExecute = True})
            KillPrcs(My.Resources.app_name)
        Catch ex As Exception
            MessageBox.Show($"インストーラーを起動できませんでした。{Environment.NewLine}{ex.Message}", "更新", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Private helpers"
    ''' <summary>
    ''' Đặt lại nội dung hiển thị và thanh tiến độ về trạng thái ban đầu.
    ''' </summary>
    Private Sub ResetProgress()
        lblCapacity.Text = ""
        lblPercent.Text = ""
        pnlProgressBar.Width = 1
    End Sub

    ''' <summary>
    ''' Đổi số byte sang MB với 2 chữ số thập phân.
    ''' </summary>
    ''' <param name="bytes">Số byte.</param>
    ''' <returns>Chuỗi MB đã định dạng.</returns>
    Private Function ToMbText(bytes As Long) As String
        Return (bytes / 1024D / 1024D).ToString("0.00")
    End Function

    ''' <summary>
    ''' Hiển thị lỗi tải cập nhật và đóng form mà không chạy trình cài đặt.
    ''' </summary>
    ''' <param name="ex">Ngoại lệ gốc.</param>
    Private Sub HandleDownloadFailure(ex As Exception)
        _downloadCompletedSuccessfully = False
        tmrMain.StopAdv()
        lblText.Text = "更新に失敗しました。"
        MessageBox.Show($"アップデートをダウンロードできませんでした。{Environment.NewLine}{ex.Message}", "更新", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Close()
    End Sub
#End Region
End Class
