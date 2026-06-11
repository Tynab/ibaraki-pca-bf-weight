Imports System.Console
Imports System.Text.Encoding
Imports System.Windows.Forms

''' <summary>
''' Điểm vào chính: kiểm tra license cục bộ rồi chạy luồng xử lý Excel.
''' </summary>
Public Module Main
    ''' <summary>
    ''' Thiết lập console UTF-8 và chỉ chạy ứng dụng khi license hợp lệ.
    ''' </summary>
    Public Sub Main()
        OutputEncoding = UTF8

        If My.Settings.Chk_Key OrElse PromptLicenseUntilValid() Then
            RunApp()
        End If
    End Sub

    ''' <summary>
    ''' Hỏi serial cho tới khi người dùng nhập đúng hoặc chọn hủy.
    ''' </summary>
    ''' <returns>True nếu license hợp lệ; False nếu người dùng hủy.</returns>
    Private Function PromptLicenseUntilValid() As Boolean
        Do
            If InputBox("シリアルを入力", "ライセンスキー") = My.Resources.key_ser Then
                UpdVldLic()
                Return True
            End If

            Dim retry = MessageBox.Show(
                "ライセンスが間違っています！",
                "エラー",
                MessageBoxButtons.RetryCancel,
                MessageBoxIcon.Error)

            If retry <> DialogResult.Retry Then
                ErrSty("終了するには、任意のキーを押してください...")
                ReadKey()
                Return False
            End If
        Loop
    End Function
End Module
