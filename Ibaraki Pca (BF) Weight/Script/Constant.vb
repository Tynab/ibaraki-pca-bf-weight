Imports System.Environment
Imports System.Environment.SpecialFolder
Imports System.IO

''' <summary>
''' Chứa hằng số và đường dẫn dùng chung cho ứng dụng.
''' </summary>
Friend Module Constant
    ''' <summary>
    ''' Tên process Excel dùng khi đóng các file Excel đang mở trước lúc xử lý.
    ''' </summary>
    Friend Const XL_NAME As String = "excel"

    ''' <summary>
    ''' Tên file cài đặt được tải về khi có phiên bản mới.
    ''' </summary>
    Friend ReadOnly FILE_SETUP_NAME As String = $"{My.Resources.app_name} Setup.msi"

    ''' <summary>
    ''' Thư mục AppData của người dùng hiện tại.
    ''' </summary>
    Friend ReadOnly BACK_PATH As String = GetFolderPath(ApplicationData)

    ''' <summary>
    ''' Thư mục riêng của ứng dụng trong AppData.
    ''' </summary>
    Friend ReadOnly FRNT_PATH As String = Path.Combine(BACK_PATH, My.Resources.co_name)

    ''' <summary>
    ''' Đường dẫn đầy đủ tới file cài đặt sau khi tải về.
    ''' </summary>
    Friend ReadOnly FILE_SETUP_ADR As String = Path.Combine(FRNT_PATH, FILE_SETUP_NAME)
End Module
