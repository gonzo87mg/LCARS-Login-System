Module Funktionen
    ' Zunächst die benötigten API-Deklarationen
    Private Declare Function FindWindow Lib "user32" _
      Alias "FindWindowA" ( _
      ByVal lpClassName As String, _
      ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowPos Lib "user32" ( _
      ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

    Private Const SWP_SHOWWINDOW = &H40
    Private Const SWP_HIDEWINDOW = &H80

    ' Taskbar ausblenden
    Public Sub HideTaskbar()
        Dim hWnd As Long
        hWnd = FindWindow("Shell_TrayWnd", "")
        Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End Sub

    ' Taskbar einblenden
    Public Sub ShowTaskbar()
        Dim hWnd As Long
        hWnd = FindWindow("Shell_TrayWnd", "")
        Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End Sub
End Module
