Public Class ClassPower
    Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" ( _
    ByVal hWnd As IntPtr, _
    ByVal wMsg As Int32, _
    ByVal wParam As Int32, _
    ByVal lParam As Int32) As Int32

    Private Const WM_SYSCOMMAND As Int32 = &H112
    Private Const SC_MONITORPOWER As Int32 = &HF170

    ''' <summary>Versetzt den Bildschirm in den Standby-Modus.</summary>
    ''' <param name="handle" >Das Handle des aufrufenden Programmes
    ''' z.b. Me.Handle.ToInt32</param>
    Public Sub Screen2Standby(ByVal handle As Int32)
        SendMessage(handle, WM_SYSCOMMAND, SC_MONITORPOWER, 1)
    End Sub

End Class
