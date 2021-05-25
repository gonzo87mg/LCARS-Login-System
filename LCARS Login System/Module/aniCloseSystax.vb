Module aniCloseSyntax
    
    Public Sub start()
        Dim size As New System.Drawing.Size
        Dim height As Integer

        size.Width = frmlogin.Size.Width
        size.Height = frmlogin.Size.Height
        frmlogin.Height = height
        size = size
        If frmlogin.Height = 10 Then
            My.Computer.Audio.Play(My.Resources.beep, _
           AudioPlayMode.Background)
            'Application.Exit()
        End If

    End Sub
    Public Sub startAni()

        frmlogin.Height -= 30
        If frmlogin.Height = 10 Then

            frmlogin.Hide()
        End If
    End Sub
End Module
