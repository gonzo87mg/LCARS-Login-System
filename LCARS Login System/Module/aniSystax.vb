Module aniSyntax


#Region "LoginSystem"
    Public Sub login()
        Dim size As New System.Drawing.Size
        Dim height As Integer

        size.Width = frmlogin.Width
        size.Height = 0
        frmlogin.Height = height
        size = size

    End Sub
    Public Sub loginAni()
        For i = 0 To 1050 Step 25
            frmlogin.Height = i
            If i = 1050 Then
                frmlogin.Timer2.Stop()
                frmlogin.Timer1.Start()
            End If
        Next

    End Sub
#End Region


End Module
