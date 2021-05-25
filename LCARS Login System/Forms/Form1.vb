Public Class Form1
    Private Sub Button1_Click(ByVal sender As System.Object, _
  ByVal e As System.EventArgs) Handles Button1.Click

        ' Array aller vorhandenen Bildschirme
        Dim oScreens() As Screen = Screen.AllScreens

        ' Anzahl vorhandener Bildschirme
        Dim nScreenCount As Integer = oScreens.Length

        ' Auflösung, WorkingArea etc. der einzelnen Bildschirme
        For i As Integer = 0 To nScreenCount - 1
            With ListBox1.Items
                .Add("Auflösung: " + oScreens(i).Bounds.Size.ToString)
                .Add("Start-Position: " & oScreens(i).Bounds.Location.ToString())
                .Add("Working Area: " & oScreens(i).WorkingArea.ToString())
                .Add("Primary Screen: " & oScreens(i).Primary.ToString())
            End With
        Next i

        Me.Text = "Anzahl Bildschirme: " & nScreenCount.ToString()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        frmlogin.Show()
        Me.Hide()

    End Sub
End Class
