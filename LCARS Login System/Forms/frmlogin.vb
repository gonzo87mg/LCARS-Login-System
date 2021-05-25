'Dieser Code ist (C) Copyright Rechtlich geschützt by Gonzo@dosentreiber.de -- www.dosentreiber.de -- Autor: Andrey Kurpiers
'### www.dosentreiber.de ###
'Wenn Sie diese Zeilen lesen können, dann ist es Ihnen gelungen, diese Software mittels Tools zu Dekompilieren.
'Ich mache darauf aufmerksam, dass jeder Dekompiliervorgang Strafrechtliche konsequenzen nach sich zieht!
'### Der Zugriff, sowie das kopieren oder ersetzen der einzelnen Ressourcen dieser Software ist untersagt! ###
'### Eingriffe ins System werden gespeichert und bei bestehender Internetverbindung direkt an den Autor gesendet, was eine nutzung der Software auf diesem Computer unmöglich macht! ###
'### Es werden keine erstellten User Accounts oder Passwörter übertragen!! ###

Imports LCARS.UI
Imports System.Runtime.InteropServices
Imports TextEffectsLib
Imports System.Threading

Public Class frmlogin

#Region "Funktionen"
    Declare Sub keybd_event Lib "user32" ( _
        ByVal bVk As Byte, _
        ByVal bScan As Byte, _
        ByVal dwFlags As Integer, _
        ByVal dwExtraInfo As Integer)

    Private Const KEYEVENTF_KEYUP = &H2

    Private Delegate Function HOOKPROCDelegate( _
      ByVal nCode As Integer, _
      ByVal wParam As IntPtr, _
      ByRef lParam As KBDLLHOOKSTRUCT) As IntPtr

    ' dauerhafte Delegaten-Variable erzeugen
    Private HookProc As New HOOKPROCDelegate(AddressOf KeyboardHookProc)

    Private Declare Unicode Function GetModuleHandleW Lib "kernel32.dll" ( _
      ByVal lpModuleName As IntPtr) As IntPtr

    ' Die Funktion, um einen globalen Hook setzen zu können:
    Private Declare Unicode Function SetWindowsHookExW Lib "user32.dll" ( _
      ByVal idHook As Integer, _
      ByVal lpfn As HOOKPROCDelegate, _
      ByVal hMod As IntPtr, _
      ByVal dwThreadId As UInteger) As IntPtr

    ' Für das Löschen des Hooks wird diese Funktion verwendet:
    Private Declare Unicode Function UnhookWindowsHookEx Lib "user32.dll" ( _
      ByVal hhk As IntPtr) As UInteger

    Private Declare Unicode Function CallNextHookEx Lib "user32.dll" ( _
      ByVal hhk As IntPtr, _
      ByVal nCode As Integer, _
      ByVal wParam As IntPtr, _
      ByRef lParam As KBDLLHOOKSTRUCT) As IntPtr

    Private Const WM_KEYDOWN As Int32 = &H100    ' Konstante für WM_KEYDOWN
    Private Const WM_KEYUP As Int32 = &H101      ' Konstante für WM_KEYUP
    Private Const HC_ACTION As Integer = 0       ' Konstante für HC_ACTION
    Private Const WH_KEYBOARD_LL As Integer = 13 ' Konstante für WH_KEYBOARD_LL

    Public PrevWndProc As Integer
    Private mHandle As IntPtr

    <StructLayout(LayoutKind.Sequential)> Public Structure KBDLLHOOKSTRUCT
        Public vkCode As Keys
        Public scanCode, flags, time, dwExtraInfo As UInteger

        Public Sub New(ByVal key As Keys, _
          ByVal scancod As UInteger, _
          ByVal flagss As UInteger, _
          ByVal zeit As UInteger, _
          ByVal extra As UInteger)

            vkCode = key
            scanCode = scancod
            flags = flagss
            time = zeit
            dwExtraInfo = extra
        End Sub
    End Structure

    ' Um den Hook ein-/ausschalten zu können:
    Public Property KeyHookEnable() As Boolean
        Get
            Return mHandle <> IntPtr.Zero
        End Get
        Set(ByVal value As Boolean)
            If KeyHookEnable = value Then Return
            If value Then
                mHandle = SetWindowsHookExW(WH_KEYBOARD_LL, HookProc, _
                  GetModuleHandleW(IntPtr.Zero), 0)
            Else
                UnhookWindowsHookEx(mHandle)
                mHandle = IntPtr.Zero
            End If
        End Set
    End Property

    ' Hiermit wird der Tastendruck VOR dem Betriebssystem abgefangen:
    ' wParam kann folgende Werte annehmen: 
    ' WM_KEYUP und WM_KEYDOWN (Taste gedrückt/losgelassen)
    ' wird fEatKeyStroke=true gesetzt, so wird dieser Tastendruck "verschluckt", 
    ' d. h. er hat für das System  NIE statt gefunden.
    Private Function KeyboardHookProc(ByVal nCode As Integer, _
      ByVal wParam As IntPtr, _
      ByRef lParam As KBDLLHOOKSTRUCT) As IntPtr

        Dim fEatKeyStroke As Boolean

        If nCode = HC_ACTION Then

            Select Case lParam.vkCode
                ' Hier z.B. die Zeichen a und b nicht zulassen 
                '(fEatKeyStroke = True)
                Case Keys.RWin
                    fEatKeyStroke = True
                Case Keys.LWin
                    fEatKeyStroke = True
                Case Keys.Alt
                    fEatKeyStroke = True
                Case Keys.Escape
                    fEatKeyStroke = True
                Case Keys.F4
                    fEatKeyStroke = True
            End Select

            If fEatKeyStroke Then
                Return New IntPtr(1)
                Exit Function
            End If

            Return CallNextHookEx(mHandle, nCode, wParam, lParam)
        End If
    End Function
#End Region

#Region "Form Region"

    Sub PlayBackgroundSoundResource()
        My.Computer.Audio.Play(My.Resources.autorisierungscode_erforderlich, _
            AudioPlayMode.Background)
    End Sub

    Private Sub frmlogin_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Funktionen.ShowTaskbar()

        Timer3.Stop()

        Application.DoEvents()

        Dispose()

        KeyHookEnable = False

        Application.Exit()

    End Sub

    Private Sub frmlogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        SpeechLibrary_StartSpeechRecognition()
        Funktionen.HideTaskbar()
        KeyHookEnable = True

        ' Call aniSyntax.login()
        'Timer2.Start()

        pnlMain.Visible = False

        Panel2.BringToFront()

        exist.Visible = True
        infolabel.Text = ("Bitte Autorisierungscode eingeben !")
        Label3.Visible = True
        btnOk.Visible = True

        'Liste mit reg. Usern füllen
        For Each found As String In My.Computer.FileSystem.GetDirectories(Application.StartupPath & "\Accounts\", FileIO.SearchOption.SearchTopLevelOnly)
            Dim title As String = System.IO.Path.GetFileNameWithoutExtension(found)
            lb1.Items.Add(title)
        Next
        lblAnzahl.Text = "Einträge: " & lb1.Items.Count
    End Sub
#End Region

#Region "User Anmeldung"
    Private Sub btnok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If user.Text = "" Then
            My.Computer.Audio.Play(My.Resources.sicherheitscode, _
                 AudioPlayMode.Background)
        Else
            Try
                Dim sa As String
                Dim sb As String

                Dim a As New System.IO.StreamReader(Application.StartupPath & "\Accounts\" + user.Text + "\048ur37.lcars")
                sa = a.ReadLine
                a.Close()
                Dim b As New System.IO.StreamReader(Application.StartupPath & "\Accounts\" + user.Text + "\159pw48.lcars")
                sb = b.ReadLine
                b.Close()

                If user.Text = sa.ToString Then
                    If pass.Text = sb.ToString Then
                        My.Computer.Audio.Play(My.Resources.zugang_genemigt, _
                           AudioPlayMode.Background)
                        Funktionen.ShowTaskbar()

                        Application.Exit()
                        End

                    Else
                        My.Computer.Audio.Play(My.Resources.keinZugang, _
                            AudioPlayMode.Background)
                    End If
                End If

            Catch ex As Exception
                Label9.Visible = True
                Label9.Text = ("Username existiert nicht ! " + ex.Message)
                My.Computer.Audio.Play(My.Resources.keinZugang, _
                            AudioPlayMode.Background)
            End Try

        End If
    End Sub
#End Region

#Region "TimerZ"
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        My.Computer.Audio.Play(My.Resources.autorisierungscode_erforderlich, _
       AudioPlayMode.Background)
        Timer1.Enabled = False
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Call aniSyntax.loginAni()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        For Each Taskmgr In Process.GetProcessesByName("TaskMgr")
            Taskmgr.Kill()
        Next
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        If Not Me.WindowState = FormWindowState.Minimized Then
            If Not Me.Bounds = Screen.PrimaryScreen.WorkingArea Then
                Me.Bounds = Screen.PrimaryScreen.WorkingArea
            End If
        End If
    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick

        If lbl1.ForeColor = Color.Yellow Then
            lbl1.ForeColor = Color.LightGray
        Else
            lbl1.ForeColor = Color.LightGray
            lbl1.ForeColor = Color.Yellow
        End If
    End Sub

    Private Sub Timer6_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer6.Tick
        'lbl2.Visible = True
        'lbl3.Visible = False
        TextEffects.Enlarge_From_Left(CType(lbl2, Object), lbl2.Text, 20, False)
        'lbl3.Visible = True
        TextEffects.Enlarge_From_Left(CType(lbl3, Object), lbl3.Text, 20, False)
        Thread.Sleep(50)
        'lbl2.Visible = False
        'lbl3.Visible = False

    End Sub
#End Region

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        newusraccount.Visible = False
        exist.Visible = True
        infolabel.Text = ("Bitte Autorisierungscode eingeben !")
        Label3.Visible = True
    End Sub


    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        btnOk.Visible = True
        newusraccount.Visible = False
        exist.Visible = True
        infolabel.Text = ("Bitte Autorisierungscode eingeben !")
        Label3.Visible = True

        My.Computer.Audio.Play(My.Resources.autorisierungscode, _
           AudioPlayMode.Background)
    End Sub

    Private Sub X32TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles X32TabPage1.Click
        infolabel.Text = ("Bitte Autorisierungscode eingeben !")
        If Not btnOk.Visible Then
            btnOk.Visible = True
        End If

        My.Computer.Audio.Play(My.Resources.autorisierungscode, _
           AudioPlayMode.Background)
    End Sub

    Private Sub X32TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles X32TabPage2.Click
        btnOk.Visible = False
    End Sub
#Region "New User"
    Private Sub ComplexButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComplexButton1.Click


        If newpass.Text = "" Then
            My.Computer.Audio.Play(My.Resources.Kommandoautorisation, _
                          AudioPlayMode.Background)
            infolabel.Text = "Das eingegebene Passwort ist nicht das Administratorpasswort. Bitte prüfen Sie ihre Eingabe!"
        End If


        If newpass2.Text = "030501040200010201010809" Then

            If My.Computer.FileSystem.DirectoryExists(Application.StartupPath & "\Accounts\") Then
                'Wenn der Ordner Accounts existiert, dann tu nichts!
            Else
                'Wenn nicht, dann erstelle neues Verzeichnis
                MkDir(Application.StartupPath & "\Accounts\")

                'Zeige Fehlerausgabe in Label3, wenn der eingegebene Benutzername bereits existiert
                If newpass2.Text = "030501040200010201010809" Then

                    If My.Computer.FileSystem.DirectoryExists(Application.StartupPath & "\Accounts\" + newuser.Text) Then
                        Label3.Text = ("Fehler! Der Account existiert bereits!")
                        My.Computer.Audio.Play(My.Resources.befehl, _
                          AudioPlayMode.Background)

                        Application.DoEvents()

                        Exit Sub

                    Else
                        'Wenn der angegebene Benutzer nicht existiert, dann lege neue Datei im Verzeichnis an
                        MkDir(Application.StartupPath & "\Accounts\" + newuser.Text)

                        Dim a As New System.IO.StreamWriter(Application.StartupPath & "\Accounts\" + newuser.Text + "\048ur37.lcars")
                        a.WriteLine(newuser.Text)
                        a.Close()

                        Dim b As New System.IO.StreamWriter(Application.StartupPath & "\Accounts\" + newuser.Text + "\159pw48.lcars")
                        b.WriteLine(newpass.Text)
                        b.Close()

                        'Zeige Meldung bei erfolgreich im Infolabel 
                        infolabel.Text = ("Account erfolgreich erstellt !")

                        LCARS.UI.MsgBox("Account erfolgreich erstellt !", MsgBoxStyle.OkOnly, "System Login")
                        If MsgBoxResult.Ok Then

                            'Hintergrundstimme als Bestätigung für erfolgreichen Zugang zum System
                            My.Computer.Audio.Play(My.Resources.newuser, _
                                           AudioPlayMode.Background)

                            Application.DoEvents()

                            'Zeigt einen Ladebildschirm

                            Funktionen.ShowTaskbar()
                            Application.DoEvents()

                            End
                        End If
                    End If
                End If
            End If
        End If
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        My.Computer.Audio.Play(My.Resources.befehl, _
           AudioPlayMode.WaitToComplete)
        My.Computer.Audio.Play(My.Resources.hauptenergie_verschlüsselt, _
           AudioPlayMode.WaitToComplete)
        Application.DoEvents()

    End Sub
#Region "DeathLine"
    Private Sub btnOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        My.Computer.Audio.Play(My.Resources.be1, _
                   AudioPlayMode.Background)
        If user.Text = "" Then
            My.Computer.Audio.Play(My.Resources.sicherheitscode, _
                 AudioPlayMode.Background)
        Else

            Try
                Dim sa As String
                Dim sb As String

                Dim a As New System.IO.StreamReader(Application.StartupPath & "\Accounts\" + user.Text + "\048ur37.lcars")
                sa = a.ReadLine
                a.Close()
                Dim b As New System.IO.StreamReader(Application.StartupPath & "\Accounts\" + user.Text + "\159pw48.lcars")
                sb = b.ReadLine
                b.Close()

                If user.Text = sa.ToString Then
                    If pass.Text = sb.ToString Then
                        My.Computer.Audio.Play(My.Resources.besteatigt, _
               AudioPlayMode.WaitToComplete)
                        My.Computer.Audio.Play(My.Resources.zugang_genemigt, _
                           AudioPlayMode.WaitToComplete)
                    Else
                        My.Computer.Audio.Play(My.Resources.keinZugang, _
                            AudioPlayMode.WaitToComplete)
                        My.Computer.Audio.Play(My.Resources.Kommandoautorisation, _
                      AudioPlayMode.Background)
                    End If
                End If

            Catch ex As Exception
                Label9.Text = ("Username existiert nicht ! " + ex.Message)
                My.Computer.Audio.Play(My.Resources.keinZugang, _
                            AudioPlayMode.Background)
            End Try
        End If
    End Sub
#End Region

    Private Sub user_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles user.Click
        btnOk.Visible = True
    End Sub

    Private Sub newuser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles newuser.Click
        btnOk.Visible = False
    End Sub

    Private Sub pass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles pass.KeyDown
        Dim myE As New System.EventArgs

        If e.KeyCode = Keys.Enter Then
            btnOk_Click_1(sender, myE)
        End If
    End Sub

    Private Sub lb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lb1.SelectedIndexChanged
        user.Text = lb1.SelectedItem
        pass.Focus()
    End Sub

    Private Sub Panel2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel2.Click
        Panel2.Hide()
        pnlMain.Visible = True
        pnlMain.BringToFront()
        My.Computer.Audio.Play(My.Resources.autorisierungscode_erforderlich, _
            AudioPlayMode.Background)
        Timer6.Stop()
        Application.DoEvents()
    End Sub
#Region "Nummernblock"
    Private Sub HalfPillButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hpb00.Click
        My.Computer.Audio.Play(My.Resources.be, _
            AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("00")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("00")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("00")
        End If
        n1.Visible = True
        n2.Visible = True

    End Sub

    Private Sub HalfPillButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hpb01.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("01")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("01")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("01")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub HalfPillButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hpb02.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("02")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("02")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("02")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub HalfPillButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hpb03.Click
        My.Computer.Audio.Play(My.Resources.be1, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("03")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("03")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("03")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub HalfPillButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hpb04.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("04")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("04")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("04")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub FlatButton18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fb05.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("05")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("05")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("05")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub FlatButton19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fb06.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("06")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("06")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("06")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub FlatButton20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fb07.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("07")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("07")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("07")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub FlatButton21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fb08.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("08")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("08")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("08")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub FlatButton22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fb09.Click
        My.Computer.Audio.Play(My.Resources.be, _
                   AudioPlayMode.Background)
        If pass.Focused Then
            pass.Text = pass.Text + ("09")
        ElseIf newpass.Focused Then
            newpass.Text = newpass.Text + ("09")
        ElseIf newpass2.Focused Then
            newpass2.Text = newpass2.Text + ("09")
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub

    Private Sub FlatButton25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclear.Click

        If pass.Focused Then
            pass.Clear()
        ElseIf newpass.Focused Then
            newpass.Clear()
        ElseIf newpass2.Focused Then
            newpass2.Clear()
        End If
        n1.Visible = True
        n2.Visible = True
    End Sub
#End Region

    Private Sub newpass2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles newpass2.GotFocus
        My.Computer.Audio.Play(My.Resources.Kommandoautorisation, _
                      AudioPlayMode.Background)
    End Sub

    Private Sub pass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pass.TextChanged

    End Sub

    Private Sub FlatButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FlatButton6.Click
        Panel2.Visible = True
        Panel2.BringToFront()

        pnlMain.Visible = False
        Timer6.Start()
        Application.DoEvents()
    End Sub
    
    Private Sub FlatButton15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FlatButton15.Click
        ShowTaskbar()
        Application.DoEvents()
        Dispose()
        Application.Exit()
    End Sub
End Class
