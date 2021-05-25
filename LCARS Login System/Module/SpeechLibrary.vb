'Copyrighted By Gonzo@dosentreiber.de

'Dieser Script arbeitet mit Microsoft's Speech Recognition feature. Das Tool enthält folgende Funktionen:
' Text-to-Speech
'Spracherkennung

'IMPORT SYSTEM.SPEECH VON DEN REFEENZEN!!
' ,andererseits wird es nicht funzen!

Imports SpeechLib 'Needed for the speech recognition to work
Module SpeechLibrary

    'DECLARING THE VARIABLES THAT ARE NEEDED FOR VOICE RECOGNITION

    Dim WithEvents RecoContext As SpSharedRecoContext      'Used to store the charectar of each letter spoken
    Dim Grammar As ISpeechRecoGrammar                      'Used to store a dictionary of words and letters so the program can recognise them
    Dim CharCount As Integer 'Counts the amount of letters used at a time

    Public SpeechLibrary_ResultToString As String 'All the results from the text will be made here

    'This is the sub which activates once the command has been recognised 
    'Use this sub to say what happens if certain words are used.
    Private Sub SpeechLibrary_COMMAND_LIST()
        Dim sender As New Object
        Dim myE As New System.EventArgs
        '- - - - - Example Use - - - - -
        'If SpeechLibrary_ResultToString = "whatever you have said" Then
        ' DO THIS 
        'End If

        If SpeechLibrary_ResultToString = "Computer" Then
            My.Computer.Audio.Play(My.Resources.stiabdruck_berechtigungscode, _
                AudioPlayMode.Background)

            frmlogin.pnlMain.Visible = True
            frmlogin.pnlMain.BringToFront()
            frmlogin.n1.Visible = True
            frmlogin.n2.Visible = True
        End If

        If SpeechLibrary_ResultToString = "berechtigungsebene" Then
            frmlogin.user.Text = ("Gonzo")
            My.Computer.Audio.Play(My.Resources.lcars_events_voicereco_ok, _
                                  AudioPlayMode.WaitToComplete)

            frmlogin.pass.Text = ("0708050200")
            My.Computer.Audio.Play(My.Resources.lcars_events_voicereco_ok, _
                                  AudioPlayMode.WaitToComplete)
        End If

        If SpeechLibrary_ResultToString = "bestätigen" Then
            frmlogin.btnOk.doClick(sender, myE)
        Else
            'My.Computer.Audio.Play(My.Resources.befehl, _
            ' AudioPlayMode.Background)
        End If


        Select Case SpeechLibrary_ResultToString
            Case "null"
                frmlogin.hpb00.doClick(sender, myE)
            Case "eins"
                frmlogin.hpb02.doClick(sender, myE)
            Case "zwei"
                frmlogin.hpb02.doClick(sender, myE)
            Case "drei"
                frmlogin.hpb03.doClick(sender, myE)
            Case "vier"
                frmlogin.hpb04.doClick(sender, myE)
            Case "fünf"
                frmlogin.fb05.doClick(sender, myE)
            Case "sechs"
                frmlogin.fb06.doClick(sender, myE)
            Case "sieben"
                frmlogin.fb07.doClick(sender, myE)
            Case "acht"
                frmlogin.fb08.doClick(sender, myE)
            Case "neun"
                frmlogin.fb09.doClick(sender, myE)
            Case "berechtigungsebene"
                frmlogin.pass.Text = ("0708050200")
                My.Computer.Audio.Play(My.Resources.lcars_events_voicereco_ok, _
                              AudioPlayMode.WaitToComplete)
                frmlogin.user.Text = ("Gonzo")
                My.Computer.Audio.Play(My.Resources.lcars_events_voicereco_ok, _
                                      AudioPlayMode.WaitToComplete)
        End Select

        If SpeechLibrary_ResultToString = "löschen" Then
            frmlogin.btnclear.doClick(sender, myE)
            'My.Computer.Audio.Play(My.Resources.befehl, _
            'AudioPlayMode.Background)
        End If

        ' If SpeechLibrary_ResultToString = "hallo" Then
        'SpeechLibrary_TextToSpeech("Hello Sir..")
        'Else
        'SpeechLibrary_TextToSpeech("This command is not recognised.. Please speak more clearly..")
        'End If
    End Sub


    'This function will convert any text entered into it into speech!
    Public Sub SpeechLibrary_TextToSpeech(ByVal TextToSpeak)
        Dim SAPI
        SAPI = CreateObject("sapi.spvoice") 'sapi.spvoice' meaning the Speech engine
        SAPI.Speak(TextToSpeak) 'Speak what the user has requested.
    End Sub

    'This function will stop the speech recognition feature.
    Public Sub SpeechLibrary_StopSpeechRecognition()
        Grammar.DictationSetState(SpeechRuleState.SGDSInactive)   'Turns off the Recognition  
    End Sub

    'This function will start the speech recognition feature.
    'NOTE: This must be turned on before you can start to use speech recognition
    Public Sub SpeechLibrary_StartSpeechRecognition()
        'First check to see if reco has been loaded before. If not lets load it.  
        If (RecoContext Is Nothing) Then
            RecoContext = New SpSharedRecoContextClass        'Erstelle eine neue RecoContextClass  
            Grammar = RecoContext.CreateGrammar(1)            'Setup Grammar  
            Grammar.DictationLoad()                          'Lade Grammar  
        End If

        Beep()                'Make the computer beep so the user knows that speech recognition has started
        Grammar.DictationSetState(SpeechRuleState.SGDSActive)   'Turns on the Recognition  
    End Sub

    'This function is here for when the speech is recognised.
    Private Sub OnReco(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult) Handles RecoContext.Recognition
        Dim recoResult As String = Result.PhraseInfo.GetText 'Create a new string, and assign the recognized text to it.  
        Dim txt As New Windows.Forms.TextBox
        txt.SelectionStart = CharCount
        txt.SelectedText = recoResult.ToLower
        CharCount = CharCount + 1 + Len(recoResult)

        SpeechLibrary_ResultToString = txt.Text
        SpeechLibrary_COMMAND_LIST()
        txt.Clear()
    End Sub
End Module
