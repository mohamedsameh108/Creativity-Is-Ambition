Imports System.Speech.Recognition
Public Class Form2

    Private Sub دخول_Click(sender As Object, e As EventArgs) Handles دخول.Click
        If (TextBox1.Text = "Egypt") Then
            form1.Show()
            Me.Hide()
        Else
            MessageBox.Show("please write or say Egypt in Text Box to contniue")
            TextBox1.Focus()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Dim speak As SpeechRecognitionEngine = New SpeechRecognitionEngine(New System.Globalization.CultureInfo("en-US"))
        speak.SetInputToDefaultAudioDevice()
        speak.LoadGrammar(New DictationGrammar())
        speak.RecognizeAsync(RecognizeMode.Multiple)
        AddHandler speak.SpeechRecognized, AddressOf speak_speak1
    End Sub
    Private Sub speak_speak1(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
        TextBox1.Text = (e.Result.Text)
    End Sub
End Class