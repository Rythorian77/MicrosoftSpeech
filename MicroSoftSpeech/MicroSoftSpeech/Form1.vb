Imports System.Speech.Recognition
Imports System.Timers
Imports Timer = System.Timers.Timer

Public Class Form1

    Private ReadOnly recEngine As New SpeechRecognitionEngine

    Public Declare Function SendMessageW Lib "user32.dll" (hWnd As IntPtr,
                                                           Msg As Integer,
                                                           wParam As IntPtr,
                                                           lParam As IntPtr) As IntPtr

    Private Sub Form1_Load(sender As Object,
                           e As EventArgs) Handles MyBase.Load
        Dim commands As New Choices
        Const V As String = "scarlett dictation mode"
        commands.Add((New String() {V}))
        Dim gramBuilder As New GrammarBuilder
        gramBuilder.Append(commands)
        Dim grammar As New Grammar(gramBuilder)
        Dim gram As Grammar = grammar
        recEngine.LoadGrammarAsync(gram)
        recEngine.SetInputToDefaultAudioDevice()
        AddHandler recEngine.SpeechRecognized, AddressOf RecEngine_SpeechRecognized
        recEngine.RecognizeAsync(RecognizeMode.Multiple)
    End Sub

    Private Sub TimerEvent(source As Object,
                           e As ElapsedEventArgs)
        SendKeys.SendWait($"{{ENTER}}") 'It is vital to use the "SendWait" method or the app will crash with error
    End Sub

    Private Sub TimerOracle(source As Object,
                            e As ElapsedEventArgs)
        SendKeys.SendWait($"%{{`}}") 'It is vital to use the "SendWait" method or the app will crash with error
        recEngine.RecognizeAsyncStop() 'Stops Recognition for app so Microsoft Word's Dictation can work without interference.
    End Sub

    Private Sub RecEngine_SpeechRecognized(sender As Object, e As SpeechRecognizedEventArgs)
        Select Case (e.Result.Text)

            Case "scarlett dictation mode" 'This will activate Microsoft Office Dictation
                Const FileName As String = "winword"
                Process.Start(FileName)

                'This will select a blank page automatically from Microsoft Word.
                Dim timer As New Timer With {
                    .Interval = &B101110111000 '3000 (3 seconds) converted to binary.
                }

                AddHandler timer.Elapsed, AddressOf TimerEvent

                timer.AutoReset = False 'This has to be kept "False" to prevent continual mis-fires
                timer.Enabled = True

                'Secondary Delayed Timer that activates mic in Microsoft Word.
                Dim timer1 As New Timer With {
                    .Interval = &B1001110001000 '5000 (5 seconds) converted to binary.
                }

                AddHandler timer1.Elapsed, AddressOf TimerOracle

                timer1.AutoReset = False 'This has to be kept "False" to prevent continual mis-fires
                timer1.Enabled = True

            Case Else
                Const Message As String = "Command not heard: scarlett dictation mode "
                Debug.WriteLine(Message)
        End Select
    End Sub

    'Don't add null check, it's a bad habit in coding, & avoiding arguments as much as possible is good coding.
    'If your code is sound, you shouldn't need un-needed arguments
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim form = TryCast(obj, Form1)
        Return form IsNot Nothing _
               AndAlso EqualityComparer(Of SpeechRecognitionEngine).Default.Equals(recEngine, form.recEngine)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Dim hashCode As Long = 146367575
        hashCode = ((hashCode _
                    * -1521134295) _
                    + EqualityComparer(Of SpeechRecognitionEngine).Default.GetHashCode(recEngine)).GetHashCode()
        Return hashCode
    End Function

End Class