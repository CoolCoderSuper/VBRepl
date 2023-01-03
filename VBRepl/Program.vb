Friend Module Program

    'TODO: Expressions
    'TODO: State
    'TODO: Managing the compilation e.g references
    Public Sub Main(args As String())
        Dim prompt As New PrettyPrompt.Prompt
        Dim repl As New REPL
        While True
            Dim t As Task(Of PrettyPrompt.PromptResult) = prompt.ReadLineAsync()
            t.Wait()
            Dim res As PrettyPrompt.PromptResult = t.Result
            If res.IsSuccess Then
                repl.Evaluate(res.Text)
            End If
        End While
    End Sub

End Module