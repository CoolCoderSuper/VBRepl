Imports Microsoft.CodeAnalysis
Imports Spectre.Console
Imports DeveloperCore.REPL

Public Module Program

    'TODO: Managing the compilation e.g references
    'TODO: Default imports
    Public Sub Main(args As String())
        Dim prompt As New PrettyPrompt.Prompt
        Dim repl As New REPL
        While True
            Dim res As PrettyPrompt.PromptResult = prompt.ReadLineAsync().Result
            If res.IsSuccess Then
                If res.Text.ToLower() = "exit" Then
                    Exit While
                ElseIf res.Text.ToLower() = "clear" Then
                    Console.Clear()
                    Continue While
                ElseIf res.Text.ToLower() = "reset" Then
                    repl.Reset()
                    Continue While
                End If
                Dim r As EvaluationResults = repl.Evaluate(res.Text)
                If TypeOf r.Result Is Exception Then
                    AnsiConsole.WriteException(r.Result)
                ElseIf r.Result IsNot Nothing Then
                    AnsiConsole.WriteLine(r.Result.ToString)
                End If
                If r.Diagnostics.Any Then
                    Dim t As New Table
                    t.AddColumn(New TableColumn("Code"))
                    t.AddColumn(New TableColumn("Description"))
                    For Each diag As Diagnostic In r.Diagnostics
                        t.AddRow($"[{GetColourFromSeverity(diag.Severity)}]{diag.Descriptor.Id}[/]", $"[{GetColourFromSeverity(diag.Severity)}]{diag.GetMessage}[/]")
                    Next
                    AnsiConsole.Write(t)
                End If
            End If
        End While
    End Sub

    Private Function GetColourFromSeverity(sev As DiagnosticSeverity) As String
        Select Case sev
            Case DiagnosticSeverity.Error
                Return "red"
            Case DiagnosticSeverity.Warning
                Return "#67885c"
            Case Else
                Return "white"
        End Select
    End Function

End Module