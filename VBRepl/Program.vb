Imports Microsoft.CodeAnalysis
Imports Spectre.Console
Imports DeveloperCore.REPL
Imports System.Reflection
Public Module Program

    'TODO: Default imports
    'TODO: Config file
    Public Sub Main(args As String())
        Console.Title = "VB REPL"
        Console.WriteLine("VB REPL by CoolCoderSuper")
        Console.WriteLine("Roslyn version: " & GetType(VisualBasic.VisualBasicCompilation).Assembly.GetCustomAttribute(Of AssemblyFileVersionAttribute).Version)
        Console.WriteLine("Type '#exit' to exit, '#clear' to clear the console and '#reset' to reset the REPL")
        Dim prompt As New PrettyPrompt.Prompt
        Dim repl As New REPL
        While True
            Dim res As PrettyPrompt.PromptResult = prompt.ReadLineAsync().Result
            If res.IsSuccess Then
                If res.Text.ToLower() = "#exit" Then
                    Exit While
                ElseIf res.Text.ToLower() = "#clear" Then
                    Console.Clear()
                    Continue While
                ElseIf res.Text.ToLower() = "#reset" Then
                    repl.Reset()
                    Continue While
                ElseIf res.Text.ToLower().StartsWith("#r") Then
                    Dim ref As String = res.Text.Substring(2).Trim
                    If ref.StartsWith("""") AndAlso ref.EndsWith("""") Then
                        ref = ref.Substring(1, ref.Length - 2)
                    End If
                    Try
                        repl.AddReference(ref).Wait()
                    Catch ex As Exception
                        AnsiConsole.WriteException(ex)
                    End Try
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