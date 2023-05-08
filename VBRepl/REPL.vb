Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports Spectre.Console
'TODO: Use random parameter name
Public Class REPL
    Private _imports As New List(Of ImportsStatementSyntax)
    Private _state As New Dictionary(Of String, Object)

    Public Sub Evaluate(str As String)
        Dim newStatement As StatementSyntax = SyntaxCreator.ParseStatement(str)
        Dim unit As CompilationUnitSyntax = SyntaxFactory.ParseCompilationUnit(GetUnit(newStatement).ToString)
        Dim method As MethodBlockSyntax = unit.DescendantNodes.OfType(Of MethodBlockSyntax).First
        unit = unit.ReplaceNode(method, UpdateMethod(method)).NormalizeWhitespace
        Dim comp As VisualBasicCompilation = CompilationCreator.GetCompilation(unit)
        Using ms As New MemoryStream
            Dim res As EmitResult = comp.Emit(ms)
            If res.Success Then
                If TypeOf newStatement Is ImportsStatementSyntax Then
                    _imports.Add(newStatement)
                End If
                Run(ms)
            End If
            Dim diagnostics As IEnumerable(Of Diagnostic) = comp.GetDiagnostics.Where(Function(x) x.Severity <> DiagnosticSeverity.Hidden)
            If diagnostics.Any Then
                Dim t As New Table
                t.AddColumn(New TableColumn("Code"))
                t.AddColumn(New TableColumn("Description"))
                For Each diag As Diagnostic In diagnostics
                    t.AddRow($"[{GetColourFromSeverity(diag.Severity)}]{diag.Descriptor.Id}[/]", $"[{GetColourFromSeverity(diag.Severity)}]{diag.GetMessage}[/]")
                Next
                AnsiConsole.Write(t)
            End If
        End Using
    End Sub

    Private Function UpdateMethod(method As MethodBlockSyntax) As MethodBlockSyntax
        Dim stateVars As New List(Of VariableDeclaratorSyntax)
        For Each key As String In _state.Keys
            Dim var As VariableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(SyntaxFactory.ModifiedIdentifier(key)).WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName(GetTypeName(key)))).WithInitializer(SyntaxFactory.EqualsValue(SyntaxFactory.ParseExpression($"state(""{key}"")")))
            stateVars.Add(var)
        Next
        method = method.WithStatements(method.Statements.Insert(0, SyntaxFactory.LocalDeclarationStatement(New SyntaxTokenList().Add(SyntaxFactory.ParseToken("Dim")), New SeparatedSyntaxList(Of VariableDeclaratorSyntax)().AddRange(stateVars)))).NormalizeWhitespace
        Dim retIndex As Integer = method.Statements.IndexOf(method.DescendantNodes.OfType(Of ReturnStatementSyntax).First)
        Dim varUpdates As New List(Of StatementSyntax)
        Dim vars As VariableDeclaratorSyntax() = method.DescendantNodes.OfType(Of VariableDeclaratorSyntax).ToArray
        For Each var As VariableDeclaratorSyntax In vars
            Dim name As String = var.Names.First.Identifier.Text
            If Not _state.ContainsKey(name) Then
                varUpdates.Add(SyntaxFactory.ExpressionStatement(SyntaxFactory.ParseExpression($"state.Add(""{name}"", {name})")))
            Else
                varUpdates.Add(SyntaxFactory.ExpressionStatement(SyntaxFactory.ParseExpression($"state(""{name}"") = {name}")))
            End If
        Next
        method = method.WithStatements(method.Statements.InsertRange(retIndex, varUpdates))
        Return method
    End Function

    Private Function GetTypeName(key As String) As String
        Dim obj As Object = _state(key)
        If obj Is Nothing Then
            Return "Object"
        Else
            Return obj.GetType.FullName
        End If
    End Function

    Private Sub Run(ms As MemoryStream)
        Dim asm As Reflection.Assembly = Reflection.Assembly.Load(ms.ToArray)
        Dim type As Type = asm.GetType("Expression")
        Dim obj As IExpression = Activator.CreateInstance(type)
        Dim result As Object = obj.Evaluate(_state)
        If result IsNot Nothing Then
            AnsiConsole.WriteLine(result.ToString)
        End If
    End Sub

    Private Function GetUnit(statement As StatementSyntax)
        Dim unit As CompilationUnitSyntax = SyntaxCreator.GetCompilationUnit(SyntaxCreator.GetModule.AddMembers(SyntaxCreator.GetMainMethod.AddStatements(If(TypeOf statement Is ImportsStatementSyntax, {}, {statement}).Concat({SyntaxFactory.ReturnStatement(SyntaxFactory.NothingLiteralExpression(SyntaxFactory.ParseToken("Nothing")))}).ToArray))).AddImports(_imports.Concat(If(TypeOf statement Is ImportsStatementSyntax, {DirectCast(statement, ImportsStatementSyntax)}, {})).ToArray).NormalizeWhitespace
        Return unit
    End Function

    Public Sub Reset()
        _imports.Clear()
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

End Class