Imports System.IO
Imports DeveloperCore.REPL
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
'TODO: Multiple statements
Public Class REPL
    Private _imports As New List(Of ImportsStatementSyntax)
    Private _state As New Dictionary(Of String, Object)

    Public Function Evaluate(str As String) As EvaluationResults
        Dim newStatement As StatementSyntax = SyntaxCreator.ParseStatement(str)
        Dim stateName As String = GetRandomString(10)
        Dim unit As CompilationUnitSyntax = SyntaxFactory.ParseCompilationUnit(GetUnit(newStatement, stateName).ToString)
        Dim method As MethodBlockSyntax = unit.DescendantNodes.OfType(Of MethodBlockSyntax).First
        unit = unit.ReplaceNode(method, UpdateMethod(method, stateName)).NormalizeWhitespace
        Dim comp As VisualBasicCompilation = CompilationCreator.GetCompilation(unit)
        Using ms As New MemoryStream
            Dim res As EmitResult = comp.Emit(ms)
            Dim results As Object
            If res.Success Then
                If TypeOf newStatement Is ImportsStatementSyntax Then
                    _imports.Add(newStatement)
                End If
                results = Run(ms)
            End If
            Return New EvaluationResults(comp.GetDiagnostics.Where(Function(x) x.Severity <> DiagnosticSeverity.Hidden).ToArray, results)
        End Using
    End Function

    Private Function UpdateMethod(method As MethodBlockSyntax, stateName As String) As MethodBlockSyntax
        Dim stateVars As New List(Of VariableDeclaratorSyntax)
        For Each key As String In _state.Keys
            Dim var As VariableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(SyntaxFactory.ModifiedIdentifier(key)).WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName(GetTypeName(key)))).WithInitializer(SyntaxFactory.EqualsValue(SyntaxFactory.ParseExpression($"{stateName}(""{key}"")")))
            stateVars.Add(var)
        Next
        method = method.WithStatements(method.Statements.Insert(0, SyntaxFactory.LocalDeclarationStatement(New SyntaxTokenList().Add(SyntaxFactory.ParseToken("Dim")), New SeparatedSyntaxList(Of VariableDeclaratorSyntax)().AddRange(stateVars)))).NormalizeWhitespace
        Dim retIndex As Integer = method.Statements.IndexOf(method.DescendantNodes.OfType(Of ReturnStatementSyntax).First)
        Dim varUpdates As New List(Of StatementSyntax)
        Dim vars As VariableDeclaratorSyntax() = method.DescendantNodes.OfType(Of VariableDeclaratorSyntax).ToArray
        For Each var As VariableDeclaratorSyntax In vars
            Dim name As String = var.Names.First.Identifier.Text
            If Not _state.ContainsKey(name) Then
                varUpdates.Add(SyntaxFactory.ExpressionStatement(SyntaxFactory.ParseExpression($"{stateName}.Add(""{name}"", {name})")))
            Else
                varUpdates.Add(SyntaxFactory.ExpressionStatement(SyntaxFactory.ParseExpression($"{stateName}(""{name}"") = {name}")))
            End If
        Next
        method = method.WithStatements(method.Statements.InsertRange(retIndex, varUpdates))
        Return method
    End Function

    Private Function GetTypeName(key As String) As String
        Dim obj As Object = _state(key)
        Return If(obj Is Nothing, "Object", obj.GetType.FullName)
    End Function

    Private Function Run(ms As MemoryStream) As Object
        Dim asm As Reflection.Assembly = Reflection.Assembly.Load(ms.ToArray)
        Dim type As Type = asm.GetType("Expression")
        Dim obj As IExpression = Activator.CreateInstance(type)
        Try
            Return obj.Evaluate(_state)
        Catch ex As Exception
            Return ex
        End Try
    End Function

    Private Function GetUnit(statement As StatementSyntax, param As String)
        Dim unit As CompilationUnitSyntax = SyntaxCreator.GetCompilationUnit(SyntaxCreator.GetModule.AddMembers(SyntaxCreator.GetMainMethod(param).AddStatements(If(TypeOf statement Is ImportsStatementSyntax, {}, {statement}).Concat({SyntaxFactory.ReturnStatement(SyntaxFactory.NothingLiteralExpression(SyntaxFactory.ParseToken("Nothing")))}).ToArray))).AddImports(_imports.Concat(If(TypeOf statement Is ImportsStatementSyntax, {DirectCast(statement, ImportsStatementSyntax)}, {})).ToArray).NormalizeWhitespace
        Return unit
    End Function

    Public Sub Reset()
        _state.Clear()
        _imports.Clear()
    End Sub

    Private Shared Function GetRandomString(intLength As Integer) As String
        Dim letters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Dim numbers As String = "1234567890"
        Dim rType As New Random
        Dim rLetter As New Random
        Dim rNumber As New Random
        Dim result As String = String.Empty
        Dim c As Integer = 0
        While c < intLength
            If rType.Next(1, 2) = 1 Then
                result &= letters(rLetter.Next(0, 25))
            Else
                result &= numbers(rNumber.Next(0, 9))
            End If
            c += 1
        End While
        Return result
    End Function

End Class