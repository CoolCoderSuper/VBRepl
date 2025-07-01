Imports System.IO
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
'TODO: Async - Can't parse Await with ParseExecutableStatement
'TODO: Non method things
'TODO: State save should be inserted before all returns
Public Class REPL
    Private _imports As New List(Of ImportsStatementSyntax)
    Private _state As New Dictionary(Of String, Object)
    Private _references As New List(Of MetadataReference)
    Private _trees As New List(Of SyntaxTree)
    Private ReadOnly _nugetCache As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DeveloperCore.REPL", "NuGetCache")

    Public Async Function Evaluate(str As String) As Task(Of EvaluationResults)
        Dim newStatement As StatementSyntax = SyntaxFactory.ParseExecutableStatement(str)
        Dim comp As VisualBasicCompilation = GetCompilation(newStatement)
        Using ms As New MemoryStream
            Dim res As EmitResult = comp.Emit(ms)
            Dim results As Object = Nothing
            If res.Success Then
                If TypeOf newStatement Is ImportsStatementSyntax Then
                    _imports.Add(newStatement)
                End If
                results = Await Run(ms)
            End If
            Return New EvaluationResults(comp.GetDiagnostics.Where(Function(x) x.Severity <> DiagnosticSeverity.Hidden).ToArray, results)
        End Using
    End Function

    Public Function GetCompilation(newStatement As StatementSyntax) As VisualBasicCompilation
        Dim stateName As String = GetRandomString(10)
        Dim unit As CompilationUnitSyntax = SyntaxFactory.ParseCompilationUnit(GetCompilationUnit(newStatement).ToString)
        Dim classNode As ClassBlockSyntax = unit.DescendantNodes.OfType(Of ClassBlockSyntax).First
        unit = unit.ReplaceNode(classNode, classNode.AddMembers(GetMethod(stateName, newStatement))).NormalizeWhitespace
        Return CompilationCreator.GetCompilation(unit).AddReferences(_references).AddSyntaxTrees(_trees)
    End Function

    Private Function GetCompilationUnit(statement As StatementSyntax) As CompilationUnitSyntax
        Dim unit As CompilationUnitSyntax = SyntaxFactory.CompilationUnit().AddMembers(SyntaxCreator.GetModule()).AddImports(_imports.Concat(If(TypeOf statement Is ImportsStatementSyntax, {DirectCast(statement, ImportsStatementSyntax)}, {})).ToArray).NormalizeWhitespace
        Return unit
    End Function

    Private Function GetMethod(stateName As String, statement As StatementSyntax) As MethodBlockSyntax
        Dim method As MethodBlockSyntax = SyntaxCreator.GetMethod(stateName)
        Dim defaultReturn As ReturnStatementSyntax = SyntaxFactory.ReturnStatement(SyntaxFactory.NothingLiteralExpression(SyntaxFactory.ParseToken("Nothing")))
        Dim taskYield As ExpressionStatementSyntax = SyntaxFactory.ExpressionStatement(SyntaxFactory.AwaitExpression(SyntaxFactory.ParseExpression("System.Threading.Tasks.Task.Yield()")))
        If TypeOf statement Is ImportsStatementSyntax Then
            method = method.AddStatements(taskYield, defaultReturn)
        Else
            Dim stateVars As New List(Of VariableDeclaratorSyntax)
            For Each key As String In _state.Keys
                Dim var As VariableDeclaratorSyntax = SyntaxFactory.VariableDeclarator(SyntaxFactory.ModifiedIdentifier(key)).WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName(GetTypeName(key)))).WithInitializer(SyntaxFactory.EqualsValue(SyntaxFactory.ParseExpression($"{stateName}(""{key}"")")))
                stateVars.Add(var)
            Next
            Dim varDeclaration As LocalDeclarationStatementSyntax = SyntaxFactory.LocalDeclarationStatement(New SyntaxTokenList().Add(SyntaxFactory.ParseToken("Dim")), New SeparatedSyntaxList(Of VariableDeclaratorSyntax)().AddRange(stateVars))
            Dim varUpdates As New List(Of StatementSyntax)
            Dim vars As VariableDeclaratorSyntax() = varDeclaration.Declarators.Concat(statement.ChildNodes.OfType(Of VariableDeclaratorSyntax)).ToArray
            For Each var As VariableDeclaratorSyntax In vars
                Dim name As String = var.Names.First.Identifier.Text
                If Not _state.ContainsKey(name) Then
                    varUpdates.Add(SyntaxFactory.ParseExecutableStatement($"{stateName}.Add(""{name}"", {name})"))
                Else
                    varUpdates.Add(SyntaxFactory.ParseExecutableStatement($"{stateName}(""{name}"") = {name}"))
                End If
            Next
            method = method.WithStatements(New SyntaxList(Of StatementSyntax)().AddRange({varDeclaration, statement}).AddRange(varUpdates).Add(taskYield).Add(defaultReturn))
        End If
        Return method
    End Function

    Private Async Function Run(ms As MemoryStream) As Task(Of Object)
        Dim asm As Reflection.Assembly = Reflection.Assembly.Load(ms.ToArray)
        Dim type As Type = asm.GetType("Expression")
        Dim obj As IExpression = Activator.CreateInstance(type)
        Try
            Return Await obj.Evaluate(_state)
        Catch ex As Exception
            Return ex
        End Try
    End Function

    Private Function GetTypeName(key As String) As String
        Dim obj As Object = _state(key)
        If obj Is Nothing Then
            Return "Object"
        Else
            Dim type As Type = obj.GetType
            Return GetTypeName(type)
        End If
    End Function

    Private Function GetTypeName(type As Type) As String
        If type.IsGenericType Then
            Return SyntaxFactory.GenericName(GetGenericNameActual(type), SyntaxFactory.TypeArgumentList(New SeparatedSyntaxList(Of TypeSyntax)().AddRange(type.GenericTypeArguments.Select(Function(x) SyntaxFactory.ParseTypeName(GetTypeName(x)))))).NormalizeWhitespace.ToFullString
        ElseIf type.IsArray Then
            Return SyntaxFactory.ArrayType(SyntaxFactory.ParseTypeName(GetTypeName(type.GetElementType))).NormalizeWhitespace.ToFullString
        Else
            Return type.FullName
        End If
    End Function

    Private Function GetGenericNameActual(type As Type) As String
        Return $"{type.Namespace}.{type.Name.Remove(type.Name.IndexOf("`"))}"
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

    Public Sub AddReference(ref As String)
        If IsValidPath(ref) AndAlso File.Exists(ref) Then
            Select Case Path.GetExtension(ref)
                Case ".dll", ".exe"
                    _references.Add(MetadataReference.CreateFromFile(ref))
                Case ".vb"
                    _trees.Add(SyntaxFactory.ParseSyntaxTree(File.ReadAllText(ref)))
            End Select
        End If
    End Sub

    Private Shared Function IsValidPath(path As String) As Boolean
        Dim fi As FileInfo = Nothing
        Try
            fi = New FileInfo(path)
        Catch __unusedArgumentException1__ As ArgumentException
        Catch __unusedPathTooLongException2__ As PathTooLongException
        Catch __unusedNotSupportedException3__ As NotSupportedException
        End Try
        Return If(fi Is Nothing, False, True)
    End Function
End Class
