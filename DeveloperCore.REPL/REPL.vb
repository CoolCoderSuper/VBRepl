Imports System.IO
Imports System.Threading
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.Emit
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax
Imports NuGet.Common
Imports NuGet.Configuration
Imports NuGet.PackageManagement
Imports NuGet.Packaging
Imports NuGet.Packaging.Core
Imports NuGet.ProjectManagement
Imports NuGet.Protocol.Core.Types
Imports NuGet.Resolver
'TODO: Multiple statements
Public Class REPL
    Private _imports As New List(Of ImportsStatementSyntax)
    Private _state As New Dictionary(Of String, Object)
    Private _references As New List(Of MetadataReference)
    Private _trees As New List(Of SyntaxTree)
    Private ReadOnly _nugetCache As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DeveloperCore.REPL", "NuGetCache")

    Public Function Evaluate(str As String) As EvaluationResults
        Dim newStatement As StatementSyntax = SyntaxCreator.ParseStatement(str)
        Dim comp As VisualBasicCompilation = GetCompilation(str)
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

    Public Function GetCompilation(str As String) As VisualBasicCompilation
        Dim newStatement As StatementSyntax = SyntaxCreator.ParseStatement(str)
        Dim stateName As String = GetRandomString(10)
        Dim unit As CompilationUnitSyntax = SyntaxFactory.ParseCompilationUnit(GetUnit(newStatement, stateName).ToString)
        Dim method As MethodBlockSyntax = unit.DescendantNodes.OfType(Of MethodBlockSyntax).First
        unit = unit.ReplaceNode(method, UpdateMethod(method, stateName)).NormalizeWhitespace
        Return CompilationCreator.GetCompilation(unit).AddReferences(_references).AddSyntaxTrees(_trees)
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

    Private Function GetTypeName(key As String) As String
        Dim obj As Object = _state(key)
        Return If(obj Is Nothing, "Object", obj.GetType.FullName)
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

    Public Async Function AddReference(ref As String) As Task
        If IsValidPath(ref) AndAlso File.Exists(ref) Then
            Select Case Path.GetExtension(ref)
                Case ".dll", ".exe"
                    _references.Add(MetadataReference.CreateFromFile(ref))
                Case ".vb"
                    _trees.Add(SyntaxFactory.ParseSyntaxTree(File.ReadAllText(ref)))
            End Select
        ElseIf ref.StartsWith("nuget:") Then
            'TODO: Add package selector
            Directory.CreateDirectory(_nugetCache)
            Dim parts As String() = ref.Split(":"c)
            Dim name As String = parts(1).Trim
            Dim version As String = If(parts.Length > 2, parts(2).Trim, "")
            Dim providers As New List(Of Lazy(Of INuGetResourceProvider))
            providers.AddRange(Repository.Provider.GetCoreV3())
            Dim sourceRepository As New SourceRepository(New PackageSource("https://api.nuget.org/v3/index.json"), providers)
            Dim log As New Logger
            Dim nugetSettings As Settings = Settings.LoadDefaultSettings(_nugetCache)
            Dim sourceProvider As New PackageSourceProvider(nugetSettings)
            Dim repositoryProvider As New SourceRepositoryProvider(sourceProvider, providers)
            Dim project As New FolderNuGetProject(_nugetCache)
            Dim manager As New NuGetPackageManager(repositoryProvider, nugetSettings, _nugetCache) With {.PackagesFolderNuGetProject = project}
            Dim searchSource As PackageSearchResource = sourceRepository.GetResource(Of PackageSearchResource)()
            Dim package As IPackageSearchMetadata = (Await searchSource.SearchAsync(name, New SearchFilter(True) With {.IncludeDelisted = False, .SupportedFrameworks = {"netstandard2.0"}}, 0, 10, log, CancellationToken.None)).FirstOrDefault
            If package Is Nothing Then Throw New Exception($"Package {name} not found")
            Dim prerelease As Boolean = True
            Dim unListed As Boolean = False
            Dim resolutionContext As New ResolutionContext(DependencyBehavior.Lowest, prerelease, unListed, VersionConstraints.None)
            Dim context As New ProjectContext
            Dim identity As New PackageIdentity(package.Identity.Id, package.Identity.Version)
            Await manager.InstallPackageAsync(project, identity, resolutionContext, context, sourceRepository, {}, CancellationToken.None)
        End If
    End Function

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

Public Class Logger
    Implements ILogger
    Private logs As New List(Of String)()

    Public Sub LogDebug(data As String) Implements ILogger.LogDebug
        logs.Add(data)
    End Sub

    Public Sub LogVerbose(data As String) Implements ILogger.LogVerbose
        logs.Add(data)
    End Sub

    Public Sub LogInformation(data As String) Implements ILogger.LogInformation
        logs.Add(data)
    End Sub

    Public Sub LogMinimal(data As String) Implements ILogger.LogMinimal
        logs.Add(data)
    End Sub

    Public Sub LogWarning(data As String) Implements ILogger.LogWarning
        logs.Add(data)
    End Sub

    Public Sub LogError(data As String) Implements ILogger.LogError
        logs.Add(data)
    End Sub

    Public Sub LogInformationSummary(data As String) Implements ILogger.LogInformationSummary
        logs.Add(data)
    End Sub

    Public Sub Log(level As LogLevel, data As String) Implements ILogger.Log
        logs.Add(data)
    End Sub

    Public Sub Log(message As ILogMessage) Implements ILogger.Log
        logs.Add(message.Message)
    End Sub

    Public Function LogAsync(level As LogLevel, data As String) As Task Implements ILogger.LogAsync
        logs.Add(data)
        Return Task.CompletedTask
    End Function

    Public Function LogAsync(message As ILogMessage) As Task Implements ILogger.LogAsync
        logs.Add(message.Message)
        Return Task.CompletedTask
    End Function
End Class

Public Class ProjectContext
    Implements INuGetProjectContext
    Private logs As New List(Of String)()

    Public Function GetLogs() As List(Of String)
        Return logs
    End Function

    Public Sub Log(level As MessageLevel, message As String, ParamArray args As Object()) Implements INuGetProjectContext.Log
        Dim formattedMessage = String.Format(message, args)
        logs.Add(formattedMessage)
    End Sub

    Public Function ResolveFileConflict(message As String) As FileConflictAction Implements INuGetProjectContext.ResolveFileConflict
        logs.Add(message)
        Return FileConflictAction.Ignore
    End Function

    Public Property PackageExtractionContext As PackageExtractionContext Implements INuGetProjectContext.PackageExtractionContext

    Public ReadOnly Property ExecutionContext As NuGet.ProjectManagement.ExecutionContext Implements INuGetProjectContext.ExecutionContext

    Public Property OriginalPackagesConfig As XDocument Implements INuGetProjectContext.OriginalPackagesConfig

    Public Property SourceControlManagerProvider As ISourceControlManagerProvider Implements INuGetProjectContext.SourceControlManagerProvider

    Public Sub ReportError(message As String) Implements INuGetProjectContext.ReportError
        logs.Add(message)
    End Sub

    Public Sub Log(message As ILogMessage) Implements INuGetProjectContext.Log
        Log(message.Level, message.Message)
    End Sub

    Public Sub ReportError(message As ILogMessage) Implements INuGetProjectContext.ReportError
        ReportError(message.Message)
    End Sub

    Public Property ActionType As NuGetActionType Implements INuGetProjectContext.ActionType

    Public Property OperationId As Guid Implements INuGetProjectContext.OperationId
End Class