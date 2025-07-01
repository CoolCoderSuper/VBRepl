Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Class SyntaxCreator

    Public Shared Function GetModule() As ClassBlockSyntax
        Return SyntaxFactory.ClassBlock(SyntaxFactory.ClassStatement(SyntaxFactory.ParseToken("Expression"))).AddImplements(SyntaxFactory.ImplementsStatement(SyntaxFactory.ParseTypeName("DeveloperCore.REPL.IExpression")))
    End Function

    Public Shared Function GetMethod(param As String) As MethodBlockSyntax
        Return SyntaxFactory.FunctionBlock(SyntaxFactory.FunctionStatement("Evaluate").AddModifiers(SyntaxFactory.Token(SyntaxKind.AsyncKeyword)).WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName("System.Threading.Tasks.Task(Of Object)"))).AddImplementsClauseInterfaceMembers(SyntaxFactory.QualifiedName(SyntaxFactory.ParseTypeName("DeveloperCore.REPL.IExpression"), SyntaxFactory.ParseName("Evaluate"))).AddParameterListParameters(SyntaxFactory.Parameter(SyntaxFactory.ModifiedIdentifier(param)).WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName("System.Collections.Generic.Dictionary(Of String, Object)")))))
    End Function

End Class
