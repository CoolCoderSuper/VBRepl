Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Class SyntaxCreator

    Public Shared Function GetModule() As ClassBlockSyntax
        Return SyntaxFactory.ClassBlock(SyntaxFactory.ClassStatement(SyntaxFactory.ParseToken("Expression"))).AddImplements(SyntaxFactory.ImplementsStatement(SyntaxFactory.ParseTypeName("VBRepl.IExpression")))
    End Function

    Public Shared Function GetMainMethod() As MethodBlockSyntax
        Return SyntaxFactory.FunctionBlock(SyntaxFactory.FunctionStatement("Evaluate").AddImplementsClauseInterfaceMembers(SyntaxFactory.QualifiedName(SyntaxFactory.ParseTypeName("VBRepl.IExpression"), SyntaxFactory.ParseName("Evaluate"))).AddParameterListParameters(SyntaxFactory.Parameter(SyntaxFactory.ModifiedIdentifier("state")).WithAsClause(SyntaxFactory.SimpleAsClause(SyntaxFactory.ParseTypeName("System.Collections.Generic.Dictionary(Of String, Object)")))))
    End Function

    Public Shared Function ParseStatement(str As String) As StatementSyntax
        Return SyntaxFactory.ParseExecutableStatement(str)
    End Function

    Public Shared Function GetCompilationUnit(node As SyntaxNode) As CompilationUnitSyntax
        Return SyntaxFactory.CompilationUnit().AddMembers(node)
    End Function

End Class