Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Class SyntaxCreator

    Public Shared Function GetModule() As ModuleBlockSyntax
        Return SyntaxFactory.ModuleBlock(SyntaxFactory.ModuleStatement(SyntaxFactory.ParseToken("Program")))
    End Function

    Public Shared Function GetMainMethod() As MethodBlockSyntax
        Return SyntaxFactory.SubBlock(SyntaxFactory.SubStatement("Main"))
    End Function

    Public Shared Function ParseStatement(str As String) As StatementSyntax
        Return SyntaxFactory.ParseExecutableStatement(str)
    End Function

    Public Shared Function GetCompilationUnit(node As SyntaxNode) As CompilationUnitSyntax
        Return SyntaxFactory.CompilationUnit().AddMembers(node)
    End Function

End Class