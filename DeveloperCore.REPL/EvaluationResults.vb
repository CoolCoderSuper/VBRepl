Imports Microsoft.CodeAnalysis

Public Class EvaluationResults
    Public Sub New(diagnostics As Diagnostic(), result As Object)
        Me.Diagnostics = diagnostics
        Me.Result = result
    End Sub
    Public ReadOnly Property Diagnostics As Diagnostic()
    Public ReadOnly Property Result As Object
End Class