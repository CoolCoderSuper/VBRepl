Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Microsoft.CodeAnalysis
Imports Microsoft.CodeAnalysis.VisualBasic
Imports Microsoft.CodeAnalysis.VisualBasic.Syntax

Public Class CompilationCreator

    Public Shared Function GetCompilation(unit As CompilationUnitSyntax) As VisualBasicCompilation
        Return VisualBasicCompilation.Create("repl", {SyntaxFactory.SyntaxTree(unit)}, GetAllReferences, New VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary))
    End Function

    Private Shared Function GetAllReferences() As List(Of MetadataReference)
        Dim referencedAssemblies As Dictionary(Of String, Reflection.Assembly) = Reflection.Assembly.GetEntryAssembly().MyGetReferencedAssembliesRecursive()
        Dim refs As New List(Of MetadataReference)
        For Each a As Reflection.Assembly In referencedAssemblies.Values
            refs.Add(MetadataReference.CreateFromFile(a.Location))
        Next
        refs.Add(MetadataReference.CreateFromFile(Reflection.Assembly.GetEntryAssembly.Location))
        Return refs
    End Function

End Class

''' <summary>
'''     Intent: Get referenced assemblies, either recursively or flat. Not thread safe, if running in a multi
'''     threaded environment must use locks.
''' </summary>
Public Module GetReferencedAssemblies

    Public Class MissingAssembly

        Public Sub New(missingAssemblyName As String, missingAssemblyNameParent As String)
            Me.MissingAssemblyName = missingAssemblyName
            Me.MissingAssemblyNameParent = missingAssemblyNameParent
        End Sub

        Public Property MissingAssemblyName As String
        Public Property MissingAssemblyNameParent As String
    End Class

    Private _dependentAssemblyList As Dictionary(Of String, Assembly)
    Private _missingAssemblyList As List(Of MissingAssembly)

    ''' <summary>
    '''     Intent: Get assemblies referenced by entry assembly. Not recursive.
    ''' </summary>
    <Extension()>
    Public Function MyGetReferencedAssembliesFlat(type As Type) As List(Of String)
        Dim results As AssemblyName() = type.Assembly.GetReferencedAssemblies()
        Return results.[Select](Function(o) o.FullName).OrderBy(Function(o) o).ToList()
    End Function

    ''' <summary>
    '''     Intent: Get assemblies currently dependent on entry assembly. Recursive.
    ''' </summary>
    <Extension()>
    Public Function MyGetReferencedAssembliesRecursive(assembly As Assembly) As Dictionary(Of String, Assembly)
        _dependentAssemblyList = New Dictionary(Of String, Assembly)()
        _missingAssemblyList = New List(Of MissingAssembly)()

        InternalGetDependentAssembliesRecursive(assembly)

        ' Only include assemblies that we wrote ourselves (ignore ones from GAC).
        Dim keysToRemove As List(Of Assembly) = _dependentAssemblyList.Values.Where(Function(o) o.GlobalAssemblyCache = True).ToList()

        For Each k As Assembly In keysToRemove
            _dependentAssemblyList.Remove(k.FullName.MyToName())
        Next

        Return _dependentAssemblyList
    End Function

    ''' <summary>
    '''     Intent: Get missing assemblies.
    ''' </summary>
    <Extension()>
    Public Function MyGetMissingAssembliesRecursive(assembly As Assembly) As List(Of MissingAssembly)
        _dependentAssemblyList = New Dictionary(Of String, Assembly)()
        _missingAssemblyList = New List(Of MissingAssembly)()
        InternalGetDependentAssembliesRecursive(assembly)

        Return _missingAssemblyList
    End Function

    ''' <summary>
    '''     Intent: Internal recursive class to get all dependent assemblies, and all dependent assemblies of
    '''     dependent assemblies, etc.
    ''' </summary>
    Private Sub InternalGetDependentAssembliesRecursive(assembly As Assembly)
        ' Load assemblies with newest versions first. Omitting the ordering results in false positives on
        ' _missingAssemblyList.
        Dim referencedAssemblies As IOrderedEnumerable(Of AssemblyName) = assembly.GetReferencedAssemblies().OrderByDescending(Function(o) o.Version)

        For Each r As AssemblyName In referencedAssemblies
            If String.IsNullOrEmpty(assembly.FullName) Then
                Continue For
            End If

            If _dependentAssemblyList.ContainsKey(r.FullName.MyToName()) = False Then
                Try
                    Dim a As Assembly = Assembly.Load(r.FullName)
                    _dependentAssemblyList(a.FullName.MyToName()) = a
                    InternalGetDependentAssembliesRecursive(a)
                Catch ex As Exception
                    _missingAssemblyList.Add(New MissingAssembly(r.FullName.Split(","c)(0), assembly.FullName.MyToName()))
                End Try
            End If
        Next
    End Sub

    <Extension()>
    Private Function MyToName(fullName As String) As String
        Return fullName.Split(","c)(0)
    End Function

End Module