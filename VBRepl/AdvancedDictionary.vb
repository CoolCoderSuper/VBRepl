
''' <summary>
''' A dictionary that will automatically insert keys when referenced.
''' </summary>
''' <typeparam name="TKey">The type of the key.</typeparam>
''' <typeparam name="TValue">The type of the value.</typeparam>
''' <seealso cref="Dictionary(Of TKey, TValue)" />
Public Class AdvancedDictionary(Of TKey, TValue)
    Inherits Dictionary(Of TKey, TValue)

    ''' <summary>
    ''' Gets or sets the value associated with the specified key.
    ''' </summary>
    Default Public Shadows Property Item(key As TKey) As TValue
        Get
            Return If(ContainsKey(key), MyBase.Item(key), Nothing)
        End Get
        Set(value As TValue)
            If ContainsKey(key) Then
                MyBase.Item(key) = value
            Else
                Add(key, value)
            End If
        End Set
    End Property

End Class