Public Interface IExpression
    Function Evaluate(state As AdvancedDictionary(Of String, Object)) As Object
End Interface