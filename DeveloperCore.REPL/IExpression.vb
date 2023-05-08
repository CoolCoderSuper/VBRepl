Public Interface IExpression
    Function Evaluate(state As Dictionary(Of String, Object)) As Object
End Interface