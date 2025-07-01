Public Interface IExpression
    Function Evaluate(state As Dictionary(Of String, Object)) As Task(Of Object)
End Interface
