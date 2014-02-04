Public Interface FichaInterface
    Function GetValorRespuestaUnica(ByVal Pregunta As String) As Integer
    Function CheckRespuestaUnica(ByVal Pregunta As String) As Boolean
    Function GetValoresRespuestaMultiple(ByVal Pregunta As String) As ArrayList
    Function CheckRespuestaMultiple(ByVal Pregunta As String) As Boolean
End Interface
