
Public Class FichaMiembro
    Private ValoresRespuestasUnicas As Dictionary(Of String, Integer)
    Private ValoresRespuestasMultiples As Dictionary(Of String, ArrayList)
    Private _IdMiembro As Integer
    Public Property IdMiembro() As Integer
        Get
            Return _IdMiembro
        End Get
        Set(ByVal value As Integer)
            _IdMiembro = value
        End Set
    End Property
    Public Sub New()
        Me.ValoresRespuestasUnicas = New Dictionary(Of String, Integer)
        Me.ValoresRespuestasMultiples = New Dictionary(Of String, ArrayList)
    End Sub
    Public Sub SetValorRespuestaUnica(ByVal Pregunta As String, ByRef Valor As Object)
        If Not TypeOf Valor Is DBNull Then
            ValoresRespuestasUnicas(Pregunta) = Valor
        Else
            ValoresRespuestasUnicas(Pregunta) = 0
        End If
    End Sub
    Public Sub AddValorRespuestaMultiple(ByVal Pregunta As String, ByVal Valor As Integer)
        If Not ValoresRespuestasMultiples.ContainsKey(Pregunta) Then
            ValoresRespuestasMultiples(Pregunta) = New ArrayList
        End If
        ValoresRespuestasMultiples(Pregunta).Add(Valor)
    End Sub
    Public Sub PrintAllValues()
        Dim SingleValuePair As KeyValuePair(Of String, Integer)
        For Each SingleValuePair In ValoresRespuestasUnicas
            Console.WriteLine(vbTab + vbTab + SingleValuePair.Key + "=" + Convert.ToString(SingleValuePair.Value))
        Next
        Dim MultiValuePair As KeyValuePair(Of String, ArrayList)
        Dim List As ArrayList
        For Each MultiValuePair In ValoresRespuestasMultiples
            Console.WriteLine(vbTab + vbTab + MultiValuePair.Key + ":")
            List = MultiValuePair.Value
            For Each value As Integer In List
                Console.Write(vbTab + vbTab + vbTab + Convert.ToString(value))
            Next
            Console.WriteLine()
        Next
    End Sub
End Class
