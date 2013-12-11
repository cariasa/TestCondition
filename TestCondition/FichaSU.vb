Public Class FichaSU
    Private ValoresRespuestasUnicas As Dictionary(Of String, Integer)
    Private ValoresRespuestasMultiples As Dictionary(Of String, ArrayList)
    Private TipoFicha As Char
    ' V para ficha que sólo tiene información de vivienda
    ' H para ficha que tiene información de vivienda y de hogar
    ' M para ficha que tiene información de vivienda, de hogar y de miembro
    Private IdFicha As Integer
    Private IdVivienda As Integer
    Private IdHogar As Integer
    Private IdMiembro As Integer

    Public Sub New(ByVal IdFicha As Integer, ByVal IdVivienda As Integer, ByVal TipoFicha As Char,
                         Optional ByVal IdHogar As Integer = 0, Optional ByVal IdMiembro As Integer = 0)
        Me.IdFicha = IdFicha
        Me.IdVivienda = IdVivienda
        Me.IdMiembro = IdMiembro
        Me.IdHogar = IdHogar
        Me.TipoFicha = TipoFicha
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
    Public Function GetValorRespuestaUnica(ByVal Pregunta As String) As Integer
        Return ValoresRespuestasUnicas(Pregunta)
    End Function
    Public Function GetValoresRespuestaMultiple(ByVal Pregunta As String) As ArrayList
        Return ValoresRespuestasMultiples(Pregunta)
    End Function
    Public Sub PrintAllValues()
        Dim SingleValuePair As KeyValuePair(Of String, Integer)
        For Each SingleValuePair In ValoresRespuestasUnicas
            Console.WriteLine(SingleValuePair.Key + "=" + Convert.ToString(SingleValuePair.Value))
        Next
        Dim MultiValuePair As KeyValuePair(Of String, ArrayList)
        Dim List As ArrayList
        For Each MultiValuePair In ValoresRespuestasMultiples
            Console.WriteLine(MultiValuePair.Key + ":")
            List = MultiValuePair.Value
            For Each value As Integer In List
                Console.Write(vbTab + Convert.ToString(value))
            Next
            Console.WriteLine()
        Next
    End Sub
End Class
