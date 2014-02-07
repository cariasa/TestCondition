
Public Class FichaVivienda : Implements FichaInterface
    Private ValoresRespuestasUnicas As Dictionary(Of String, Integer)
    Private ValoresRespuestasMultiples As Dictionary(Of String, ArrayList)
    Private HogaresEnVivienda As ArrayList

    Private _IdVivienda As Integer
    Public Property IdVivienda() As Integer
        Get
            Return _IdVivienda
        End Get
        Set(ByVal value As Integer)
            _IdVivienda = value
        End Set
    End Property
    Public Sub New()
        Me.ValoresRespuestasUnicas = New Dictionary(Of String, Integer)
        Me.ValoresRespuestasMultiples = New Dictionary(Of String, ArrayList)
    End Sub
    Public Function GetHogaresEnVivienda() As ArrayList
        Return HogaresEnVivienda
    End Function
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
    Public Sub SetHogares(ByRef Hogares As ArrayList)
        HogaresEnVivienda = Hogares
    End Sub
    Public Function GetValorRespuestaUnica(ByVal Pregunta As String) As Integer Implements FichaInterface.GetValorRespuestaUnica
        Return ValoresRespuestasUnicas(Pregunta)
    End Function
    Public Function CheckRespuestaUnica(ByVal Pregunta As String) As Boolean Implements FichaInterface.CheckRespuestaUnica
        Return ValoresRespuestasUnicas.ContainsKey(Pregunta)
    End Function
    Public Function GetValoresRespuestaMultiple(ByVal Pregunta As String) As ArrayList Implements FichaInterface.GetValoresRespuestaMultiple
        Return ValoresRespuestasMultiples(Pregunta)
    End Function
    Public Function CheckRespuestaMultiple(ByVal Pregunta As String) As Boolean Implements FichaInterface.CheckRespuestaMultiple
        Return ValoresRespuestasMultiples.ContainsKey(Pregunta)
    End Function

    Public Function GetUbicacionGeografica() As VariablesDesagregadas
        Dim Ubicacion As New VariablesDesagregadas
        Ubicacion.IdDepartamento = ValoresRespuestasUnicas("V01")
        Ubicacion.IdMunicipio = ValoresRespuestasUnicas("V02")
        Ubicacion.IdAldea = ValoresRespuestasUnicas("V03")
        Ubicacion.IdCaserio = ValoresRespuestasUnicas("V04")
        Ubicacion.IdBarrio = ValoresRespuestasUnicas("V05")
        Ubicacion.IdRegion = GetValorRespuestaUnica("V06")
        Return Ubicacion
    End Function
    Public Function GetDepartamento() As VariableDepartamento
        Dim VarDepto As New VariableDepartamento
        VarDepto.Departamento = ValoresRespuestasUnicas("V01")
        Return VarDepto
    End Function
    Public Function GetDepartamentoMunicipio() As VariableDepartamentoMunicipio
        Dim VarDeptoMuni As New VariableDepartamentoMunicipio
        VarDeptoMuni.Departamento = ValoresRespuestasUnicas("V01")
        VarDeptoMuni.Municipio = ValoresRespuestasUnicas("V02")
        Return VarDeptoMuni
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
