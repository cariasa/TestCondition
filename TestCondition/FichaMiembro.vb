﻿
Public Class FichaMiembro : Implements FichaInterface
    Private ValoresRespuestasUnicas As Dictionary(Of String, Integer)
    Private ValoresRespuestasMultiples As Dictionary(Of String, ArrayList)
    Private HogarReference As FichaHogar
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
    Public Sub SetHogarReference(ByRef HogarReference As FichaHogar)
        Me.HogarReference = HogarReference
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
    Public Function GetSexo() As VariableSexo
        Dim VarSexo As VariableSexo
        If ValoresRespuestasUnicas.ContainsKey("P9") Then
            VarSexo = New VariableSexo
            VarSexo.Sexo = ValoresRespuestasUnicas("P9")
        Else
            VarSexo = Nothing
        End If
        Return VarSexo
    End Function
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
