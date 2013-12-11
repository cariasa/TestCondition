Public Class Condicion
    Public IdCondicion As Integer
    Public IdVariable As Integer
    Public IdTipoCondicion As Integer
    Public NombreCondicion As String
    Public Raiz As Integer
    Public Total As Integer
    Public Operando1 As String
    Public Operador As String
    Public Operando2 As String
    Public Sub New(IdCondicion As Integer, IdVariable As Integer, IdTipoCondicion As Integer, NombreCondicion As String,
                    Raiz As Integer, Total As Integer, Operando1 As String, Operador As String, Operando2 As String)
        Me.IdCondicion = IdCondicion
        Me.IdVariable = IdVariable
        Me.IdTipoCondicion = IdTipoCondicion
        Me.NombreCondicion = NombreCondicion
        Me.Raiz = Raiz
        Me.Total = Total
        Me.Operador = Operador
        Me.Operando1 = Operando1
        Me.Operando2 = Operando2
    End Sub
    Public Overrides Function ToString() As String
        If (IdTipoCondicion <> 3) Then
            Return Convert.ToString(IdCondicion) + " " + NombreCondicion + " " + Operando1 + " " + Operador + " " + Operando2 _
                & " " + Convert.ToString(Raiz)
        Else
            Return Convert.ToString(IdCondicion) + " " + NombreCondicion + " " + Operando1 + " " + Operador + " " + Operando2 _
                & " " + Convert.ToString(Raiz) + " " + Convert.ToString(Convert.ToUInt64(Operando1)) + " " + Convert.ToString(Convert.ToUInt64(Operando2))
        End If

    End Function
End Class
