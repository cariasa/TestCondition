Public Class FormulaIndicador
    Public IdIndicador As Integer
    Public IdIndicadoresEvaluacionPorPrograma As Integer
    Public IdVariableNumerador As Integer
    Public Numerador As String
    Public IdVariableDenominador As Integer
    Public Denominador As String
    Public UsaVariableMacroNumerador, UsaVariableMacroDenominador As Integer
    Public Sub New(ByVal IdIndicador As Integer, ByVal IdIndicadoresEvaluacionPorPrograma As Integer,
                   ByVal IdVariableNumerador As Integer, ByVal Numerador As String, ByVal IdVariableDenominador As Integer,
                   ByVal Denominador As String, ByVal UsaVariableMacroNumerador As Integer,
                   ByVal UsaVariableMacroDenominador As Integer)
        Me.IdIndicador = IdIndicador
        Me.IdIndicadoresEvaluacionPorPrograma = IdIndicadoresEvaluacionPorPrograma
        Me.IdVariableDenominador = IdVariableDenominador
        Me.IdVariableNumerador = IdVariableNumerador
        Me.Numerador = Numerador
        Me.Denominador = Denominador
        Me.UsaVariableMacroDenominador = UsaVariableMacroDenominador
        Me.UsaVariableMacroNumerador = UsaVariableMacroNumerador
    End Sub
    Public Overrides Function ToString() As String
        Return Numerador + "/" + Denominador
    End Function
End Class
