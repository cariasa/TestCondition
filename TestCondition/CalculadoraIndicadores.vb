Imports System.Data.SqlClient
Imports System.Data

Public Class CalculadoraIndicadores
    Private ConnectionString As String
    Private QueryCondiciones As String = "SELECT * FROM Condiciones WHERE IdVariable = @IdVariable"
    Private QueryValoresOficiales As String = "SELECT TOP 1 Valor FROM VariablesMacro WHERE IdVariable = @IdVariable ORDER BY FechaCreacion DESC"
    Private QueryValoresPrograma As String = "SELECT TOP 1 Valor FROM VariablesPrograma WHERE IdVariable = @IdVariable ORDER BY FechaCreacion DESC"
    Private QueryRaiz As String = "SELECT * FROM Condiciones WHERE Raiz = 1 AND IdVariable = @IdVariable"
    Private QueryFormulas As String = "SELECT  " _
                                & "	IEP.IdIndicador,  " _
                                & "	IEP.IdIndicadoresEvaluacionPorPrograma,  " _
                                & "	FI.IdVariableNumerador,  " _
                                & "	V1.NombreVariable as Numerador,  " _
                                & "	FI.IdVariableDenominador,  " _
                                & "	V2.NombreVariable as Denominador,  " _
                                & "	FI.UsaVariableMacroNumerador,  " _
                                & "	FI.UsaVariableMacroDenominador  " _
                                & "FROM  " _
                                & "	IndicadoresEvaluacionPorPrograma IEP  " _
                                & "	JOIN FormulaIndicador FI ON IEP.IdIndicador=FI.IdIndicador  " _
                                & "	JOIN Variables V1 ON FI.IdVariableNumerador = V1.IdVariable  " _
                                & "	JOIN Variables V2 ON FI.IdVariableDenominador = V2.IdVariable  " _
                                & "WHERE  " _
                                & "	IEP.IdPrograma = @IdPrograma AND " _
                                & "	FI.BasadoEnFSU = @BasadoEnFSU"
    Private QueryRecuperaFichaIE As String = "select		" _
                                & "	PPI.IdAmigable,	" _
                                & "	PPI.IdTipoDePregunta,	" _
                                & "	OPI.Valor	" _
                                & " from	" _
                                & "	EncabezadoRespuesta ER	" _
                                & "	join RespuestasInstrumento RI on ER.IdEncabezadoRespuesta = RI.IdEncabezadoRespuesta	" _
                                & "	join PreguntasPorInstrumento PPI on RI.IdPreguntaPorInstrumento = PPI.IdPreguntaPorInstrumento	" _
                                & "	join OpcionesRespuestaInstrumento ORI on RI.IdRespuestaInstrumento = ORI.IdRespuestaInstrumento	" _
                                & "	join OpcionesPreguntaPorInstrumento OPI on ORI.IdOpcionPreguntaPorInstrumento = OPI.IdOpcionPreguntaPorInstrumento	" _
                                & "where	" _
                                & "	ER.IdEncabezadoRespuesta = @IdEncabezadoRespuesta" _
                                & " order by PPI.IdAmigable"

    Private InsertValoresDepartameto As String = "INSERT INTO ValoresDepartamento " _
                               & "(IdAplicacionInstrumento, IdIndicadorEvaluacionPorPrograma, IdDepartamento, Valor, CreadoPor) " _
                               & "VALUES (@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdDepartamento,@Valor,@CreadoPor)"
    Private InsertValoresMunicipio As String = "INSERT INTO ValoresMunicipio " _
                               & "(IdAplicacionInstrumento, IdIndicadorEvaluacionPorPrograma, IdDepartamento, IdMunicipio,Valor, CreadoPor) " _
                               & "VALUES (@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdDepartamento, @IdMunicipio, @Valor,@CreadoPor)"
    Private InsertValoresSexo As String = "INSERT INTO ValoresSexo " _
                               & "(IdAplicacionInstrumento, IdIndicadorEvaluacionPorPrograma, IdSexo, Valor, CreadoPor) " _
                               & "VALUES (@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdSexo,@Valor,@CreadoPor)"
    Private QueryUsaFSU As String = "SELECT UsaFSU FROM AplicacionInstrumento WHERE IdAplicacionInstrumento=@IdLevantamiento"
    Private IdPrograma As Integer
    Private IdLevantamiento As Integer
    Private UsaFSU As Boolean
    Public Sub New(ByVal ConnectionString As String, ByVal IdPrograma As Integer, ByVal IdLevantamiento As Integer)
        Me.ConnectionString = ConnectionString
        Me.IdPrograma = IdPrograma
        Me.IdLevantamiento = IdLevantamiento

        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As SqlCommand = New SqlCommand(QueryUsaFSU, SqlConn)
        Command.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
        Dim Reader As SqlDataReader = Command.ExecuteReader
        If Reader.Read Then
            UsaFSU = Convert.ToBoolean(Reader("UsaFSU"))
        Else
            Throw New Exception("No Hay Levantamiento con ID= " + Convert.ToString(IdLevantamiento))
        End If
        SqlConn.Close()
    End Sub
    Sub ComputeIEAcumulators(ByVal Ficha As FichaSU, ByVal VarTreePair As KeyValuePair(Of String, ConditionTreeNode), ByRef VariableAcum As Dictionary(Of String, Double))
        If VarTreePair.Value.Evaluate(Ficha) Then
            If VariableAcum.ContainsKey(VarTreePair.Key) Then
                VariableAcum(VarTreePair.Key) = VariableAcum(VarTreePair.Key) + 1
            Else
                VariableAcum(VarTreePair.Key) = 1
            End If
        End If
    End Sub
    Sub ComputeFSUAcumulators(ByVal Ficha As Ficha, ByVal VarTreePair As KeyValuePair(Of String, ConditionTreeNode), ByRef VariableAcum As Dictionary(Of String, Double), ByRef VariableDeptoAcum As Dictionary(Of VariableDepartamento, Double), ByRef VariableMuniAcum As Dictionary(Of VariableDepartamentoMunicipio, Double), ByRef VariableSexoAcum As Dictionary(Of VariableSexo, Double))
        Dim VarDepto As VariableDepartamento = Ficha.GetDepartamento
        VarDepto.Variable = VarTreePair.Key
        Dim VarDeptoMuni As VariableDepartamentoMunicipio = Ficha.GetDepartamentoMunicipio
        VarDeptoMuni.Variable = VarTreePair.Key

        If VarTreePair.Value.Level = "V" Then
            If VarTreePair.Value.Evaluate(Ficha.GetFichaVivienda) Then
                'Desagregacion por Ubicacion Geográfica
                If VariableDeptoAcum.ContainsKey(VarDepto) Then
                    VariableDeptoAcum(VarDepto) = VariableDeptoAcum(VarDepto) + 1
                Else
                    VariableDeptoAcum(VarDepto) = 1
                End If
                If VariableMuniAcum.ContainsKey(VarDeptoMuni) Then
                    VariableMuniAcum(VarDeptoMuni) = VariableMuniAcum(VarDeptoMuni) + 1
                Else
                    VariableMuniAcum(VarDeptoMuni) = 1
                End If
                '=========================================================
                If VariableAcum.ContainsKey(VarTreePair.Key) Then
                    VariableAcum(VarTreePair.Key) = VariableAcum(VarTreePair.Key) + 1
                Else
                    VariableAcum(VarTreePair.Key) = 1
                End If
            End If
        ElseIf VarTreePair.Value.Level = "H" Then
            Dim Hogares As ArrayList = Ficha.GetFichaVivienda.GetHogaresEnVivienda
            For Each Hogar As FichaHogar In Hogares
                If VarTreePair.Value.Evaluate(Hogar) Then
                    'Desagregacion por Ubicacion Geográfica
                    If VariableDeptoAcum.ContainsKey(VarDepto) Then
                        VariableDeptoAcum(VarDepto) = VariableDeptoAcum(VarDepto) + 1
                    Else
                        VariableDeptoAcum(VarDepto) = 1
                    End If
                    If VariableMuniAcum.ContainsKey(VarDeptoMuni) Then
                        VariableMuniAcum(VarDeptoMuni) = VariableMuniAcum(VarDeptoMuni) + 1
                    Else
                        VariableMuniAcum(VarDeptoMuni) = 1
                    End If
                    '=========================================================
                    If VariableAcum.ContainsKey(VarTreePair.Key) Then
                        VariableAcum(VarTreePair.Key) = VariableAcum(VarTreePair.Key) + 1
                    Else
                        VariableAcum(VarTreePair.Key) = 1
                    End If
                End If
            Next
        Else
            Dim Hogares As ArrayList = Ficha.GetFichaVivienda.GetHogaresEnVivienda
            For Each Hogar As FichaHogar In Hogares
                Dim Miembros As ArrayList = Hogar.GetMiembrosEnHogar
                For Each Miembro As FichaMiembro In Miembros
                    Dim VarSexo As VariableSexo
                    VarSexo = Miembro.GetSexo
                    VarSexo.Variable = VarTreePair.Key
                    If VarTreePair.Value.Evaluate(Miembro) Then
                        'Desagregacion por Ubicacion Geográfica
                        If VariableDeptoAcum.ContainsKey(VarDepto) Then
                            VariableDeptoAcum(VarDepto) = VariableDeptoAcum(VarDepto) + 1
                        Else
                            VariableDeptoAcum(VarDepto) = 1
                        End If
                        If VariableMuniAcum.ContainsKey(VarDeptoMuni) Then
                            VariableMuniAcum(VarDeptoMuni) = VariableMuniAcum(VarDeptoMuni) + 1
                        Else
                            VariableMuniAcum(VarDeptoMuni) = 1
                        End If
                        '=========================================================
                        'Desagregacion por Sexo
                        If VariableSexoAcum.ContainsKey(VarSexo) Then
                            VariableSexoAcum(VarSexo) = VariableSexoAcum(VarSexo) + 1
                        Else
                            VariableSexoAcum(VarSexo) = 1
                        End If
                        '=========================================================
                        If VariableAcum.ContainsKey(VarTreePair.Key) Then
                            VariableAcum(VarTreePair.Key) = VariableAcum(VarTreePair.Key) + 1
                        Else
                            VariableAcum(VarTreePair.Key) = 1
                        End If
                    End If
                Next
            Next

        End If

    End Sub
    'OficialPrograma = 1 significa calcular Valor Oficial, sino Programa
    Function GetValor(ByVal IdVariable As Integer, ByVal OficialPrograma As Integer) As Double
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As SqlCommand
        If OficialPrograma = 1 Then
            Command = New SqlCommand(QueryValoresOficiales, SqlConn)
        Else
            Command = New SqlCommand(QueryValoresPrograma, SqlConn)
        End If
        Command.Parameters.AddWithValue("@IdVariable", IdVariable)
        Dim Reader As SqlDataReader = Command.ExecuteReader
        If Reader.Read Then
            Return Reader("Valor")
        Else
            Return -1.0
        End If
    End Function
    Function ComputeIndicadorOficialPrograma(ByVal Formula As FormulaIndicador, ByVal OficialPrograma As Integer) As Double
        Dim num, den, res As Double
        num = GetValor(Formula.IdVariableNumerador, OficialPrograma)
        If num = -1 Then
            Return -1
        End If
        den = GetValor(Formula.IdVariableDenominador, OficialPrograma)
        If den = -1 Then
            Return -1
        End If
        res = num / den
        Return res
    End Function
    Sub StoreIndicador(ByVal Formula As FormulaIndicador, ByVal res As Double, ByVal CreadoPor As String, ByVal OficialPrograma As Integer)
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("InsertarValoresIndicadores", SqlConn)
        Command.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
        Command.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
        Command.Parameters.AddWithValue("@ValorIndicador", res)
        Command.Parameters.AddWithValue("@FechaCalculo", Date.Now)
        Command.Parameters.AddWithValue("@CreadoPor", CreadoPor)
        If OficialPrograma = 1 Then
            Command.Parameters.AddWithValue("@IndicadorCalculado", False)
            Command.Parameters.AddWithValue("@IndicadorOficial", True)
            Command.Parameters.AddWithValue("@IndicadorPrograma", False)
        Else
            Command.Parameters.AddWithValue("@IndicadorCalculado", False)
            Command.Parameters.AddWithValue("@IndicadorOficial", False)
            Command.Parameters.AddWithValue("@IndicadorPrograma", True)
        End If
        Command.CommandType = CommandType.StoredProcedure
        Command.ExecuteNonQuery()
        SqlConn.Close()
    End Sub
    Sub ComputeIndicatorsAndInsertOficialProgram(ByVal Formula As FormulaIndicador, ByVal CreadoPor As String)
        Dim value As Double
        value = ComputeIndicadorOficialPrograma(Formula, 1)
        If value <> -1 Then
            StoreIndicador(Formula, value, CreadoPor, 1)
        End If
        value = ComputeIndicadorOficialPrograma(Formula, 2)
        If value <> -1 Then
            StoreIndicador(Formula, value, CreadoPor, 2)
        End If
    End Sub
    Sub ComputeIndicatorsAndInsert(ByVal Formulas As ArrayList, ByVal CreadoPor As String, ByRef VariableAcum As Dictionary(Of String, Double))
        Dim SqlConn As SqlConnection = GetConnection()
        ' NO se está desagregando en este nivel, pues no se sabe cuál es el campo de ubicación geográfica (Departamento, Municipio), ya que la encuesta fue creada
        ' dinámicamente
        ' Se podría sacar de la FSU si estuviera asociada, pero no lo está ya que se especifica que no usa FSU
        ' Para poder desagregar, se puede usar la tabla de pivoteo, recuperando los datos completos.
        For Each Formula As FormulaIndicador In Formulas
            ComputeIndicatorsAndInsertOficialProgram(Formula, CreadoPor)
            Dim Command As New SqlCommand("InsertarValoresIndicadores", SqlConn)
            Dim num As Double
            Dim den As Double
            Dim res As Double
            If VariableAcum.ContainsKey(Formula.Numerador) Then
                num = VariableAcum(Formula.Numerador)
            Else
                num = 0
            End If
            If VariableAcum.ContainsKey(Formula.Denominador) Then
                den = VariableAcum(Formula.Denominador)
            Else
                den = 0
            End If
            If den = 0 Then
                res = -1 ' Indica que no se calculó el indicador. Tal vez no habían datos para calcular
            Else
                res = num / den
            End If
            Command.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
            Command.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
            Command.Parameters.AddWithValue("@ValorIndicador", res)
            Command.Parameters.AddWithValue("@FechaCalculo", Date.Now)
            Command.Parameters.AddWithValue("@CreadoPor", CreadoPor)
            Command.Parameters.AddWithValue("@IndicadorCalculado", True)
            Command.Parameters.AddWithValue("@IndicadorOficial", False)
            Command.Parameters.AddWithValue("@IndicadorPrograma", False)
            Command.CommandType = CommandType.StoredProcedure
            Command.ExecuteNonQuery()
        Next
        SqlConn.Close()
    End Sub
    Sub ComputeIndicatorsAndInsert(ByVal Formulas As ArrayList, ByVal CreadoPor As String, ByRef VariableAcum As Dictionary(Of String, Double), ByRef VariableDeptoAcum As Dictionary(Of VariableDepartamento, Double), ByRef VariableMuniAcum As Dictionary(Of VariableDepartamentoMunicipio, Double), ByRef VariableSexoAcum As Dictionary(Of VariableSexo, Double))
        Dim SqlConn As SqlConnection = GetConnection()

        Dim NumsSexo As New Dictionary(Of Integer, Double)
        Dim DensSexo As New Dictionary(Of Integer, Double)
        Dim NumsDepto As New Dictionary(Of Integer, Double)
        Dim DensDepto As New Dictionary(Of Integer, Double)
        Dim NumsMunis As New Dictionary(Of VariableDepartamentoMunicipio, Double)(New VariableDepartamentoMunicipioComparer)
        Dim DensMunis As New Dictionary(Of VariableDepartamentoMunicipio, Double)(New VariableDepartamentoMunicipioComparer)
        Dim DeptoPair As KeyValuePair(Of VariableDepartamento, Double)
        Dim MuniPair As KeyValuePair(Of VariableDepartamentoMunicipio, Double)
        Dim SexoPair As KeyValuePair(Of VariableSexo, Double)
        For Each Formula As FormulaIndicador In Formulas
            ComputeIndicatorsAndInsertOficialProgram(Formula, CreadoPor)
            Dim Command As New SqlCommand("InsertarValoresIndicadores", SqlConn)
            Dim num As Double
            Dim den As Double
            Dim res As Double
            If VariableAcum.ContainsKey(Formula.Numerador) Then
                num = VariableAcum(Formula.Numerador)
            Else
                num = 0
            End If
            If VariableAcum.ContainsKey(Formula.Denominador) Then
                den = VariableAcum(Formula.Denominador)
            Else
                den = 0
            End If
            If den = 0 Then
                Console.WriteLine("Algo anda mal, el denominador quedó en cero, no hay muestra")
                res = 0
            Else
                res = num / den
            End If
            Command.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
            Command.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
            Command.Parameters.AddWithValue("@ValorIndicador", res)
            Command.Parameters.AddWithValue("@FechaCalculo", Date.Now)
            Command.Parameters.AddWithValue("@CreadoPor", CreadoPor)
            Command.Parameters.AddWithValue("@IndicadorCalculado", True)
            Command.Parameters.AddWithValue("@IndicadorOficial", False)
            Command.Parameters.AddWithValue("@IndicadorPrograma", False)
            Command.CommandType = CommandType.StoredProcedure
            Command.ExecuteNonQuery()
            Console.WriteLine("IdIndicador = " + Convert.ToString(Formula.IdIndicador) + " = " + Convert.ToString(res))
            'Insertar Valores Desagregados
            'Insertar Valores por Departamento ===================================================================
            For Each DeptoPair In VariableDeptoAcum
                If DeptoPair.Key.Variable = Formula.Numerador Then
                    NumsDepto(DeptoPair.Key.Departamento) = DeptoPair.Value
                End If
                If DeptoPair.Key.Variable = Formula.Denominador Then
                    DensDepto(DeptoPair.Key.Departamento) = DeptoPair.Value
                End If
            Next
            If NumsDepto.Count = 0 And DensDepto.Count > 0 Then
                For Each DenNum As KeyValuePair(Of Integer, Double) In DensDepto
                    res = 0
                    Dim CommandDepto As New SqlCommand(InsertValoresDepartameto, SqlConn)
                    CommandDepto.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
                    CommandDepto.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
                    CommandDepto.Parameters.AddWithValue("@IdDepartamento", DenNum.Key)
                    CommandDepto.Parameters.AddWithValue("@Valor", res)
                    CommandDepto.Parameters.AddWithValue("@CreadoPor", CreadoPor)
                    CommandDepto.ExecuteNonQuery()
                Next
            Else
                For Each NumDen As KeyValuePair(Of Integer, Double) In NumsDepto
                    num = NumDen.Value
                    If (DensDepto.ContainsKey(NumDen.Key)) Then
                        den = DensDepto(NumDen.Key)
                    Else
                        num = 0
                    End If
                    If num = 0 Then
                        res = 0
                    ElseIf den = 0 Then
                        res = 0
                    Else
                        res = num / den
                    End If
                    Dim CommandDepto As New SqlCommand(InsertValoresDepartameto, SqlConn)
                    CommandDepto.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
                    CommandDepto.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
                    CommandDepto.Parameters.AddWithValue("@IdDepartamento", NumDen.Key)
                    CommandDepto.Parameters.AddWithValue("@Valor", res)
                    CommandDepto.Parameters.AddWithValue("@CreadoPor", CreadoPor)
                    CommandDepto.ExecuteNonQuery()

                Next
            End If
            '=====================================================================================================
            'Insertar Valores por Municipio ======================================================================
            NumsMunis.Clear()
            DensMunis.Clear()
            For Each MuniPair In VariableMuniAcum
                Dim MunDep As VariableDepartamentoMunicipio
                If MuniPair.Key.Variable = Formula.Numerador Then
                    MunDep = New VariableDepartamentoMunicipio
                    MunDep.Variable = "N"
                    MunDep.Departamento = MuniPair.Key.Departamento
                    MunDep.Municipio = MuniPair.Key.Municipio
                    NumsMunis(MunDep) = MuniPair.Value
                End If
                If MuniPair.Key.Variable = Formula.Denominador Then
                    MunDep = New VariableDepartamentoMunicipio
                    MunDep.Variable = "N"
                    MunDep.Departamento = MuniPair.Key.Departamento
                    MunDep.Municipio = MuniPair.Key.Municipio
                    DensMunis(MunDep) = MuniPair.Value
                End If
            Next
            If NumsMunis.Count = 0 And DensMunis.Count > 0 Then
                For Each DenNum As KeyValuePair(Of VariableDepartamentoMunicipio, Double) In DensMunis
                    res = 0
                    Dim CommandMuni As New SqlCommand(InsertValoresMunicipio, SqlConn)
                    CommandMuni.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
                    CommandMuni.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
                    CommandMuni.Parameters.AddWithValue("@IdDepartamento", DenNum.Key.Departamento)
                    CommandMuni.Parameters.AddWithValue("@IdMunicipio", DenNum.Key.Municipio)
                    CommandMuni.Parameters.AddWithValue("@Valor", res)
                    CommandMuni.Parameters.AddWithValue("@CreadoPor", CreadoPor)
                    CommandMuni.ExecuteNonQuery()
                Next
            Else
                For Each NumDen As KeyValuePair(Of VariableDepartamentoMunicipio, Double) In NumsMunis
                    num = NumDen.Value
                    DensMunis.Keys(0).Equals(NumDen.Key)
                    If DensMunis.ContainsKey(NumDen.Key) Then
                        den = DensMunis(NumDen.Key)
                    Else
                        num = 0
                    End If
                    If num = 0 Then
                        res = 0
                    ElseIf den = 0 Then
                        res = 0
                    Else
                        res = num / den
                    End If
                    Dim CommandMuni As New SqlCommand(InsertValoresMunicipio, SqlConn)
                    CommandMuni.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
                    CommandMuni.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
                    CommandMuni.Parameters.AddWithValue("@IdDepartamento", NumDen.Key.Departamento)
                    CommandMuni.Parameters.AddWithValue("@IdMunicipio", NumDen.Key.Municipio)
                    CommandMuni.Parameters.AddWithValue("@Valor", res)
                    CommandMuni.Parameters.AddWithValue("@CreadoPor", CreadoPor)
                    CommandMuni.ExecuteNonQuery()
                Next

            End If
            '=====================================================================================================
            'Insertar Valores por Sexo ===========================================================================
            For Each SexoPair In VariableSexoAcum
                If SexoPair.Key.Variable = Formula.Numerador Then
                    NumsSexo(SexoPair.Key.Sexo) = SexoPair.Value
                End If
                If SexoPair.Key.Variable = Formula.Denominador Then
                    DensSexo(SexoPair.Key.Sexo) = SexoPair.Value
                End If
            Next
            If NumsSexo.Count = 0 And DensSexo.Count > 0 Then
                For Each DenNum As KeyValuePair(Of Integer, Double) In DensSexo
                    res = 0
                    Dim CommandSexo As New SqlCommand(InsertValoresSexo, SqlConn)
                    CommandSexo.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
                    CommandSexo.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
                    CommandSexo.Parameters.AddWithValue("@IdSexo", DenNum.Key)
                    CommandSexo.Parameters.AddWithValue("@Valor", res)
                    CommandSexo.Parameters.AddWithValue("@CreadoPor", CreadoPor)
                    CommandSexo.ExecuteNonQuery()
                Next
            Else
                For Each NumDen As KeyValuePair(Of Integer, Double) In NumsSexo
                    num = NumDen.Value
                    If DensSexo.ContainsKey(NumDen.Key) Then
                        den = DensSexo(NumDen.Key)
                    Else
                        num = 0
                    End If
                    If num = 0 Then
                        res = 0
                    ElseIf den = 0 Then
                        res = 0
                    Else
                        res = num / den
                    End If
                    Dim CommandSexo As New SqlCommand(InsertValoresSexo, SqlConn)
                    CommandSexo.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
                    CommandSexo.Parameters.AddWithValue("@IdIndicadorEvaluacionPorPrograma", Formula.IdIndicadoresEvaluacionPorPrograma)
                    CommandSexo.Parameters.AddWithValue("@IdSexo", NumDen.Key)
                    CommandSexo.Parameters.AddWithValue("@Valor", res)
                    CommandSexo.Parameters.AddWithValue("@CreadoPor", CreadoPor)
                    CommandSexo.ExecuteNonQuery()
                Next

            End If
            '=====================================================================================================
        Next
        SqlConn.Close()
    End Sub
    Sub Run(ByVal CreadoPor As String)
        Dim ListFichasID As ArrayList
        ListFichasID = GetFichasLevantamiento(IdLevantamiento)
        Dim Formulas As ArrayList
        Dim VariablesConditions As Dictionary(Of String, ConditionTreeNode)
        Dim VarTreePair As KeyValuePair(Of String, ConditionTreeNode)

        Dim VariableAcum As Dictionary(Of String, Double)

        ' ========================================================================================================
        ' Trabajar con Fórmulas de la Ficha Socioeconómica Única
        Dim VariableDeptoAcum As Dictionary(Of VariableDepartamento, Double)
        Dim VariableMuniAcum As Dictionary(Of VariableDepartamentoMunicipio, Double)
        Dim VariableSexoAcum As Dictionary(Of VariableSexo, Double)

        Formulas = GetFormulasFromPrograma(IdPrograma, True)
        VariablesConditions = GetConditionsFromFormulas(Formulas)
        VariableAcum = New Dictionary(Of String, Double)
        VariableDeptoAcum = New Dictionary(Of VariableDepartamento, Double)(New VariableDepartamentoComparer)
        VariableMuniAcum = New Dictionary(Of VariableDepartamentoMunicipio, Double)(New VariableDepartamentoMunicipioComparer)
        VariableSexoAcum = New Dictionary(Of VariableSexo, Double)(New VariableSexoComparer)

        For Each f As ParFSU_IE In ListFichasID
            Dim Ficha As New Ficha(f.CodigoFSU, ConnectionString)
            For Each VarTreePair In VariablesConditions
                ComputeFSUAcumulators(Ficha, VarTreePair, VariableAcum, VariableDeptoAcum, VariableMuniAcum, VariableSexoAcum)
            Next
        Next
        ComputeIndicatorsAndInsert(Formulas, CreadoPor, VariableAcum, VariableDeptoAcum, VariableMuniAcum, VariableSexoAcum)

        ' ========================================================================================================
        ' Trabajar con Fórmulas de los Instrumentos de Evaluación SUEPPS
        Formulas = GetFormulasFromPrograma(IdPrograma, False)
        VariablesConditions = GetConditionsFromFormulas(Formulas)
        VariableAcum.Clear()

        For Each f As ParFSU_IE In ListFichasID
            'Recuperar parte Instrumento de Evaluacion
            Dim Ficha As FichaSU = RetrieveSingleFichaIE(f.CodigoIE)
            For Each VarTreePair In VariablesConditions
                ComputeIEAcumulators(Ficha, VarTreePair, VariableAcum)
            Next
        Next
        ComputeIndicatorsAndInsert(Formulas, CreadoPor, VariableAcum)
    End Sub
    Function GetFormulasFromPrograma(ByVal IdPrograma As Integer, ByVal UsaFSU As Boolean) As ArrayList
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryFormulas, SqlConn)
        Command.Parameters.AddWithValue("@IdPrograma", IdPrograma)
        Command.Parameters.AddWithValue("@BasadoEnFSU", UsaFSU)
        Dim Reader As SqlDataReader = Command.ExecuteReader
        Dim Formulas As New ArrayList
        While Reader.Read
            Dim Formula As New FormulaIndicador(Reader("IdIndicador"), Reader("IdIndicadoresEvaluacionPorPrograma"),
                                                Reader("IdVariableNumerador"), Reader("Numerador"),
                                                Reader("IdVariableDenominador"), Reader("Denominador"),
                                                Reader("UsaVariableMacroNumerador"), Reader("UsaVariableMacroDenominador"))
            Formulas.Add(Formula)
        End While
        Reader.Close()
        SqlConn.Close()
        Return Formulas
    End Function

    Private Function GetConditionsFromFormulas(ByRef Formulas As ArrayList) As Dictionary(Of String, ConditionTreeNode)
        Dim Condiciones As ArrayList
        Dim ConditionRoot As Condicion
        Dim VariablesConditions As New Dictionary(Of String, ConditionTreeNode)
        For Each Formula As FormulaIndicador In Formulas
            'Si no se definido la varible del numerador
            If Not VariablesConditions.ContainsKey(Formula.Numerador) Then
                'Consigue todas las condiciones del numerador
                Condiciones = GetCondiciones(Formula.IdVariableNumerador)
                'Consigue la condición raíz del numerador
                ConditionRoot = GetCondicionesRaiz(Formula.IdVariableNumerador)
                'Crea el par variable, árbol de condición, llama a CreateConditionTree con la raíz y la lista de condiciones
                VariablesConditions.Add(Formula.Numerador, CreateConditionTree(ConditionRoot, Condiciones))
            End If
            'Si no se ha definido la variable del denominador
            If Not VariablesConditions.ContainsKey(Formula.Denominador) Then
                'Consigue todas las condiciones del denominador
                Condiciones = GetCondiciones(Formula.IdVariableDenominador)
                'Consigue la condición raíz del denominador
                ConditionRoot = GetCondicionesRaiz(Formula.IdVariableDenominador)
                'Crea el par variable, árbol de condición, llama a CreateConditionTree con la raíz y la lista de condiciones
                VariablesConditions.Add(Formula.Denominador, CreateConditionTree(ConditionRoot, Condiciones))
            End If
        Next
        Return VariablesConditions
    End Function
    Private Function CreateConditionTree(ByRef C As Condicion, ByRef List As ArrayList) As ConditionTreeNode
        If C.Total Then
            Dim Tree As New ConditionTreeNode("T")
            Tree.Total = True
            Tree.Level = C.Operando1(0)
            Return Tree
        End If
        Dim Fuente As String
        Dim Level As Char
        If C.IdTipoCondicion <> 3 Then
            If (C.IdTipoCondicion = 1) Then
                Fuente = "FSU"
                If C.Operando1(0) = "P" Then
                    Level = "P"
                ElseIf C.Operando1(0) = "H" Then
                    Level = "H"
                Else
                    Level = "V"
                End If
            Else
                Fuente = "IE"
            End If
            Return New ConditionTreeNode(Fuente, C.Operando1, C.Operador, C.Operando2, Level)
        Else
            Dim root As ConditionTreeNode
            If C.Operador = "AND" Then
                root = New ConditionTreeNode("Y")
            Else
                root = New ConditionTreeNode("O")
            End If
            Dim C1, C2 As Condicion
            C1 = Nothing
            C2 = Nothing
            Dim f1, f2 As Boolean
            f1 = False
            f2 = False
            For Each R As Condicion In List
                If R.IdCondicion = Convert.ToUInt64(C.Operando1) Then
                    C1 = R
                    f1 = True
                End If
                If R.IdCondicion = Convert.ToUInt64(C.Operando2) Then
                    C2 = R
                    f2 = True
                End If
                If f1 And f2 Then
                    Exit For
                End If
            Next

            root.LeftNode = CreateConditionTree(C1, List)
            root.RightNode = CreateConditionTree(C2, List)
            If root.LeftNode.Level = "P" Or root.RightNode.Level = "P" Then
                root.Level = "P"
            ElseIf root.LeftNode.Level = "H" Or root.RightNode.Level = "H" Then
                root.Level = "H"
            Else
                root.Level = "V"
            End If
            Return root
        End If
    End Function
    Private Function GetFormulasFromPrograma() As ArrayList
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand("RecuperarFormulas", SqlConn)
        Command.Parameters.AddWithValue("@IdPrograma", IdPrograma)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader = Command.ExecuteReader
        Dim Formulas As New ArrayList
        While Reader.Read
            Dim Formula As New FormulaIndicador(Reader("IdIndicador"), Reader("IdIndicadoresEvaluacionPorPrograma"),
                                                Reader("IdVariableNumerador"), Reader("Numerador"),
                                                Reader("IdVariableDenominador"), Reader("Denominador"),
                                                Reader("UsaVariableMacroNumerador"), Reader("UsaVariableMacroDenominador"))
            Formulas.Add(Formula)
        End While
        Reader.Close()
        SqlConn.Close()
        Return Formulas
    End Function
    Private Function GetConnection() As SqlConnection
        Dim SqlConn As New SqlConnection
        SqlConn.ConnectionString = ConnectionString
        SqlConn.Open()
        Return SqlConn
    End Function
    Private Function GetFichasLevantamiento(ByVal IdLevantamiento As Integer) As ArrayList
        'Retorna todas las fichas de la muestra para la prueba
        'En realidad este debería de tener como parámetro el levantamiento y a partir de ahí traer las fichas
        'de dicho levantamiento, que son el par(FichaFSU, FichaIE)

        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand("RecuperarFichasPorLevantamiento", SqlConn)
        Command.Parameters.AddWithValue("@IdLevantamiento", IdLevantamiento)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader = Command.ExecuteReader
        Dim List As New ArrayList
        While Reader.Read
            Dim Ficha As New ParFSU_IE(Reader("CodigoFSU"), Reader("IdVivienda"), Reader("IdHogar"), Reader("IdMiembro"), Reader("IdEncabezadoRespuesta"))
            List.Add(Ficha)
        End While
        Reader.Close()
        SqlConn.Close()
        Return List

        'Temporalmente la funcion retornara todas las fichas de 1 a IdLevantamiento de la muestra FSU provista
        'Dim List As New ArrayList
        'For i = 1 To IdLevantamiento
        '    List.Add(New ParFSU_IE(i, i))
        'Next
        'Return List
    End Function
    Private Function GetCondiciones(ByVal IdVariable As Integer) As ArrayList
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryCondiciones, SqlConn)
        Command.Parameters.AddWithValue("@IdVariable", Convert.ToString(IdVariable))
        Dim Reader As SqlDataReader = Command.ExecuteReader
        Dim Condiciones As New ArrayList
        While Reader.Read
            Dim Condition As New Condicion(Reader("IdCondicion"), Reader("IdVariable"), Reader("IdTipoCondicion"),
                                           Reader("NombreCondicion"), Reader("Raiz"), Reader("Total"), Reader("Operando1"),
                                           Reader("Operador"), Reader("Operando2"))
            Condiciones.Add(Condition)
        End While
        Reader.Close()
        SqlConn.Close()
        Return Condiciones
    End Function
    Private Function GetCondicionesRaiz(ByVal IdVariable As Integer) As Condicion
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryRaiz, SqlConn)
        Command.Parameters.AddWithValue("@IdVariable", Convert.ToString(IdVariable))
        Dim Reader As SqlDataReader = Command.ExecuteReader
        Dim Condition As Condicion
        If Reader.Read Then
            Condition = New Condicion(Reader("IdCondicion"), Reader("IdVariable"), Reader("IdTipoCondicion"),
                                           Reader("NombreCondicion"), Reader("Raiz"), Reader("Total"), Reader("Operando1"),
                                           Reader("Operador"), Reader("Operando2"))
        Else
            Condition = Nothing
        End If
        Return Condition
    End Function

    Private Function RetrieveSingleFichaIE(CodigoIE As Integer) As FichaSU
        Dim Ficha As New FichaSU(CodigoIE)
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand(QueryRecuperaFichaIE, SqlConn)
        Command.Parameters.AddWithValue("@IdEncabezadoRespuesta", CodigoIE)
        Dim Reader As SqlDataReader = Command.ExecuteReader
        While Reader.Read
            Dim TipoPregunta As Integer = Reader("IdTipoDePregunta")
            If TipoPregunta = 4 Or TipoPregunta = 6 Or TipoPregunta = 7 Then
                Ficha.SetValorRespuestaUnica(Reader("IdAmigable"), Reader("Valor"))
            Else
                Ficha.AddValorRespuestaMultiple(Reader("IdAmigable"), Reader("Valor"))
            End If
        End While
        Return Ficha
    End Function

End Class
