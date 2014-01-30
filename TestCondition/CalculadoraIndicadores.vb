Imports System.Data.SqlClient
Imports System.Data

Public Class CalculadoraIndicadores
    Private ConnectionString As String
    Private QueryCondiciones As String = "SELECT * FROM Condiciones WHERE IdVariable = @IdVariable"
    Private QueryRaiz As String = "SELECT * FROM Condiciones WHERE Raiz = 1 AND IdVariable = @IdVariable"
    Private QueryFormulas As String = "SELECT IEP.IdIndicador, IEP.IdIndicadoresEvaluacionPorPrograma, FI.IdVariableNumerador, " _
                               & "V1.NombreVariable as Numerador, FI.IdVariableDenominador, V2.NombreVariable as Denominador, " _
                               & "FI.UsaVariableMacroNumerador, FI.UsaVariableMacroDenominador " _
                               & "FROM IndicadoresEvaluacionPorPrograma IEP " _
                               & "JOIN FormulaIndicador FI ON IEP.IdIndicador=FI.IdIndicador " _
                               & "JOIN Variables V1 ON FI.IdVariableNumerador = V1.IdVariable " _
                               & "JOIN Variables V2 ON FI.IdVariableDenominador = V2.IdVariable " _
                               & "WHERE IEP.IdPrograma = @IdPrograma"
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

    Sub Run(ByVal CreadoPor As String)
        If (UsaFSU) Then 'Cuando se usa la Ficha Socioeconómica Única
            Dim Formulas As ArrayList
            Formulas = GetFormulasFromPrograma(IdPrograma)
            Dim VariablesConditions As Dictionary(Of String, ConditionTreeNode)
            VariablesConditions = GetConditionsFromFormulas(Formulas)
            'Para pruebas se barrerán las fichas provistas, ignorando el levantamiento
            Dim ListFichasID As ArrayList
            ' OJO ESTE DEBE RECIBIR IdLevantamiento ****************************************************************************************
            ListFichasID = GetFichasLevantamiento(IdLevantamiento)
            'Dim ListFichas As New ArrayList
            'Agrega todas las fichas del levantamiento a ListFichas
            'Ya no vamos a traer todas las fichas, las vamos a recuperar de una en una
            'For Each f As ParFSU_IE In ListFichasID
            '    ListFichas.AddRange(RetrieveSingleFichaAllMembers(f.CodigoFSU))
            'Next
            'Verificar cuál es el niveles de recuperación de fichas necesarios, hay que traer todos los necesarios
            Dim VarTreePair As KeyValuePair(Of String, ConditionTreeNode)

            'Acá se almacenarán todos los valores
            Dim VariableAcum As New Dictionary(Of String, Double)
            Dim VariableDeptoAcum As New Dictionary(Of VariableDepartamento, Double)(New VariableDepartamentoComparer)
            Dim VariableMuniAcum As New Dictionary(Of VariableDepartamentoMunicipio, Double)(New VariableDepartamentoMunicipioComparer)
            Dim VariableSexoAcum As New Dictionary(Of VariableSexo, Double)(New VariableSexoComparer)
            For Each f As ParFSU_IE In ListFichasID

                'Ahora hay información de IdVivienda, IdHogar, IdMiembro en ParFSU_IE... los IE son todos de población o pueden ser de hogar?
                'Habrá cálculo de indicadores sin que haya IE?
                'Se combinarán datos de fsu con ie?


                'Recuperar parte Instrumento de Evaluacion
                Dim FichaIE As FichaSU = RetrieveSingleFichaIE(f.CodigoIE)
                'Recuperar la parte vivienda de la ficha UNA INSTANCIA
                Dim FichaVivienda As FichaSU = RetrieveSingleFichaForVivienda(f.CodigoFSU)
                'Recuperar los hogares de la ficha       UNA LISTA
                Dim FichasHogares As ArrayList = RetrieveSingleFichaAllHogares(f.CodigoFSU)
                'Recuperar los miembros de la ficha      UNA LISTA
                Dim FichasMiembros As ArrayList = RetrieveSingleFichaAllMembers(f.CodigoFSU)
                Dim VariablePoblacion As Boolean
                For Each VarTreePair In VariablesConditions
                    'Si la Condicion, en la raiz tiene el nivel máximo, que se ve es de vivienda entonces usar la de vivienda
                    'Si la Condicion, en la raiz tiene el nivel máximo, que se ve es de hogar hacer por cada hogar
                    'Si la Condicion, en la raiz tiene el nivel máximo, que se ve es de miembro hacer por cada miembro
                    Dim ListaFichas As ArrayList
                    If VarTreePair.Value.Level = "V" Then
                        ListaFichas = New ArrayList
                        ListaFichas.Add(FichaVivienda)
                        VariablePoblacion = False
                    ElseIf VarTreePair.Value.Level = "H" Then
                        ListaFichas = FichasHogares
                        VariablePoblacion = False
                    Else
                        ListaFichas = FichasMiembros
                        VariablePoblacion = True
                    End If
                    ' A cada una de las fichas de la lista agregar la parte de IE
                    For Each Ficha As FichaSU In ListaFichas
                        Ficha.MergeFSUWithIE(FichaIE)
                    Next
                    ' Hay que verificar si hay lista de FSU, pues puede venir la IE sola la verificación se hace arriba.
                    For Each Ficha As FichaSU In ListaFichas
                        Dim VarDepto As VariableDepartamento = Ficha.GetDepartamento
                        VarDepto.Variable = VarTreePair.Key
                        Dim VarDeptoMuni As VariableDepartamentoMunicipio = Ficha.GetDepartamentoMunicipio
                        VarDeptoMuni.Variable = VarTreePair.Key
                        Dim VarSexo As VariableSexo
                        If VariablePoblacion Then
                            VarSexo = Ficha.GetSexo
                            VarSexo.Variable = VarTreePair.Key
                        End If
                        If VarTreePair.Value.Evaluate(Ficha) Then
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
                            If VariablePoblacion Then
                                If VariableSexoAcum.ContainsKey(VarSexo) Then
                                    VariableSexoAcum(VarSexo) = VariableSexoAcum(VarSexo) + 1
                                Else
                                    VariableSexoAcum(VarSexo) = 1
                                End If
                            End If
                            '=========================================================
                            If VariableAcum.ContainsKey(VarTreePair.Key) Then
                                VariableAcum(VarTreePair.Key) = VariableAcum(VarTreePair.Key) + 1
                            Else
                                VariableAcum(VarTreePair.Key) = 1
                            End If
                        End If
                        'Console.WriteLine(VarTreePair.Key)
                        'VarTreePair.Value.PrintTree()
                    Next
                Next
            Next
            'Dim VarAcumPair As KeyValuePair(Of String, Double)
            'For Each VarAcumPair In VariableAcum
            '    Console.WriteLine(VarAcumPair.Key + " " + Convert.ToString(VarAcumPair.Value))
            'Next
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
                        '@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdDepartamento,@Valor,@CreadorPor
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
                        '@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdDepartamento,@Valor,@CreadorPor
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
                        '@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdDepartamento, @IdMunicipio, @Valor,@CreadorPor
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
                        '@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdDepartamento, @IdMunicipio, @Valor,@CreadorPor
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
                        '@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdSexo,@Valor,@CreadorPor
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
                        '@IdLevantamiento,@IdIndicadorEvaluacionPorPrograma,@IdSexo,@Valor,@CreadorPor
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
        Else 'Sólo usa el Instrumento de Evaluación
            Dim Formulas As ArrayList
            Formulas = GetFormulasFromPrograma(IdPrograma)
            Dim VariablesConditions As Dictionary(Of String, ConditionTreeNode)
            VariablesConditions = GetConditionsFromFormulas(Formulas)
            'Para pruebas se barrerán las fichas provistas, ignorando el levantamiento
            Dim ListFichasID As ArrayList
            ' OJO ESTE DEBE RECIBIR IdLevantamiento *****************************************************************************************
            ListFichasID = GetFichasLevantamiento(IdLevantamiento)
            'Dim ListFichas As New ArrayList
            'Agrega todas las fichas del levantamiento a ListFichas
            'Ya no vamos a traer todas las fichas, las vamos a recuperar de una en una
            'For Each f As ParFSU_IE In ListFichasID
            '    ListFichas.AddRange(RetrieveSingleFichaAllMembers(f.CodigoFSU))
            'Next
            'Verificar cuál es el niveles de recuperación de fichas necesarios, hay que traer todos los necesarios
            Dim VarTreePair As KeyValuePair(Of String, ConditionTreeNode)

            'Acá se almacenarán todos los valores
            Dim VariableAcum As New Dictionary(Of String, Double)
            For Each f As ParFSU_IE In ListFichasID
                'Recuperar parte Instrumento de Evaluacion
                Dim Ficha As FichaSU = RetrieveSingleFichaIE(f.CodigoIE)
                For Each VarTreePair In VariablesConditions
                    If VarTreePair.Value.Evaluate(Ficha) Then
                        If VariableAcum.ContainsKey(VarTreePair.Key) Then
                            VariableAcum(VarTreePair.Key) = VariableAcum(VarTreePair.Key) + 1
                        Else
                            VariableAcum(VarTreePair.Key) = 1
                        End If
                    End If
                Next
            Next
            Dim SqlConn As SqlConnection = GetConnection()

            For Each Formula As FormulaIndicador In Formulas
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
                Command.CommandType = CommandType.StoredProcedure
                Command.ExecuteNonQuery()
            Next
            SqlConn.Close()
        End If

    End Sub
    Function GetFormulasFromPrograma(ByVal IdPrograma As Integer) As ArrayList
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryFormulas, SqlConn)
        Command.Parameters.AddWithValue("@IdPrograma", IdPrograma)
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
    Private Function RetrieveSingleFichaForVivienda(ByVal IdFicha As Integer) As FichaSU
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("ViviendaPorFicha", SqlConn)
        Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader
        Reader = Command.ExecuteReader
        Dim Ficha As FichaSU
        While Reader.Read
            Ficha = New FichaSU(IdFicha, Reader("IdVivienda"), "V")
            Ficha.SetValorRespuestaUnica("V1", Reader("v1"))
            Ficha.SetValorRespuestaUnica("V2", Reader("v2"))
            Ficha.SetValorRespuestaUnica("V3", Reader("v3"))
            Ficha.SetValorRespuestaUnica("V4", Reader("v4"))
            Ficha.SetValorRespuestaUnica("V5", Reader("v5"))
            Ficha.SetValorRespuestaUnica("V6", Reader("v6"))
            Ficha.SetValorRespuestaUnica("V7", Reader("v7"))
            Ficha.SetValorRespuestaUnica("V8", Reader("v8"))
            Ficha.SetValorRespuestaUnica("V8_Pago", Reader("v8_Pago"))
            Ficha.SetValorRespuestaUnica("V9", Reader("v9"))
            Ficha.SetValorRespuestaUnica("V10", Reader("v10"))
            Ficha.SetValorRespuestaUnica("V12", Reader("v12"))
            'Valores Desagregación por Ubicación Geográfica
            Ficha.SetValorRespuestaUnica("V01", Reader("V01"))  'Departamento
            Ficha.SetValorRespuestaUnica("V02", Reader("V02"))  'Municipio
            Ficha.SetValorRespuestaUnica("V03", Reader("V03"))  'Aldea    
            Ficha.SetValorRespuestaUnica("V04", Reader("V04"))  'Caserio  
            Ficha.SetValorRespuestaUnica("V05", Reader("V05"))  'Barrio   
            Ficha.SetValorRespuestaUnica("V06", Reader("V06"))  'Region
            ' =================================================
            Dim ASqlConn As SqlConnection = GetConnection()
            Dim ACommand As New SqlCommand("AmenazasPorViviendaPorFicha", ASqlConn)
            ACommand.Parameters.AddWithValue("@CodigoFSU", IdFicha)
            ACommand.Parameters.AddWithValue("@IdVivienda", Reader("IdVivienda"))
            ACommand.CommandType = CommandType.StoredProcedure
            Dim AReader As SqlDataReader
            AReader = ACommand.ExecuteReader
            While AReader.Read
                Ficha.AddValorRespuestaMultiple("V11", AReader("v11"))
            End While
            AReader.Close()
            ASqlConn.Close()
        End While
        Reader.Close()
        SqlConn.Close()
        Return Ficha
    End Function
    Private Function RetrieveSingleFichaAllHogares(ByVal IdFicha As Integer) As ArrayList
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("HogaresPorVivienda", SqlConn)
        Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader
        Reader = Command.ExecuteReader
        Dim Ficha As FichaSU
        Dim ListaFichas As New ArrayList
        While Reader.Read
            Ficha = New FichaSU(IdFicha, Reader("IdVivienda"), "M", Reader("IdHogar"))
            'Datos de Vivienda - Encabezado
            Ficha.SetValorRespuestaUnica("V1", Reader("v1"))
            Ficha.SetValorRespuestaUnica("V2", Reader("v2"))
            Ficha.SetValorRespuestaUnica("V3", Reader("v3"))
            Ficha.SetValorRespuestaUnica("V4", Reader("v4"))
            Ficha.SetValorRespuestaUnica("V5", Reader("v5"))
            Ficha.SetValorRespuestaUnica("V6", Reader("v6"))
            Ficha.SetValorRespuestaUnica("V7", Reader("v7"))
            Ficha.SetValorRespuestaUnica("V8", Reader("v8"))
            Ficha.SetValorRespuestaUnica("V8_Pago", Reader("v8_Pago"))
            Ficha.SetValorRespuestaUnica("V9", Reader("v9"))
            Ficha.SetValorRespuestaUnica("V10", Reader("v10"))
            Ficha.SetValorRespuestaUnica("V12", Reader("v12"))
            'Valores Desagregación por Ubicación Geográfica
            Ficha.SetValorRespuestaUnica("V01", Reader("V01"))  'Departamento
            Ficha.SetValorRespuestaUnica("V02", Reader("V02"))  'Municipio
            Ficha.SetValorRespuestaUnica("V03", Reader("V03"))  'Aldea    
            Ficha.SetValorRespuestaUnica("V04", Reader("V04"))  'Caserio  
            Ficha.SetValorRespuestaUnica("V05", Reader("V05"))  'Barrio   
            Ficha.SetValorRespuestaUnica("V06", Reader("V06"))  'Region
            ' =================================================
            Dim ASqlConn As SqlConnection = GetConnection()
            Dim ACommand As New SqlCommand("AmenazasPorViviendaPorFicha", ASqlConn)
            ACommand.Parameters.AddWithValue("@CodigoFSU", IdFicha)
            ACommand.Parameters.AddWithValue("@IdVivienda", Reader("IdVivienda"))
            ACommand.CommandType = CommandType.StoredProcedure
            Dim AReader As SqlDataReader
            AReader = ACommand.ExecuteReader
            While AReader.Read
                Ficha.AddValorRespuestaMultiple("V11", AReader("v11"))
            End While
            AReader.Close()
            ASqlConn.Close()
            Ficha.SetValorRespuestaUnica("H1", Reader("H1"))
            Ficha.SetValorRespuestaUnica("H2_Hombres", Reader("H2_Hombres"))
            Ficha.SetValorRespuestaUnica("H2_Mujeres", Reader("H2_Mujeres"))
            Ficha.SetValorRespuestaUnica("H2_Total", Reader("H2_Total"))
            Ficha.SetValorRespuestaUnica("H3", Reader("H3"))
            Ficha.SetValorRespuestaUnica("H6", Reader("H6"))
            'Ficha.SetValorRespuestaUnica("H7", Reader("H7"))
            Ficha.SetValorRespuestaUnica("H8", Reader("H8"))
            Ficha.SetValorRespuestaUnica("H8_Estud", Reader("H8_Estud"))
            Ficha.SetValorRespuestaUnica("H9_Bus", Reader("H9_Bus"))
            Ficha.SetValorRespuestaUnica("H9_Taxi", Reader("H9_Taxi"))
            Ficha.SetValorRespuestaUnica("H10", Reader("H10"))

            Dim SqlConn2 As SqlConnection = GetConnection()
            Dim Command2 As New SqlCommand("OrganizacionesComunitariasPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Dim Reader2 As SqlDataReader
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("H4", Reader2("H4"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("BienesPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("H5", Reader2("H5"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("ServiciosFinancierosPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("H11", Reader2("H11"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            ListaFichas.Add(Ficha)

        End While
        Reader.Close()
        SqlConn.Close()
        Return ListaFichas
    End Function

    Private Function RetrieveSingleFichaAllMembers(ByVal IdFicha As Integer) As ArrayList
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("PoblacionPorHogar", SqlConn)
        Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader
        Reader = Command.ExecuteReader
        Dim Ficha As FichaSU
        Dim ListaFichas As New ArrayList
        While Reader.Read
            Ficha = New FichaSU(IdFicha, Reader("IdVivienda"), "M", Reader("IdHogar"), Reader("IdMiembro"))
            'Datos de Vivienda - Encabezado
            Ficha.SetValorRespuestaUnica("V1", Reader("v1"))
            Ficha.SetValorRespuestaUnica("V2", Reader("v2"))
            Ficha.SetValorRespuestaUnica("V3", Reader("v3"))
            Ficha.SetValorRespuestaUnica("V4", Reader("v4"))
            Ficha.SetValorRespuestaUnica("V5", Reader("v5"))
            Ficha.SetValorRespuestaUnica("V6", Reader("v6"))
            Ficha.SetValorRespuestaUnica("V7", Reader("v7"))
            Ficha.SetValorRespuestaUnica("V8", Reader("v8"))
            Ficha.SetValorRespuestaUnica("V8_Pago", Reader("v8_Pago"))
            Ficha.SetValorRespuestaUnica("V9", Reader("v9"))
            Ficha.SetValorRespuestaUnica("V10", Reader("v10"))
            Ficha.SetValorRespuestaUnica("V12", Reader("v12"))
            'Valores Desagregación por Ubicación Geográfica
            Ficha.SetValorRespuestaUnica("V01", Reader("V01"))  'Departamento
            Ficha.SetValorRespuestaUnica("V02", Reader("V02"))  'Municipio
            Ficha.SetValorRespuestaUnica("V03", Reader("V03"))  'Aldea    
            Ficha.SetValorRespuestaUnica("V04", Reader("V04"))  'Caserio  
            Ficha.SetValorRespuestaUnica("V05", Reader("V05"))  'Barrio   
            Ficha.SetValorRespuestaUnica("V06", Reader("V06"))  'Region
            ' =================================================
            Dim ASqlConn As SqlConnection = GetConnection()
            Dim ACommand As New SqlCommand("AmenazasPorViviendaPorFicha", ASqlConn)
            ACommand.Parameters.AddWithValue("@CodigoFSU", IdFicha)
            ACommand.Parameters.AddWithValue("@IdVivienda", Reader("IdVivienda"))
            ACommand.CommandType = CommandType.StoredProcedure
            Dim AReader As SqlDataReader
            AReader = ACommand.ExecuteReader
            While AReader.Read
                Ficha.AddValorRespuestaMultiple("V11", AReader("v11"))
            End While
            AReader.Close()
            ASqlConn.Close()
            Ficha.SetValorRespuestaUnica("H1", Reader("H1"))
            Ficha.SetValorRespuestaUnica("H2_Hombres", Reader("H2_Hombres"))
            Ficha.SetValorRespuestaUnica("H2_Mujeres", Reader("H2_Mujeres"))
            Ficha.SetValorRespuestaUnica("H2_Total", Reader("H2_Total"))
            Ficha.SetValorRespuestaUnica("H3", Reader("H3"))
            Ficha.SetValorRespuestaUnica("H6", Reader("H6"))
            'Ficha.SetValorRespuestaUnica("H7", Reader("H7"))
            Ficha.SetValorRespuestaUnica("H8", Reader("H8"))
            Ficha.SetValorRespuestaUnica("H8_Estud", Reader("H8_Estud"))
            Ficha.SetValorRespuestaUnica("H9_Bus", Reader("H9_Bus"))
            Ficha.SetValorRespuestaUnica("H9_Taxi", Reader("H9_Taxi"))
            Ficha.SetValorRespuestaUnica("H10", Reader("H10"))

            Dim SqlConn2 As SqlConnection = GetConnection()
            Dim Command2 As New SqlCommand("OrganizacionesComunitariasPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Dim Reader2 As SqlDataReader
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("H4", Reader2("H4"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("BienesPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("H5", Reader2("H5"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("ServiciosFinancierosPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("H11", Reader2("H11"))
            End While
            Reader2.Close()
            SqlConn2.Close()
            Ficha.SetValorRespuestaUnica("P5", Reader("P5"))
            Ficha.SetValorRespuestaUnica("P8", Reader("P8"))
            Ficha.SetValorRespuestaUnica("P9", Reader("P9"))
            Ficha.SetValorRespuestaUnica("P9_Emb", Reader("P9_Emb"))
            Ficha.SetValorRespuestaUnica("P10", Reader("P10"))
            Ficha.SetValorRespuestaUnica("P11", Reader("P11"))
            Ficha.SetValorRespuestaUnica("P12", Reader("P12"))
            Ficha.SetValorRespuestaUnica("P13", Reader("P13"))
            Ficha.SetValorRespuestaUnica("P14", Reader("P14"))
            Ficha.SetValorRespuestaUnica("P15", Reader("P15"))
            Ficha.SetValorRespuestaUnica("P15_Razon", Reader("P15_Razon"))
            Ficha.SetValorRespuestaUnica("P16", Reader("P16"))
            Ficha.SetValorRespuestaUnica("P17", Reader("P17"))
            Ficha.SetValorRespuestaUnica("P18", Reader("P18"))
            Ficha.SetValorRespuestaUnica("P19", Reader("P19"))
            Dim SqlConn3 As SqlConnection = GetConnection()
            Dim Command3 As New SqlCommand("DiscapacidadesPorMiembro", SqlConn3)
            Command3.Parameters.AddWithValue("@IdMiembro", Reader("IdMiembro"))
            Command3.CommandType = CommandType.StoredProcedure
            Dim Reader3 As SqlDataReader
            Reader3 = Command3.ExecuteReader
            While Reader3.Read
                Ficha.AddValorRespuestaMultiple("P20", Reader3("P20"))
            End While
            Reader3.Close()
            SqlConn3.Close()
            SqlConn3 = GetConnection()
            Command3 = New SqlCommand("ProgramasSocialesPorMiembro", SqlConn3)
            Command3.Parameters.AddWithValue("@IdMiembro", Reader("IdMiembro"))
            Command3.CommandType = CommandType.StoredProcedure
            Reader3 = Command3.ExecuteReader
            While Reader3.Read
                Ficha.AddValorRespuestaMultiple("P21", Reader3("P21"))
            End While
            Reader3.Close()
            SqlConn3.Close()
            ListaFichas.Add(Ficha)

        End While
        Reader.Close()
        SqlConn.Close()
        Return ListaFichas
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
