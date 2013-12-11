Imports System.Data.SqlClient

Module Module1
    Dim ConnectionString As String = "Data Source=localhost;Initial Catalog=SUEPPS-BD;Integrated Security=True"
    Dim QueryFichas As String = "SELECT ER.CodigoFSU, ER.IdEncabezadoRespuesta " _
                                & "FROM EncabezadoRespuesta ER " _
                                & "JOIN AplicacionInstrumento AI ON ER.IdAplicacionInstrumento = AI.IdAplicacionInstrumento " _
                                & "WHERE AI.IdAplicacionInstrumento ="
    Dim QueryFormulas As String = "SELECT IEP.IdIndicador, IEP.IdIndicadoresEvaluacionPorPrograma, FI.IdVariableNumerador, " _
                                  & "V1.NombreVariable as Numerador, FI.IdVariableDenominador, V2.NombreVariable as Denominador, " _
                                  & "FI.UsaVariableMacroNumerador, FI.UsaVariableMacroDenominador " _
                                  & "FROM IndicadoresEvaluacionPorPrograma IEP " _
                                  & "JOIN FormulaIndicador FI ON IEP.IdIndicador=FI.IdIndicador " _
                                  & "JOIN Variables V1 ON FI.IdVariableNumerador = V1.IdVariable " _
                                  & "JOIN Variables V2 ON FI.IdVariableDenominador = V2.IdVariable " _
                                  & "WHERE IEP.IdPrograma ="
    Dim QueryCondiciones As String = "SELECT * FROM Condiciones WHERE IdVariable = "
    Dim QueryRaiz As String = "SELECT * FROM Condiciones WHERE Raiz = 1 AND IdVariable = "
    Sub PrevTest()
        Dim Formulas As ArrayList
        Formulas = GetFormulasFromPrograma(133)
        Dim VariablesConditions As Dictionary(Of String, ConditionTreeNode)
        Dim Ficha As FichaSU
        Dim ListaFichas As ArrayList
        VariablesConditions = DisplayTable(Formulas)
        Console.ReadKey()
        GetMiembrosFromFicha(27)
        '59
        Console.ReadKey()
        Console.WriteLine("Vivienda Ficha")
        Ficha = RetrieveSingleFichaForVivienda(7)
        Ficha.PrintAllValues()
        Console.ReadKey()
        Console.WriteLine("Poblacion Ficha")
        ListaFichas = RetrieveSingleFichaAllMembers(7)
        Dim i As Integer = 1
        For Each Ficha In ListaFichas
            Console.WriteLine("Miembro # " + Convert.ToString(i))
            Ficha.PrintAllValues()
            i = i + 1
        Next

    End Sub
    Sub Main()
        'PrevTest()
        'TestCheckFichas(133, 1)
        'Programa 133 que tiene asociados dos indicadores, 5 levantamiento que está en la bd de prueba, sólo
        'para que pase el FK constraint de ValoresIndicadores
        Dim Calculadora As New CalculadoraIndicadores(ConnectionString, 133, 5)
        Calculadora.Run("arias-test")
        Console.ReadKey()

    End Sub
    Sub TestCheckFichas(ByVal IdPrograma As Integer, ByVal IdLevantamiento As Integer)
        Dim Formulas As ArrayList
        Formulas = GetFormulasFromPrograma(IdPrograma)
        Dim VariablesConditions As Dictionary(Of String, ConditionTreeNode)
        VariablesConditions = GetConditionsFromFormulas(Formulas)
        'Para pruebas se barrerán las fichas provistas, ignorando el levantamiento
        Dim ListFichasID As ArrayList
        ListFichasID = GetFichasLevantamiento(50)
        Dim ListFichas As New ArrayList
        'Agrega todas las fichas del levantamiento a ListFichas
        For Each f In ListFichasID
            ListFichas.AddRange(RetrieveSingleFichaAllMembers(f))
        Next
        'Acá se almacenarán todos los valores
        Dim VariableAcum As New Dictionary(Of String, Double)
        Dim VarTreePair As KeyValuePair(Of String, ConditionTreeNode)
        For Each f As FichaSU In ListFichas
            For Each VarTreePair In VariablesConditions
                If VarTreePair.Value.Evaluate(f) Then
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
        Dim VarAcumPair As KeyValuePair(Of String, Double)
        For Each VarAcumPair In VariableAcum
            Console.WriteLine(VarAcumPair.Key + " " + Convert.ToString(VarAcumPair.Value))
        Next
        For Each Formula As FormulaIndicador In Formulas
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
            Console.WriteLine("IdIndicador = " + Convert.ToString(Formula.IdIndicador) + " = " + Convert.ToString(res))
        Next
    End Sub
    Function GetFichasLevantamiento(ByVal Count As Integer) As ArrayList
        'Retorna todas las fichas de la muestra para la prueba
        'En realidad este debería de tener como parámetro el levantamiento y a partir de ahí traer las fichas
        'de dicho levantamiento
        Dim List As New ArrayList
        For i = 1 To Count
            List.Add(i)
        Next
        Return List
    End Function
    Sub GetMiembrosFromFicha(ByVal IdFicha)
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("ViviendaPorFicha", SqlConn)
        Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader
        Reader = Command.ExecuteReader
        While Reader.Read
            Console.WriteLine(
                Convert.ToString(Reader("CodigoFSU")) + " " +
                Convert.ToString(Reader("V1")) + " " +
                Convert.ToString(Reader("V2")) + " " +
                Convert.ToString(("V3")))
        End While
        Reader.Close()
        SqlConn.Close()
    End Sub
    Function RetrieveSingleFichaForVivienda(ByVal IdFicha As Integer) As FichaSU
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

    Function RetrieveSingleFichaAllMembers(ByVal IdFicha As Integer) As ArrayList
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
            Dim SqlConn2 As SqlConnection = GetConnection()
            Dim Command2 As New SqlCommand("DiscapacidadesPorMiembro", SqlConn2)
            Command2.Parameters.AddWithValue("@IdMiembro", Reader("IdMiembro"))
            Command2.CommandType = CommandType.StoredProcedure
            Dim Reader2 As SqlDataReader
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("P20", Reader2("P20"))
            End While
            Reader2.Close()
            SqlConn2.Close()
            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("ProgramasSocialesPorMiembro", SqlConn2)
            Command2.Parameters.AddWithValue("@IdMiembro", Reader("IdMiembro"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                Ficha.AddValorRespuestaMultiple("P21", Reader2("P21"))
            End While
            Reader2.Close()
            SqlConn2.Close()
            ListaFichas.Add(Ficha)

        End While
        Reader.Close()
        SqlConn.Close()
        Return ListaFichas
    End Function
    Function DisplayTable(ByRef Formulas As ArrayList) As Dictionary(Of String, ConditionTreeNode)
        Dim Condiciones As ArrayList
        Dim ConditionRoot As Condicion
        Dim VariablesConditions As New Dictionary(Of String, ConditionTreeNode)
        For Each Formula As FormulaIndicador In Formulas
            Console.WriteLine(Formula.ToString)
            Condiciones = GetCondiciones(Formula.IdVariableNumerador)
            ConditionRoot = GetCondicionesRaiz(Formula.IdVariableNumerador)
            'Crea el par variable, árbol de condición
            VariablesConditions.Add(Formula.Numerador, CreateConditionTree(ConditionRoot, Condiciones))
            Console.WriteLine("Condiciones Numerador:")
            For Each Condition As Condicion In Condiciones
                Console.WriteLine(Condition.ToString)
            Next
            Condiciones = GetCondiciones(Formula.IdVariableDenominador)
            ConditionRoot = GetCondicionesRaiz(Formula.IdVariableDenominador)
            VariablesConditions.Add(Formula.Denominador, CreateConditionTree(ConditionRoot, Condiciones))
            Console.WriteLine("Condiciones Denominador:")
            For Each Condition As Condicion In Condiciones
                Console.WriteLine(Condition.ToString)
            Next
        Next
        Return VariablesConditions
    End Function
    Function GetConditionsFromFormulas(ByRef Formulas As ArrayList) As Dictionary(Of String, ConditionTreeNode)
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
    Function CreateConditionTree(ByRef C As Condicion, ByRef List As ArrayList) As ConditionTreeNode
        Dim Fuente As String
        If C.IdTipoCondicion <> 3 Then
            If (C.IdTipoCondicion = 1) Then
                Fuente = "FSU"
            Else
                Fuente = "IE"
            End If
            Return New ConditionTreeNode(Fuente, C.Operando1, C.Operador, C.Operando2)
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
            Return root
        End If
    End Function
    Function GetConnection() As SqlConnection
        Dim SqlConn As New SqlConnection
        SqlConn.ConnectionString = ConnectionString
        SqlConn.Open()
        Return SqlConn
    End Function
    Function GetFichas(ByVal IdLevantamiento As Integer) As SqlDataReader
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryFichas + Convert.ToString(IdLevantamiento), SqlConn)
        Dim Reader As SqlDataReader = Command.ExecuteReader
        Return Reader
    End Function
    Function GetFormulasFromPrograma(ByVal IdPrograma As Integer) As ArrayList
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryFormulas + Convert.ToString(IdPrograma), SqlConn)
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
    Function GetCondiciones(ByVal IdVariable As Integer) As ArrayList
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryCondiciones + Convert.ToString(IdVariable), SqlConn)
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
    Function GetCondicionesRaiz(ByVal IdVariable As Integer) As Condicion
        Dim SqlConn As SqlConnection
        SqlConn = GetConnection()
        Dim Command As New SqlCommand(QueryRaiz + Convert.ToString(IdVariable), SqlConn)
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
End Module
