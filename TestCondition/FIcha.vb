Imports System.Data.SqlClient

Public Class Ficha
    Private Ficha As FichaVivienda
    Private IdFicha As Integer
    Private ConnectionString As String

    Public Sub New(ByVal IdFicha As Integer, ByVal ConnectionString As String)
        Me.IdFicha = IdFicha
        Me.ConnectionString = ConnectionString
        Ficha = New FichaVivienda
        RetrieveVivienda()
        RetrieveViviendaHogares()
        RetrieveHogarMiembros()
    End Sub
    Private Function GetConnection() As SqlConnection
        Dim SqlConn As New SqlConnection
        SqlConn.ConnectionString = ConnectionString
        SqlConn.Open()
        Return SqlConn
    End Function
    Private Sub RetrieveVivienda()
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("ViviendaPorFicha", SqlConn)
        Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader
        Reader = Command.ExecuteReader
        Dim Schema As DataTable
        Schema = Reader.GetSchemaTable
        Dim IdVivienda As Integer
        While Reader.Read
            For Each Field As DataRow In Schema.Rows
                If Field("ColumnName").ToString.Equals("IdVivienda") Then
                    IdVivienda = Reader(Field("ColumnName"))
                Else
                    Ficha.SetValorRespuestaUnica(Field("ColumnName"), Reader(Field("ColumnName")))
                End If
            Next
            Ficha.IdVivienda = IdVivienda
            ' Preguntas de Opción Múltiple
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

    End Sub
    Private Sub RetrieveViviendaHogares()
        Dim ListaHogares As New ArrayList
        Dim SqlConn As SqlConnection = GetConnection()
        Dim Command As New SqlCommand("HogaresPorVivienda", SqlConn)
        Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
        Command.Parameters.AddWithValue("@IdVivienda", Ficha.IdVivienda)
        Command.CommandType = CommandType.StoredProcedure
        Dim Reader As SqlDataReader
        Reader = Command.ExecuteReader
        Dim Schema As DataTable
        Schema = Reader.GetSchemaTable
        Dim IdHogar As Integer
        Dim FichaH As New FichaHogar
        While Reader.Read
            For Each Field As DataRow In Schema.Rows
                If Field("ColumnName").ToString.Equals("IdHogar") Then
                    IdHogar = Reader(Field("ColumnName"))
                ElseIf Field("ColumnName").ToString.Equals("H7") Then
                    'Se salta el campo H7, Centro de Salud, pues es un campo String.
                Else
                    FichaH.SetValorRespuestaUnica(Field("ColumnName"), Reader(Field("ColumnName")))
                End If
            Next
            FichaH.IdHogar = IdHogar

            Dim SqlConn2 As SqlConnection = GetConnection()
            Dim Command2 As New SqlCommand("OrganizacionesComunitariasPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Dim Reader2 As SqlDataReader
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                FichaH.AddValorRespuestaMultiple("H4", Reader2("H4"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("BienesPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                FichaH.AddValorRespuestaMultiple("H5", Reader2("H5"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            SqlConn2 = GetConnection()
            Command2 = New SqlCommand("ServiciosFinancierosPorHogar", SqlConn2)
            Command2.Parameters.AddWithValue("@IdHogar", Reader("IdHogar"))
            Command2.CommandType = CommandType.StoredProcedure
            Reader2 = Command2.ExecuteReader
            While Reader2.Read
                FichaH.AddValorRespuestaMultiple("H11", Reader2("H11"))
            End While
            Reader2.Close()
            SqlConn2.Close()

            ListaHogares.Add(FichaH)


        End While
        Reader.Close()
        SqlConn.Close()
        Ficha.SetHogares(ListaHogares)
    End Sub

    Private Sub RetrieveHogarMiembros()

        For Each Hogar As FichaHogar In Ficha.GetHogaresEnVivienda
            Dim ListaMiembros As New ArrayList
            Dim SqlConn As SqlConnection = GetConnection()
            Dim Command As New SqlCommand("PoblacionPorHogar", SqlConn)
            Command.Parameters.AddWithValue("@CodigoFSU", IdFicha)
            Command.Parameters.AddWithValue("@IdVivienda", Ficha.IdVivienda)
            Command.Parameters.AddWithValue("@IdHogar", Hogar.IdHogar)
            Command.CommandType = CommandType.StoredProcedure
            Dim Reader As SqlDataReader
            Reader = Command.ExecuteReader
            Dim Schema As DataTable
            Schema = Reader.GetSchemaTable
            Dim IdMiembro As Integer
            While Reader.Read
                Dim FichaM As New FichaMiembro
                For Each Field As DataRow In Schema.Rows
                    If Field("ColumnName").ToString.Equals("IdMiembro") Then
                        IdMiembro = Reader(Field("ColumnName"))
                    Else
                        FichaM.SetValorRespuestaUnica(Field("ColumnName"), Reader(Field("ColumnName")))
                    End If
                Next
                FichaM.IdMiembro = IdMiembro
                Dim SqlConn3 As SqlConnection = GetConnection()
                Dim Command3 As New SqlCommand("DiscapacidadesPorMiembro", SqlConn3)
                Command3.Parameters.AddWithValue("@IdMiembro", Reader("IdMiembro"))
                Command3.CommandType = CommandType.StoredProcedure
                Dim Reader3 As SqlDataReader
                Reader3 = Command3.ExecuteReader
                While Reader3.Read
                    FichaM.AddValorRespuestaMultiple("P20", Reader3("P20"))
                End While
                Reader3.Close()
                SqlConn3.Close()
                SqlConn3 = GetConnection()
                Command3 = New SqlCommand("ProgramasSocialesPorMiembro", SqlConn3)
                Command3.Parameters.AddWithValue("@IdMiembro", Reader("IdMiembro"))
                Command3.CommandType = CommandType.StoredProcedure
                Reader3 = Command3.ExecuteReader
                While Reader3.Read
                    FichaM.AddValorRespuestaMultiple("P21", Reader3("P21"))
                End While
                Reader3.Close()
                SqlConn3.Close()
                ListaMiembros.Add(FichaM)
            End While
            Hogar.SetMiembros(ListaMiembros)
        Next
    End Sub
    Public Function GetDepartamento() As VariableDepartamento
        Dim VarDepto As New VariableDepartamento
        VarDepto.Departamento = Ficha.GetValorRespuestaUnica("V01")
        Return VarDepto
    End Function
    Public Function GetDepartamentoMunicipio() As VariableDepartamentoMunicipio
        Dim VarDeptoMuni As New VariableDepartamentoMunicipio
        VarDeptoMuni.Departamento = Ficha.GetValorRespuestaUnica("V01")
        VarDeptoMuni.Municipio = Ficha.GetValorRespuestaUnica("V02")
        Return VarDeptoMuni
    End Function
    
    Public Sub PrintFullFicha()
        Console.WriteLine("Id Ficha = " + IdFicha.ToString)
        Console.WriteLine("--------------------------------------------------------------------------------------------")
        Console.WriteLine("Vivienda... IdVivienda = " + Ficha.IdVivienda.ToString)
        Ficha.PrintAllValues()
        Console.WriteLine("--------------------------------------------------------------------------------------------")
        For Each Hogar As FichaHogar In Ficha.GetHogaresEnVivienda
            Console.WriteLine(vbTab + "Hogar... IdHogar = " + Hogar.IdHogar.ToString)
            Hogar.PrintAllValues()
            Console.WriteLine("----------------------------------------------------------------------------------------")
            For Each Miembro As FichaMiembro In Hogar.GetMiembrosEnHogar
                Console.WriteLine(vbTab + vbTab + "Miembro... IdMiembro = " + Miembro.IdMiembro.ToString)
                Miembro.PrintAllValues()
            Next
        Next
    End Sub
End Class
