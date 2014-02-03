Imports System.Data.SqlClient

Module Module1
    Dim ConnectionString As String = "Data Source=localhost;Initial Catalog=SUEPPS-BD;Integrated Security=True"
    Sub Main()
        Dim F As New Ficha(58, ConnectionString)
        F.PrintFullFicha()
        'PrevTest()
        'TestCheckFichas(133, 1)
        'Programa 133 que tiene asociados dos indicadores, 5 levantamiento que está en la bd de prueba, sólo
        'para que pase el FK constraint de ValoresIndicadores
        'Dim Calculadora As New CalculadoraIndicadores(ConnectionString, 133, 6)
        'Calculadora.Run("arias-test")
        Console.ReadKey()

    End Sub


End Module
