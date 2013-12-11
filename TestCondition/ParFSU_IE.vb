Public Class ParFSU_IE
    Public CodigoFSU As Integer
    Public CodigoIE As Integer
    Public UsaFSU As Boolean
    Public Sub New(ByVal CodigoFSU As Integer, ByVal CodigoIE As Integer)
        Me.CodigoFSU = CodigoFSU
        Me.CodigoIE = CodigoIE
        Me.UsaFSU = True
    End Sub
    Public Sub New(ByVal CodigoIE As Integer)
        Me.CodigoFSU = 0
        Me.CodigoIE = CodigoIE
        Me.UsaFSU = False
    End Sub
End Class
