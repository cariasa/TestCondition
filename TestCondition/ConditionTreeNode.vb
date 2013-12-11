Public Class ConditionTreeNode
    Private NodeLevel As Char
    'NodeLevel determina qué nivel de ficha es necesario recuperar para FSU: V vivienda, H Hogar, M Miembros
    Private NodeType As Char
    'NodeType = 'Q' o 'C' Q representa una pregunta, C representa una condicion
    Private LogicalOperation As Char
    'LogicalOperation = {Y, O}
    Private LogicalValue As Boolean

    Private Source As String
    'Source = {FSU, IE}
    Private SourceQName As String
    'SourceQName es el nombre amigable de la pregunta, por ejemplo P11, EM09, etc
    Private QOperation As String
    'QOperation = {=, <, <=, >, >=, C}
    Private QValues As String
    'QValues es la lista de valores posibles que se van a verificar con la pregunta, son separados por coma
    'cuando se usa el operador contenido, de lo contrario sólo debería haber uno
    Private Shared Separator() As Char = {","c, " "c}
    Private Left, Right As ConditionTreeNode
    Public ReadOnly Property Value As Boolean
        Get
            Return LogicalValue
        End Get
    End Property
    Public Property Level As Char
        Get
            Return NodeLevel
        End Get
        Set(ByVal value As Char)
            NodeLevel = value
        End Set
    End Property

    Public Property LeftNode As ConditionTreeNode
        Get
            Return Left
        End Get
        Set(value As ConditionTreeNode)
            Left = value
        End Set
    End Property
    Public Property RightNode As ConditionTreeNode
        Get
            Return Right
        End Get
        Set(value As ConditionTreeNode)
            Right = value
        End Set
    End Property

    Public Sub New(ByVal LogicalOperation As Char)
        Me.NodeType = "C"
        Me.LogicalOperation = LogicalOperation
        Me.LogicalValue = False
        Me.Source = Nothing
        Me.SourceQName = Nothing
        Me.QOperation = Nothing
        Me.QValues = Nothing
        Me.Left = Nothing
        Me.Right = Nothing
    End Sub

    Public Sub New(ByVal Source As String, ByVal SourceQName As String, ByVal QOperation As String, ByVal QValues As String)
        Me.NodeType = "Q"
        Me.LogicalOperation = Nothing
        Me.LogicalValue = False
        Me.Source = Source
        Me.SourceQName = SourceQName
        Me.QOperation = QOperation
        Me.QValues = QValues
        Me.Left = Nothing
        Me.Right = Nothing
    End Sub
    Public Sub New(ByVal Source As String, ByVal SourceQName As String, ByVal QOperation As String, ByVal QValues As String, ByVal Level As Char)
        Me.NodeType = "Q"
        Me.LogicalOperation = Nothing
        Me.LogicalValue = False
        Me.Source = Source
        Me.SourceQName = SourceQName
        Me.QOperation = QOperation
        Me.QValues = QValues
        Me.Left = Nothing
        Me.Right = Nothing
        Me.Level = Level
    End Sub
    Public Overrides Function ToString() As String
        Return NodeType + " " + LogicalOperation + " " + Source + " " + SourceQName + " " + QOperation + " " + QValues
    End Function
    Public Sub PrintTree()
        PrintTree(Me)
    End Sub
    Private Sub PrintTree(ByRef Node As ConditionTreeNode, Optional ByVal Espacio As String = "")
        Console.WriteLine(Node.ToString)
        If Not Node.Left Is Nothing Then
            PrintTree(Node.Left, vbTab)
        End If
        If Not Node.Right Is Nothing Then
            PrintTree(Node.Right, vbTab)
        End If
    End Sub
    'Hay que pasar como parametro la Ficha y el Instrumento de Evaluación a verificar
    Public Function Evaluate(ByRef FSU As FichaSU) As Boolean
        If NodeType = "C" Then
            Dim opLeft, opRight As Boolean
            opLeft = Left.Evaluate(FSU)
            opRight = Right.Evaluate(FSU)
            If LogicalOperation = "Y" Then
                Return opLeft And opRight
            Else
                Return opLeft Or opRight
            End If
        Else
            If Source = "FSU" Then
                'Se verifica SourceQName para sacar el campo que viene de FSU
                'Luego se ve si el valor del campo que viene de FSU tiene el valor guardado en QValues de ser así
                'retorna true, de lo contrario false
                Dim Respuesta As Integer
                Respuesta = FSU.GetValorRespuestaUnica(SourceQName)
                'QOperation = {=, <, <=, >, >=, C}
                If QOperation = "=" Then
                    If Convert.ToUInt64(QValues) = Respuesta Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf QOperation = "<" Then
                    If Respuesta < Convert.ToUInt64(QValues) Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf QOperation = "<=" Then
                    If Respuesta <= Convert.ToUInt64(QValues) Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf QOperation = ">" Then
                    If Respuesta > Convert.ToUInt64(QValues) Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf QOperation = ">=" Then
                    If Respuesta >= Convert.ToUInt64(QValues) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Dim values As String()
                    values = QValues.Split(Separator)
                    For Each v In values
                        If Convert.ToInt64(v) = Respuesta Then
                            Return True
                        End If
                    Next
                    Return False
                End If
            Else
                'Parecido al anterior pero se verifica el Instrumento de Evaluacion IE
            End If
            Return False
        End If
    End Function
End Class
