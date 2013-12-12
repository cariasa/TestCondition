Public Class VariablesDesagregadas
    Private pVariable As String
    Public Property Variable() As String
        Get
            Return pVariable
        End Get
        Set(ByVal value As String)
            pVariable = value
        End Set
    End Property
    Private pIdDepartamento As String
    Public Property IdDepartamento() As String
        Get
            Return pIdDepartamento
        End Get
        Set(ByVal value As String)
            pIdDepartamento = value
        End Set
    End Property
    Private pIdMunicipio As String
    Public Property IdMunicipio() As String
        Get
            Return pIdMunicipio
        End Get
        Set(ByVal value As String)
            pIdMunicipio = value
        End Set
    End Property
    Private pIdRegion As String
    Public Property IdRegion() As String
        Get
            Return pIdRegion
        End Get
        Set(ByVal value As String)
            pIdRegion = value
        End Set
    End Property
    Private pIdAldea As String
    Public Property IdAldea() As String
        Get
            Return pIdAldea
        End Get
        Set(ByVal value As String)
            pIdAldea = value
        End Set
    End Property
    Private pIdCaserio As String
    Public Property IdCaserio() As String
        Get
            Return pIdCaserio
        End Get
        Set(ByVal value As String)
            pIdCaserio = value
        End Set
    End Property
    Private pIdBarrio As String
    Public Property IdBarrio() As String
        Get
            Return pIdBarrio
        End Get
        Set(ByVal value As String)
            pIdBarrio = value
        End Set
    End Property

End Class
Public Class UbicacionComparer
    Implements Generic.IEqualityComparer(Of VariablesDesagregadas)
    Public Overloads Function Equals(a As VariablesDesagregadas, b As VariablesDesagregadas) As Boolean Implements System.Collections.Generic.IEqualityComparer(Of TestCondition.VariablesDesagregadas).Equals
        Return a.IdAldea = b.IdAldea And b.IdBarrio = a.IdBarrio And a.IdCaserio = b.IdCaserio And a.IdDepartamento = b.IdDepartamento And a.IdMunicipio = b.IdMunicipio And a.IdRegion = b.IdRegion And a.Variable.Equals(b.Variable)
    End Function
    Public Overloads Function GetHashCode(obj As VariablesDesagregadas) As Integer Implements System.Collections.Generic.IEqualityComparer(Of TestCondition.VariablesDesagregadas).GetHashCode
        Return obj.IdDepartamento Xor obj.IdMunicipio Xor obj.IdAldea Xor obj.IdCaserio Xor obj.IdRegion Xor obj.IdBarrio Xor obj.Variable.GetHashCode
    End Function
End Class
Public Class VariableDepartamento
    Private pVariable As String
    Public Property Variable() As String
        Get
            Return pVariable
        End Get
        Set(ByVal value As String)
            pVariable = value
        End Set
    End Property
    Private pDepartamento As Integer
    Public Property Departamento() As Integer
        Get
            Return pDepartamento
        End Get
        Set(ByVal value As Integer)
            pDepartamento = value
        End Set
    End Property

End Class
Public Class VariableDepartamentoComparer
    Implements Generic.IEqualityComparer(Of VariableDepartamento)

    Public Shadows Function Equals(x As VariableDepartamento, y As VariableDepartamento) As Boolean Implements IEqualityComparer(Of VariableDepartamento).Equals
        Return x.Departamento = y.Departamento And x.Variable.Equals(y.Variable)
    End Function

    Public Overloads Function GetHashCode(obj As VariableDepartamento) As Integer Implements IEqualityComparer(Of VariableDepartamento).GetHashCode
        Return obj.Departamento Xor obj.Variable.GetHashCode
    End Function
End Class
Public Class VariableDepartamentoMunicipio
    Private pVariable As String
    Public Property Variable() As String
        Get
            Return pVariable
        End Get
        Set(ByVal value As String)
            pVariable = value
        End Set
    End Property
    Private pDepartamento As Integer
    Public Property Departamento() As Integer
        Get
            Return pDepartamento
        End Get
        Set(ByVal value As Integer)
            pDepartamento = value
        End Set
    End Property
    Private pMunicipio As Integer
    Public Property Municipio() As Integer
        Get
            Return pMunicipio
        End Get
        Set(ByVal value As Integer)
            pMunicipio = value
        End Set
    End Property

End Class
Public Class VariableDepartamentoMunicipioComparer
    Implements Generic.IEqualityComparer(Of VariableDepartamentoMunicipio)

    Public Shadows Function Equals(x As VariableDepartamentoMunicipio, y As VariableDepartamentoMunicipio) As Boolean Implements IEqualityComparer(Of VariableDepartamentoMunicipio).Equals
        Return x.Departamento = y.Departamento And x.Municipio = y.Municipio And x.Variable.Equals(y.Variable)
    End Function

    Public Overloads Function GetHashCode(obj As VariableDepartamentoMunicipio) As Integer Implements IEqualityComparer(Of VariableDepartamentoMunicipio).GetHashCode
        Return obj.Departamento Xor obj.Variable.GetHashCode Xor obj.Municipio
    End Function
End Class
Public Class VariableSexo
    Private pVariable As String
    Public Property Variable() As String
        Get
            Return pVariable
        End Get
        Set(ByVal value As String)
            pVariable = value
        End Set
    End Property
    Private pSexo As Integer
    Public Property Sexo() As Integer
        Get
            Return pSexo
        End Get
        Set(ByVal value As Integer)
            pSexo = value
        End Set
    End Property
End Class
Public Class VariableSexoComparer
    Implements Generic.IEqualityComparer(Of VariableSexo)

    Public Shadows Function Equals(x As VariableSexo, y As VariableSexo) As Boolean Implements IEqualityComparer(Of VariableSexo).Equals
        Return x.Sexo = y.Sexo And x.Variable.Equals(y.Variable)
    End Function

    Public Overloads Function GetHashCode(obj As VariableSexo) As Integer Implements IEqualityComparer(Of VariableSexo).GetHashCode
        Return obj.Sexo Xor obj.Variable.GetHashCode
    End Function
End Class