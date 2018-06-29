Public MustInherit Class cls_BaseList(Of T)

#Region "Costanti & Variabili protette"
    Protected _ListaValori As Dictionary(Of String, T)
#End Region '"Costanti"

#Region "Costruttore"

    Public Sub New()
        _ListaValori = New Dictionary(Of String, T)

    End Sub

#End Region


#Region "Property"

    Default Public Property Item(ByVal Key As String) As T
        Get
            Key = Key.ToUpper.Trim
            If _ListaValori.ContainsKey(Key) Then
                Return _ListaValori.Item(Key)
            Else
                Return Nothing
            End If
        End Get

        Set(value As T)
            Key = Key.ToUpper.Trim
            If _ListaValori.ContainsKey(Key) Then
                'value.NomeParametro = NomeParametro

                _ListaValori(Key) = value
            Else
                _ListaValori.Add(Key, value)
            End If
        End Set
    End Property

    Public ReadOnly Property Items As List(Of T)
        Get
            Dim obj_Lista As New List(Of T)
            For Each obj_x As KeyValuePair(Of String, T) In _ListaValori
                obj_Lista.Add(obj_x.Value)
            Next
            Return obj_Lista
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return _ListaValori.Count
        End Get
    End Property

#End Region '"Property"


#Region "Property Object - Liste"

    'Public MustOverride ReadOnly Property ListOf() As List(Of T)
    Public ReadOnly Property ListOf() As List(Of T)
        Get
            Dim obj_x As New List(Of T)
            Try
                For Each x As KeyValuePair(Of String, T) In _ListaValori
                    obj_x.Add(x.Value)
                Next
            Catch ex As Exception

            End Try

            Return obj_x
        End Get
    End Property


    'Public MustOverride ReadOnly Property Lista() As Dictionary(Of String, T)
    Public ReadOnly Property Lista As Dictionary(Of String, T)
        Get
            Return _ListaValori
        End Get
    End Property

#End Region


#Region "Base Function"

    Public Function Clear() As Boolean
        Dim bool_Return As Boolean = False
        Try
            _ListaValori.Clear()
            bool_Return = True
        Catch ex As Exception
            bool_Return = False
        End Try
        Return bool_Return
    End Function

    Public Function ContainsKey(ByVal Key As String) As Boolean
        Key = Key.ToUpper.Trim
        Try
            Return _ListaValori.ContainsKey(Key)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Delete(ByVal Key As String) As Boolean
        Dim bool_Return As Boolean = False
        Key = Key.ToUpper.Trim

        Try
            bool_Return = _ListaValori.Remove(Key)
        Catch ex As Exception
            bool_Return = False
        End Try

        Return bool_Return
    End Function

#End Region '"Base Function"


#Region "ADD Function"

    Public MustOverride Function Add(ByRef Valore As T) As T

#End Region '"ADD Function"

End Class
