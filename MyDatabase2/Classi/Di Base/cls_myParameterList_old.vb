Public Class cls_myParameterList_old


#Region "Elementi Privati"

    Protected _ListaParametri As Dictionary(Of String, System.Data.Common.DbParameter) = Nothing
    Private _DBConnection As System.Data.Common.DbConnection

#End Region '"Elementi Privati"


#Region "Costruttore"

    Public Sub New(DBConnection As System.Data.Common.DbConnection)
        _ListaParametri = New Dictionary(Of String, Data.Common.DbParameter)
        _DBConnection = DBConnection
    End Sub

#End Region '"Costruttore"


#Region "Property"

    Default Public Property Item(ByVal Nome As String) As System.Data.Common.DbParameter
        Get
            Nome = Nome.ToUpper.Trim
            If _ListaParametri.ContainsKey(Nome) Then
                Return _ListaParametri.Item(Nome)
            Else
                Return Nothing
            End If
        End Get

        Set(value As System.Data.Common.DbParameter)
            Nome = Nome.ToUpper.Trim
            If _ListaParametri.ContainsKey(Nome) Then
                value.ParameterName = Nome
                _ListaParametri(Nome) = value
            Else
                _ListaParametri.Add(Nome, value)
            End If
        End Set
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return _ListaParametri.Count
        End Get
    End Property

    Public ReadOnly Property Items As List(Of System.Data.Common.DbParameter)
        Get
            Dim obj_Lista As New List(Of System.Data.Common.DbParameter)
            For Each obj_x As KeyValuePair(Of String, System.Data.Common.DbParameter) In _ListaParametri
                obj_Lista.Add(obj_x.Value)
            Next
            Return obj_Lista
        End Get
    End Property

#End Region '"Property"


#Region "Metodi Pubblici"

    Public Function Clear() As Boolean
        Dim bool_Return As Boolean = False
        Try
            _ListaParametri.Clear()
            bool_Return = True
        Catch ex As Exception
            bool_Return = False
        End Try
        Return bool_Return
    End Function

    Public Function ContainsKey(ByVal Key As String) As Boolean
        Key = Key.ToUpper.Trim
        Try
            Return _ListaParametri.ContainsKey(Key)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Add(ByVal Nome As String, ByVal Valore As Object, ByVal Tipo As System.Data.DbType) As Boolean 'System.Data.Common.DbParameter
        Dim bool_Return As Boolean = False
        Nome = Nome.ToUpper.Trim
        Dim obj_RetValue As System.Data.Common.DbParameter

        Try
            obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
            With obj_RetValue
                .ParameterName = Nome
                .Value = Valore
                .DbType = Tipo
            End With

        Catch ex As Exception
            obj_RetValue = Nothing
            bool_Return = False
        End Try

        If Not (obj_RetValue Is Nothing) Then
            If InizializzaListaParametri() Then
                Try
                    _ListaParametri.Add(Nome, obj_RetValue)
                    bool_Return = True
                Catch ex As Exception
                    bool_Return = False
                End Try

            Else
                bool_Return = False
            End If
        End If

        Return bool_Return
    End Function

    Public Function Delete(ByVal Nome As String) As Boolean
        Dim bool_Return As Boolean = False
        Nome = Nome.ToUpper.Trim

        Try
            bool_Return = _ListaParametri.Remove(Nome)
        Catch ex As Exception
            bool_Return = False
        End Try

        Return bool_Return
    End Function

#End Region '"Metodi Pubblici"


#Region "Metodi Privati"

    Private Function InizializzaListaParametri() As Boolean
        Dim bool_RetValue As Boolean = False

        If _ListaParametri Is Nothing Then
            Try
                _ListaParametri = New Dictionary(Of String, Data.Common.DbParameter)
                bool_RetValue = True
            Catch ex As Exception
                bool_RetValue = False
            End Try
        Else
            bool_RetValue = True
        End If

        Return bool_RetValue
    End Function

#End Region '"Metodi Privati"

End Class


