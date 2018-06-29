Public Class cls_myParameterList
    Inherits cls_BaseList(Of System.Data.Common.DbParameter)


#Region "Elementi Privati"

    '    Protected _ListaParametri As Dictionary(Of String, System.Data.Common.DbParameter) = Nothing
    Private _DBConnection As System.Data.Common.DbConnection

#End Region '"Elementi Privati"


#Region "Costruttore"

    Public Sub New(DBConnection As System.Data.Common.DbConnection)
        MyBase.New()
        _DBConnection = DBConnection
    End Sub

#End Region '"Costruttore"


#Region "Property"


#End Region '"Property"


#Region "Metodi Pubblici"


    Public Overrides Function Add(ByRef Valore As Data.Common.DbParameter) As Data.Common.DbParameter
        Try
            MyBase._ListaValori.Add(Valore.ParameterName.ToUpper.Trim, Valore)
            Return Valore
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    Public Overloads Function Add(ByVal Nome As String, ByVal Valore As Object, ByVal Tipo As System.Data.DbType) As System.Data.Common.DbParameter
        Dim bool_Return As Boolean = False
        Nome = Nome.ToUpper.Trim
        Dim obj_RetValue As System.Data.Common.DbParameter = Nothing

        Try
            If _DBConnection Is Nothing Then
                Throw New SystemException("Connessione a DB non inizializzata")
            Else
                obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
            End If

            If Not IsNumeric(Valore) Then
                If Valore.ToString = "" Then
                    Select Case Tipo
                        Case Data.DbType.Date, Data.DbType.DateTime
                            Valore = DBNull.Value
                    End Select
                End If
            End If

            With obj_RetValue
                .ParameterName = Nome
                .Value = Valore
                .DbType = Tipo
            End With

        Catch ex As Exception
            MsgBox(ex.Message)  'obj_RetValue = Nothing
            bool_Return = False
        End Try

        If Not (obj_RetValue Is Nothing) Then
            Try
                If InizializzaListaParametri() Then
                    obj_RetValue = Add(obj_RetValue)
                Else
                    obj_RetValue = Nothing
                End If
            Catch ex As Exception
                obj_RetValue = Nothing
            End Try
        End If

        Return obj_RetValue
    End Function



#End Region '"Metodi Pubblici"


#Region "Metodi Privati"

    Private Function InizializzaListaParametri() As Boolean
        Dim bool_RetValue As Boolean = False

        If MyBase._ListaValori Is Nothing Then
            Try
                MyBase._ListaValori = New Dictionary(Of String, Data.Common.DbParameter)
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



