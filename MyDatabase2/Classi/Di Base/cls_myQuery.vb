Public Class cls_myQuery
    Implements IDisposable

#Region "Enumerati"

    Public Enum en_TipoRisultatoQuery
        Nessuno = 0
        DataTable = 1
        DataReader = 2
        HashTable = 3
        CommandQuery = 4
        BindingSource = 5
        Scalare = 6
    End Enum

#End Region '"Enumerati"

#Region "Variabili Private"
    Protected _CurrentQuery As String = ""
    Protected _ListaParametri As cls_myParameterList
    Protected _CurrentCommand As System.Data.Common.DbCommand

    Protected _obj_myDatabaseClass As cls_myDatabase
    Protected _obj_Connessione As System.Data.Common.DbConnection
    Protected _LastErrorMessage As String = ""



#End Region '"Variabili Private"

#Region "Costruttore"

    Public Sub New()
        'MyBase.New()
        _CurrentQuery = ""
        _obj_Connessione = Nothing
        _ListaParametri = Nothing
        _CurrentCommand = Nothing
    End Sub

    Public Sub New(ConnessioneDB As cls_myDatabase)
        _CurrentQuery = ""
        _obj_myDatabaseClass = ConnessioneDB

        Try
            _obj_Connessione = _obj_myDatabaseClass.ConnessioneDB
            _ListaParametri = New cls_myParameterList(_obj_Connessione)
        Catch ex As Exception
            _obj_Connessione = Nothing

        End Try

        _CurrentCommand = Nothing
    End Sub


    Public Sub New(ConnessioneDB As System.Data.Common.DbConnection)
        _CurrentQuery = ""
        _obj_myDatabaseClass = Nothing
        _obj_Connessione = ConnessioneDB
        _ListaParametri = New cls_myParameterList(_obj_Connessione)
        _CurrentCommand = Nothing
    End Sub

#End Region '"Costruttore"

#Region "Property"

    Public Property Parametri As cls_myParameterList
        Get
            If _ListaParametri Is Nothing Then
                _ListaParametri = New cls_myParameterList(_obj_Connessione)
            End If
            Return _ListaParametri
        End Get
        Set(value As cls_myParameterList)
            _ListaParametri = value
        End Set
    End Property

    Public Property MyDBClass As cls_myDatabase
        Get
            Return _obj_myDatabaseClass
        End Get
        Set(value As cls_myDatabase)
            _obj_myDatabaseClass = value

            _obj_Connessione = _obj_myDatabaseClass.ConnessioneDB
            _ListaParametri = New cls_myParameterList(_obj_Connessione)
        End Set
    End Property

    'Public Property DBConnection As System.Data.Common.DbConnection
    '    Get
    '        Return _obj_Connessione
    '    End Get
    '    Set(value As System.Data.Common.DbConnection)
    '        _obj_Connessione = value
    '        _obj_myDatabaseClass = Nothing
    '        _ListaParametri = New cls_myParameterList(_obj_Connessione)
    '    End Set
    'End Property

    Public Overridable Property Query() As String
        Get
            Return _CurrentQuery
        End Get
        Set(ByVal value As String)
            _CurrentQuery = value
        End Set
    End Property

    Public ReadOnly Property IsError As Boolean
        Get
            Return (_LastErrorMessage <> "")
        End Get
    End Property

    Public ReadOnly Property Errore As String
        Get
            Return _LastErrorMessage
        End Get
    End Property


#End Region '"Property"

#Region "funzioni"

    Protected Function CreaCommand(ByVal TestoQuery As String) As System.Data.Common.DbCommand
        Dim obj_Return As System.Data.Common.DbCommand = Nothing
        Try
            If Not (_obj_myDatabaseClass Is Nothing) Then
                obj_Return = _obj_myDatabaseClass.CreaCommand()
            Else
                obj_Return = _obj_Connessione.CreateCommand()
            End If

            If Not (obj_Return Is Nothing) Then
                With obj_Return
                    If .Connection Is Nothing Then
                        .Connection = _obj_Connessione
                    End If

                    If Query.Length > 0 Then
                        .CommandText = TestoQuery
                        If _ListaParametri.Count > 0 Then
                            With obj_Return
                                For Each obj_param As System.Data.Common.DbParameter In _ListaParametri.Items
                                    .Parameters.Add(obj_param)
                                Next
                            End With
                        End If
                    End If
                End With
            End If
        Catch ex As Exception
            obj_Return = Nothing
        End Try

        Return obj_Return
    End Function

#Region "ESEGUI QUERY"

    Public Overridable Function eseguiQuery(TestoQuery As String, ListaParametri As List(Of System.Data.Common.DbParameter),
                                        Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Dim obj_RetValue As Object = Nothing

        '-- Ricrea Elenco Parametri
        If Not (ListaParametri Is Nothing) Then
            _ListaParametri.Clear()
            For Each x As System.Data.Common.DbParameter In ListaParametri
                If x Is Nothing Then
                    MsgBox("[cls_myquery.EseguiQuery] Lista Parametri nulla")
                Else
                    _ListaParametri.Add(x.ParameterName, x.Value, x.DbType)
                End If

            Next
        End If

        '-- Memorizza Query da eseguire
        Query = TestoQuery

        '-- Crea Command
        _CurrentCommand = CreaCommand(Query)

        '-- Esegui Query in base al risultato
        If Query.Length > 0 Then
            obj_RetValue = Esegui(_CurrentCommand, TipoRisultato)
        Else
            obj_RetValue = Nothing
            _LastErrorMessage = "Query non specificata"
        End If

        Return obj_RetValue
    End Function

    Public Overridable Function eseguiQuery(ListaParametri As List(Of System.Data.Common.DbParameter), Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Return eseguiQuery(Query, ListaParametri, TipoRisultato)
    End Function

    Public Overridable Function eseguiQuery(TestoQuery As String, Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Return eseguiQuery(TestoQuery, Nothing, TipoRisultato)
    End Function

    Public Overridable Function eseguiQuery(Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Return eseguiQuery(Query, Nothing, TipoRisultato)
    End Function


#End Region '"ESEGUI QUERY"

#Region "SPECIFICI ESEGUI QUERY"

    Protected Function Esegui(ByVal _DBCommand As System.Data.Common.DbCommand, Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Dim obj_RetValue As Object = Nothing

        If Not (_DBCommand Is Nothing) Then
            Select Case TipoRisultato
                Case en_TipoRisultatoQuery.DataReader
                    obj_RetValue = EseguiQuery_DataReader(_DBCommand)
                Case en_TipoRisultatoQuery.DataTable
                    obj_RetValue = EseguiQuery_DataTable(_DBCommand)
                Case en_TipoRisultatoQuery.HashTable
                    obj_RetValue = EseguiQuery_HashTable(_DBCommand)
                Case en_TipoRisultatoQuery.CommandQuery
                    obj_RetValue = EseguiQueryComando(_DBCommand)
                Case en_TipoRisultatoQuery.BindingSource
                    obj_RetValue = EseguiBindingSource(_DBCommand)
                Case en_TipoRisultatoQuery.Nessuno
                    obj_RetValue = EseguiQueryComando(_DBCommand)
                Case en_TipoRisultatoQuery.Scalare
                    obj_RetValue = EseguiQueryScalare(_DBCommand)
            End Select
        End If

        Return obj_RetValue
    End Function

    Protected Function EseguiQuery_DataReader(dataCommand As System.Data.Common.DbCommand) As System.Data.Common.DbDataReader
        Dim dataReader As System.Data.Common.DbDataReader = Nothing
        'Dim dataCommand As System.Data.Common.DbCommand = CreaCommandParam(TestoQuery, SequenzaParametri)
        _LastErrorMessage = ""
        If Not (dataCommand Is Nothing) Then
            'Esegui Query
            Try
                dataReader = dataCommand.ExecuteReader()
            Catch ex As Exception
                _LastErrorMessage = ex.Message
                dataReader = Nothing
            End Try
        End If

        Return dataReader
    End Function

    Protected Function EseguiQuery_DataTable(dataCommand As System.Data.Common.DbCommand) As System.Data.DataTable
        Dim dataReader As System.Data.Common.DbDataReader
        Dim obj_DataTable = New System.Data.DataTable
        _LastErrorMessage = ""
        dataReader = EseguiQuery_DataReader(dataCommand)
        'Popola DataTable
        If Not (dataReader Is Nothing) Then
            obj_DataTable.Load(dataReader)

            'Chiudi datareader
            dataReader.Close()
        Else
            obj_DataTable = Nothing
        End If
        'Ritorna DataTable
        Return obj_DataTable
    End Function

    Protected Function EseguiQueryComando(dataCommand As System.Data.Common.DbCommand) As Integer
        Dim int_Return As Integer = 0
        _LastErrorMessage = ""
        'Esegui Query
        If Not (dataCommand Is Nothing) Then
            Try
                int_Return = dataCommand.ExecuteNonQuery
            Catch ex As Exception
                _LastErrorMessage = ex.Message
                Debug.Print(Query)
                int_Return = -1
            End Try
        End If
        Return int_Return
    End Function

    Protected Function EseguiQueryScalare(dataCommand As System.Data.Common.DbCommand) As Object
        Dim obj_Return As Object = Nothing
        _LastErrorMessage = ""
        'Esegui Query
        If Not (dataCommand Is Nothing) Then
            Try
                obj_Return = dataCommand.ExecuteScalar
            Catch ex As Exception
                _LastErrorMessage = ex.Message
                obj_Return = Nothing
            End Try
        End If
        Return obj_Return
    End Function

    Protected Function EseguiQuery_HashTable(dataCommand As System.Data.Common.DbCommand) As Collections.Hashtable
        Dim dataReader As System.Data.Common.DbDataReader
        Dim obj_hash As New Hashtable
        _LastErrorMessage = ""
        dataReader = EseguiQuery_DataReader(dataCommand)

        If Not (dataReader Is Nothing) Then
            While dataReader.Read
                obj_hash.Add(dataReader(0), dataReader(1))
            End While
            dataReader.Close()
            dataReader = Nothing
        End If

        'Ritorna DataTable
        Return obj_hash
    End Function

    Protected Function EseguiBindingSource(dataCommand As System.Data.Common.DbCommand) As System.Windows.Forms.BindingSource
        Dim dataReader As System.Data.Common.DbDataReader
        Dim obj_DataTable = New System.Data.DataTable
        Dim obj_BindingSource = New System.Windows.Forms.BindingSource

        _LastErrorMessage = ""
        dataReader = EseguiQuery_DataReader(dataCommand)
        'Popola DataTable
        If Not (dataReader Is Nothing) Then
            obj_DataTable.Load(dataReader)

            'Chiudi datareader
            dataReader.Close()
            '.
            obj_BindingSource.DataSource = obj_DataTable
        Else
            obj_DataTable = Nothing
            obj_BindingSource = Nothing
        End If
        'Ritorna DataTable
        Return obj_BindingSource
    End Function

#End Region '"SPECIFICI ESEGUI QUERY"


#End Region '"funzioni"

#Region "Funzioni Shared"

    Public Function isNullValue(ByRef Valore As Object) As Object
        If Valore Is Nothing Then
            Return DBNull.Value

        ElseIf Valore.ToString.Length = 0 Then
            Return DBNull.Value
        Else
            Return Valore
        End If


    End Function

#End Region '"Funzioni Shared"


#Region "IDisposable Support"
    Private disposedValue As Boolean ' Per rilevare chiamate ridondanti

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: eliminare stato gestito (oggetti gestiti).
            End If

            ' TODO: liberare risorse non gestite (oggetti non gestiti) ed eseguire l'override del seguente Finalize().
            ' TODO: impostare campi di grandi dimensioni su null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: eseguire l'override di Finalize() solo se Dispose(ByVal disposing As Boolean) dispone del codice per liberare risorse non gestite.
    'Protected Overrides Sub Finalize()
    '    ' Non modificare questo codice. Inserire il codice di pulizia in Dispose(ByVal disposing As Boolean).
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Questo codice è aggiunto da Visual Basic per implementare in modo corretto il modello Disposable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Non modificare questo codice. Inserire il codice di pulizia in Dispose(disposing As Boolean).
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
