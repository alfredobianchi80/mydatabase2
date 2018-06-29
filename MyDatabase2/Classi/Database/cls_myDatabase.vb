Public Class cls_myDatabase
    Implements IDisposable


#Region "Enum"

    Public Enum en_TipoDB
        ACCESS = 0
        MSSQL = 1
        MYSQL = 2
        DBF = 3
        DB2 = 4
        ODBC = 10
        UNDEF = -1
    End Enum

#End Region '"Enum"



#Region "Variabili Private"

    Private _DBConnection As System.Data.Common.DbConnection = Nothing
    Private _ClasseDB As String = ""
    Private _ClasseInizializzata As Boolean = False

    Private _DBName As String = ""
    Private _DBServerPath As String = ""
    Private _TipoDatabase As en_TipoDB = en_TipoDB.UNDEF

    Private _StringConnessione As String = ""

    Private _USerName As String = ""
    Private _Password As String = ""
    Private _Trusted As Boolean = False

    Private _ProviderFactory As System.Data.Common.DbProviderFactory = Nothing

#End Region '"Variabili Private"


#Region "Property"

    Public Property ConnessioneDB() As System.Data.Common.DbConnection
        Get
            Return _DBConnection
        End Get
        Set(ByVal value As System.Data.Common.DbConnection)
            _DBConnection = value
            '_ClasseDB =ctype(_DBConnection)
        End Set
    End Property

    Public WriteOnly Property NomeDatabase
        Set(value)
            _DBName = value
        End Set
    End Property

    Public WriteOnly Property PercorsoDatabase
        Set(value)
            _DBServerPath = value
        End Set
    End Property

    Public WriteOnly Property NomeServer
        Set(value)
            _DBServerPath = value
        End Set
    End Property

    Public ReadOnly Property ClasseDB
        Get
            Return _ClasseDB
        End Get
    End Property

    Public WriteOnly Property UserID As String
        Set(value As String)
            _USerName = value
        End Set
    End Property

    Public WriteOnly Property UserPWD As String
        Set(value As String)
            _Password = value
        End Set
    End Property

    Public WriteOnly Property TrustedConnection As Boolean
        Set(value As Boolean)
            _Trusted = value
        End Set
    End Property

    Public Property StringConnessione
        Set(value)
            If _TipoDatabase < en_TipoDB.ODBC Then
                Throw New System.Exception("Stringa di connessione specificabile solo per TipoDB = ODBC")
            Else
                _StringConnessione = value
            End If
        End Set
        Get
            Return _StringConnessione
        End Get
    End Property

    Public Property TipoDatabase As en_TipoDB
        Get
            Return _TipoDatabase
        End Get
        Set(value As en_TipoDB)
            _TipoDatabase = value
            _ClasseDB = ""
            Select Case value
                Case en_TipoDB.UNDEF
                    _ClasseDB = ""
                Case en_TipoDB.ACCESS
                    _ClasseDB = "System.Data.OleDb" '.OleDbConnection"
                Case en_TipoDB.MSSQL
                    _ClasseDB = "System.Data.SqlClient" '.SqlConnection"
                Case en_TipoDB.MYSQL
                    _ClasseDB = ""
                Case en_TipoDB.DBF
                    _ClasseDB = "System.Data.OleDb" '.OleDbConnection"
                Case en_TipoDB.DB2
                    _ClasseDB = "IBM.Data.DB2" '.DB2Connection"
                Case en_TipoDB.ODBC
                    _ClasseDB = "System.Data.OleDb" '.OleDbConnection"

            End Select
        End Set



    End Property

#End Region '"Property"

#Region "Costruttore"

    Public Sub New()
        _DBConnection = Nothing
        _ClasseDB = ""
        _ClasseInizializzata = False
        _DBName = ""
        _DBServerPath = ""

        _Trusted = False
        _USerName = ""
        _Password = ""

        _TipoDatabase = en_TipoDB.UNDEF
    End Sub

#End Region '"Costruttore"

#Region "Funzioni"

    Public Function Inizializza() As Boolean
        Dim bool_Return As Boolean = False
        Dim obj_ProviderFactory As System.Data.Common.DbProviderFactory = Nothing

        'Chiudi eventuali connessioni già aperte
        Try
            If Not (_DBConnection Is Nothing) Then
                _DBConnection.Close()
                _DBConnection = Nothing
            End If
        Catch ex As Exception

        End Try

        'Crea Classe accesso a Database
        Try
            'Crea Classe
            'obj_ProviderFactory = System.Data.Common.DbProviderFactories.GetFactory("System.Data.SqlClient") '_ClasseDB)
            obj_ProviderFactory = System.Data.Common.DbProviderFactories.GetFactory(_ClasseDB)
            _DBConnection = obj_ProviderFactory.CreateConnection
            _ClasseInizializzata = True
            bool_Return = True
        Catch ex As Exception
            Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Return bool_Return
    End Function


    Public Function FullDBPath() As String
        Return _DBServerPath & IIf(_DBServerPath.EndsWith("\"), "", "\") & _DBName
    End Function

    Public Function Connetti() As Boolean
        Dim bool_Return As Boolean = False
        Dim bool_Errore As Boolean = False


        'Chiudi eventuali connessioni già aperte
        Try
            If Not (_DBConnection Is Nothing) Then
                _DBConnection.Close()
                _DBConnection = Nothing
            End If
        Catch ex As Exception
            bool_Errore = True
        End Try


        'Crea Classe accesso a Database
        If Not bool_Errore Then
            Try
                _ProviderFactory = System.Data.Common.DbProviderFactories.GetFactory(_ClasseDB)
                _DBConnection = _ProviderFactory.CreateConnection
                _ClasseInizializzata = True

            Catch ex As Exception
                bool_Errore = True
            End Try
        End If

        'Apri connessione a database
        If Not bool_Errore Then

            'crea Stringa di connessione a database

            Dim obj_ConStr As System.Data.Common.DbConnectionStringBuilder = _ProviderFactory.CreateConnectionStringBuilder
            Dim bool_UsaCostruttoreStringa As Boolean = True
            Dim str_ConnectionString As String = ""

            'ACCESS: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=[DB_PATH]\[DB_NAME];Persist Security Info=False;
            'MS SQL: "Server={DB_SERVER};Database={DB_NAME};User ID={USER_ID};Password={USER_PWD};Trusted_Connection=False;"
            'MS SQL (trusted): "Server={DB_SERVER};Database={DB_NAME};Trusted_Connection=True;"
            'DBF: "Provider=vfpoledb;SourceType=DBF;SourceDB={DB_NAME};Persist Security Info=False"
            'DB2: "Server={DB_SERVER}; Database={DB_NAME}; UID={USER_ID}; PWD={USER_PWD}"
            Select Case _TipoDatabase
                Case en_TipoDB.ACCESS
                    Dim str_Percorso As String = ""
                    str_Percorso = FullDBPath() ' _DBServerPath & IIf(_DBServerPath.EndsWith("\"), "", "\") & _DBName
                    obj_ConStr.Add("Provider", "Microsoft.ACE.OLEDB.12.0")
                    obj_ConStr.Add("Data Source", str_Percorso)
                    obj_ConStr.Add("Persist Security Info", "false")
                Case en_TipoDB.DB2
                    obj_ConStr.Add("Server", _DBServerPath)
                    obj_ConStr.Add("Database", _DBName)
                    obj_ConStr.Add("UID", _USerName)
                    obj_ConStr.Add("PWD", _Password)
                    
                Case en_TipoDB.DBF
                    obj_ConStr.Add("Provider", "vfpoledb")
                    obj_constr.add("data source", _dbserverpath)
                    'obj_ConStr.Add("SourceType", "DBF")
                    'obj_ConStr.Add("SourceDB", _DBServerPath)
                    'obj_ConStr.Add("Persist Security Info", False)
                    bool_UsaCostruttoreStringa = False
                    _DBConnection = New System.Data.OleDb.OleDbConnection '_ProviderFactory.CreateConnection
                    str_ConnectionString = "Provider=vfpoledb;data source={DB_PATH}" ';Persist Security Info=False"
                    str_ConnectionString = Replace(str_ConnectionString, "{DB_PATH}", _DBServerPath)

                Case en_TipoDB.MSSQL
                    obj_ConStr.Add("Server", _DBServerPath)
                    obj_ConStr.Add("Database", _DBName)
                    obj_ConStr.Add("Trusted_Connection", _Trusted.ToString)
                    If _Trusted Then
                        'obj_ConStr.Add("User ID", "")
                        'obj_ConStr.Add("Password", "")
                    Else
                        obj_ConStr.Add("User ID", _USerName)
                        obj_ConStr.Add("Password", _Password)
                    End If
                    
                Case en_TipoDB.MYSQL
                    'obj_ConStr.Add("xxx", "")
                    'obj_ConStr.Add("xxx", "")
                    'obj_ConStr.Add("xxx", "")
                    bool_Errore = True
                    Error 10001
                Case en_TipoDB.ODBC
                    obj_ConStr.ConnectionString = _StringConnessione

                Case en_TipoDB.UNDEF
                    bool_Errore = True
                    Error 10000
            End Select

            If Not bool_Errore Then
                Try
                    If bool_UsaCostruttoreStringa Then
                        _StringConnessione = obj_ConStr.ConnectionString
                    Else
                        _StringConnessione = str_ConnectionString
                    End If

                    _DBConnection.ConnectionString = _StringConnessione
                    _DBConnection.Open()
                    bool_Return = True
                Catch ex As Exception
                    bool_Errore = True
                    Debug.Print(ex.Message)
                    MsgBox(ex.Message)
                End Try
            End If
        End If

        If bool_Errore And Not (bool_Return) Then
            bool_Return = False
        End If



        If _DBConnection Is Nothing Then
            Windows.Forms.MessageBox.Show("HELLO")
        End If
        Return bool_Return
    End Function

    Public Function Disconnetti() As Boolean
        Dim bool_Return As Boolean = False


        If (_DBConnection Is Nothing) Then
            'Già disconnesso
            bool_Return = True
        Else
            Try
                _DBConnection.Close()

            Catch ex As Exception
            Finally
                _DBConnection = Nothing
                bool_Return = True
            End Try
        End If
        Return bool_Return
    End Function

    Public Function isInizializzato() As Boolean
        Return _ClasseInizializzata
    End Function

    Public Function isConnessioneAperta() As Boolean
        Dim bool_Return As Boolean = False

        If Not _DBConnection Is Nothing Then

            If _DBConnection.State = Data.ConnectionState.Open Then
                '-- [TODO] Testare se la connessione è effettivamente aperta?!?

                bool_Return = True

            End If
        End If
        Return bool_Return
    End Function

    Public Function CreaCommand() As System.Data.Common.DbCommand
        Try
            Return _DBConnection.CreateCommand
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function CreateDataAdapter() As System.Data.Common.DbDataAdapter
        Dim obj_ProviderFactory As System.Data.Common.DbProviderFactory
        Dim obj_DataAdapter As System.Data.Common.DbDataAdapter
        obj_ProviderFactory = System.Data.Common.DbProviderFactories.GetFactory(_ClasseDB)

        obj_DataAdapter = obj_ProviderFactory.CreateDataAdapter

        Return obj_DataAdapter
    End Function


    Public Function ElencaTabelle() As ArrayList
        Dim obj_SchemaTable As System.Data.DataTable = Nothing
        Dim obj_SchemaSchema As Object = Nothing
        'MessageBox.Show("Inizio Creazione Elenco Tabelle")
        ElencaTabelle = New ArrayList

        'Ottiene tutte le tabelle
        Try
            Select Case _TipoDatabase
                Case en_TipoDB.ACCESS, en_TipoDB.DBF
                    Dim obj_cn As System.Data.OleDb.OleDbConnection = DirectCast(_DBConnection, System.Data.OleDb.OleDbConnection)
                    obj_SchemaTable = obj_cn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, Nothing})
                Case en_TipoDB.DB2
                    Dim obj_cn As IBM.Data.DB2.DB2Connection = DirectCast(_DBConnection, IBM.Data.DB2.DB2Connection)
                    'Dim obj_cn As Object = DirectCast(_DBConnection, IBM.Data.DB2.DB2Connection)
                    obj_SchemaSchema = obj_cn.GetSchema(IBM.Data.DB2.DB2MetaDataCollectionNames.Schemas)
                    obj_SchemaTable = obj_cn.GetSchema(IBM.Data.DB2.DB2MetaDataCollectionNames.Tables)

                    obj_SchemaSchema = Nothing
                Case en_TipoDB.DBF
                    Dim obj_cn As System.Data.OleDb.OleDbConnection = DirectCast(_DBConnection, System.Data.OleDb.OleDbConnection)
                    obj_SchemaTable = obj_cn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, Nothing})

                Case en_TipoDB.MSSQL
                    Dim obj_cn As System.Data.SqlClient.SqlConnection = DirectCast(_DBConnection, System.Data.SqlClient.SqlConnection)
                    'obj_SchemaTable = obj_cn.GetSchema(IBM.Data.DB2.DB2MetaDataCollectionNames.Tables)
                    obj_SchemaTable = obj_cn.GetSchema(System.Data.SqlClient.SqlClientMetaDataCollectionNames.Tables)



                Case Else
                    Windows.Forms.MessageBox.Show("[ElencaTabelle] TIPO DB non gestito")
            End Select
        Catch ex As Exception
            Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        'MessageBox.Show("Fine Creazione Elenco Tabelle")


        If obj_SchemaSchema IsNot Nothing Then
            Dim str_NomeSchema As String = ""
            Beep()
            For I As Int16 = 0 To obj_SchemaSchema.Rows.Count - 1
                str_NomeSchema = Convert.ToString(obj_SchemaSchema.Rows.item(I).item(0))
                For Each obj_x In obj_SchemaSchema.rows(I).table.rows.item(1)
                    Beep()
                Next

                If Convert.ToString(obj_SchemaTable.Rows(I).Item(3)).ToString Like "*TABLE*" Then
                    ElencaTabelle.Add(obj_SchemaTable.Rows(I).Item(2))
                End If
            Next
        Else
            If Not (obj_SchemaTable Is Nothing) Then
                For I As Int16 = 0 To obj_SchemaTable.Rows.Count - 1
                    If Convert.ToString(obj_SchemaTable.Rows(I).Item(3)).ToString Like "*TABLE*" Then
                        ElencaTabelle.Add(obj_SchemaTable.Rows(I).Item(2))
                    End If
                Next

            End If
        End If

        ''Le aggiunge alla collezione
        'If Not (obj_SchemaTable Is Nothing) Then
        '    For I As Int16 = 0 To obj_SchemaTable.Rows.Count - 1
        '        If Convert.ToString(obj_SchemaTable.Rows(I).Item(3)).ToString Like "*TABLE*" Then
        '            ElencaTabelle.Add(obj_SchemaTable.Rows(I).Item(2))
        '        End If
        '    Next

        'End If
    End Function

#End Region '"Funzioni"


#Region "TEST"


    Public Function GetClassName() As String
        Try
            'Return ctype(_DBConnection)
            'Return GetType(_DBConnection)
            'Return TypeName(_DBConnection)
            Return VbTypeName(TypeName(_DBConnection))
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

    Sub p()


        Dim factory As System.Data.Common.DbProviderFactory
        factory = System.Data.Common.DbProviderFactories.GetFactory("System.Data.SqlClient")

        Dim conn As System.Data.Common.DbConnection = factory.CreateConnection()
        conn.ConnectionString = "connectionString"
        conn.Open()


    End Sub

#End Region '"TEST"


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
