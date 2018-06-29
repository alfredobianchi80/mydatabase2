Public Class cls_myEasyQuery
    Inherits cls_myQuery

#Region "Enumerati e Strutture"

    Private Structure tp_WhereField
        Dim testo As String
        Dim RefParametro As String
        Dim CatenaCondizione As String
    End Structure

    Public Enum en_TipoQuery
        UNDEF = 0
        mySELECT = 1
        myUPDATE = 2
        myINSERT = 4
        myDELETE = 8
    End Enum

#End Region '"Enumerati e Strutture"

#Region "Attributi"

    Protected _NomeTabella As String
    Protected _ListaCampi As cls_myFieldsList
    Protected _ListaWhere As cls_myWhereList
    Protected _ListaOrderBy As cls_myOrderByList
    Protected _ListaGroupBY As cls_myGroupByList
    Protected _ListaHaving As cls_myHavingList
    Protected _ListaTabelleAddizionali As cls_myAdditionalTablesList

    Protected _ListaCampiPersUpdarte As cls_myUpdateFieldList

    Protected _RiordinaParametri As Boolean
    Protected _TipoQuery As en_TipoQuery
    Protected _ListaSequenzaParametri As List(Of String)

#End Region '"Attributi"

#Region "Property"

    Public Property TipoQuery As en_TipoQuery
        Get
            Return _TipoQuery
        End Get
        Set(value As en_TipoQuery)
            _TipoQuery = value
        End Set
    End Property

    Public Property Tabella()
        Get
            Return _NomeTabella
        End Get
        Set(value)
            _NomeTabella = value
        End Set
    End Property

    Public Property OrdinaParametri As Boolean
        Get
            Return _RiordinaParametri
        End Get
        Set(value As Boolean)
            _RiordinaParametri = value
        End Set
    End Property

#End Region '"Property"


#Region "Property Sub-Elementi"

    Public Property Field As cls_myFieldsList
        Get
            If _ListaCampi Is Nothing Then
                _ListaCampi = New cls_myFieldsList
            End If
            Return _ListaCampi
        End Get
        Set(value As cls_myFieldsList)
            _ListaCampi = value
        End Set
    End Property


    Public Property OtherUpdateField As cls_myUpdateFieldList
        Get
            If _ListaCampiPersUpdarte Is Nothing Then
                _ListaCampiPersUpdarte = New cls_myUpdateFieldList
            End If
            Return _ListaCampiPersUpdarte
        End Get
        Set(value As cls_myUpdateFieldList)
            _ListaCampiPersUpdarte = value
        End Set
    End Property


    Public Property Where As cls_myWhereList
        Get
            If _ListaWhere Is Nothing Then
                _ListaWhere = New cls_myWhereList
            End If

            Return _ListaWhere
        End Get
        Set(value As cls_myWhereList)

        End Set
    End Property

    Public Property OrderBy As cls_myOrderByList
        Get
            If _ListaOrderBy Is Nothing Then
                _ListaOrderBy = New cls_myOrderByList
            End If

            Return _ListaOrderBy
        End Get
        Set(value As cls_myOrderByList)

        End Set
    End Property

    Public Property Havinhg As cls_myHavingList
        Get
            If _ListaHaving Is Nothing Then
                _ListaHaving = New cls_myHavingList
            End If

            Return _ListaHaving
        End Get
        Set(value As cls_myHavingList)

        End Set
    End Property

    Public Property GroupBy As cls_myGroupByList
        Get
            If _ListaGroupBY Is Nothing Then
                _ListaGroupBY = New cls_myGroupByList
            End If

            Return _ListaGroupBY
        End Get
        Set(value As cls_myGroupByList)

        End Set
    End Property

    Public Property TabelleAddizionali As cls_myAdditionalTablesList
        Get
            If _ListaTabelleAddizionali Is Nothing Then
                _ListaTabelleAddizionali = New cls_myAdditionalTablesList
            End If
            Return _ListaTabelleAddizionali
        End Get
        Set(value As cls_myAdditionalTablesList)

        End Set
    End Property

#End Region '"Property Sub-Elementi"

#Region "Costruttore"

    Public Sub New()
        MyBase.New()
        '_ListaCampi = New List(Of String)
        _ListaCampi = New cls_myFieldsList
        _ListaWhere = New cls_myWhereList
        _ListaGroupBY = New cls_myGroupByList
        _ListaHaving = New cls_myHavingList
        _ListaOrderBy = New cls_myOrderByList
        _ListaTabelleAddizionali = New cls_myAdditionalTablesList
        _ListaCampiPersUpdarte = New cls_myUpdateFieldList

        _NomeTabella = ""
        _ListaSequenzaParametri = New List(Of String)
        _RiordinaParametri = True
    End Sub

#End Region '"Costruttore"

#Region "Overrides Elements"

    Public Overrides Property Query() As String
        Get
            Dim str_Query As String = ""

            _ListaSequenzaParametri.Clear()

            Select Case _TipoQuery
                Case en_TipoQuery.mySELECT
                    str_Query = CreaQuery_SELECT()

                Case en_TipoQuery.myINSERT
                    str_Query = CreaQuery_INSERT()

                Case en_TipoQuery.myUPDATE
                    str_Query = CreaQuery_UPDATE()

                Case en_TipoQuery.myDELETE
                    str_Query = CreaQuery_DELETE()

                Case Else
                    Throw New System.Exception("Casistica non gestita!!!!")
                    str_Query = ""
            End Select

            _CurrentQuery = str_Query
            Return str_Query
        End Get

        Set(value As String)

        End Set
    End Property

    Public Overrides Function eseguiQuery(Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Dim obj_RetValue As Object = Nothing
        Dim str_TestoQuery As String = Query
        Dim obj_ListaParametri As List(Of System.Data.Common.DbParameter) = Nothing
        Dim obj_param As System.Data.Common.DbParameter = Nothing

        If _ListaSequenzaParametri.Count > 0 Then
            For Each p As String In _ListaSequenzaParametri
                p = p.ToUpper.Trim
                obj_param = MyBase._ListaParametri.Item(p)
                If obj_ListaParametri Is Nothing Then
                    obj_ListaParametri = New List(Of Data.Common.DbParameter)
                End If

                obj_ListaParametri.Add(obj_param)
            Next
        End If
        'qua sopra
        obj_RetValue = MyBase.eseguiQuery(str_TestoQuery, obj_ListaParametri, TipoRisultato)

        Return obj_RetValue
    End Function

#End Region '"Overrides Elements"

#Region "Crea Query Function"

    Private Function CreaQuery_SELECT() As String
        Dim str_Query As String = ""
        Dim str_Fields As String = ""
        Dim str_Where As String = ""
        Dim str_Having As String = ""
        Dim str_GroupBy As String = ""
        Dim str_OrderBy As String = ""
        Dim str_Join As String = ""

        _ListaSequenzaParametri.Clear()

        '** FIELDS
        str_Fields = ""
        For Each o As cls_myField In _ListaCampi.Items
            'For Each o As String In _ListaCampi
            If str_Fields.Length > 0 Then
                str_Fields = str_Fields & ","
            End If
            If o.TabellaRiferimento.Length > 0 Then
                str_Fields = str_Fields & o.TabellaRiferimento.Trim & "."
            End If
            str_Fields = str_Fields & o.NomeCampo

            If o.Etichetta.Length > 0 Then
                str_Fields = str_Fields & " AS " & o.Etichetta
            End If
        Next

        '** WHERE
        str_Where = ""
        For Each obj_Where As cls_myWhere In _ListaWhere.Items
            If str_Where.Length > 0 Then
                str_Where = str_Where & " AND "
            End If
            str_Where = str_Where & " (" & obj_Where.Condizione & ")"

            If obj_Where.RifParametro.ToUpper.Trim.Length > 0 Then
                If Not (_ListaSequenzaParametri.Contains(obj_Where.RifParametro.ToUpper.Trim)) Then
                    _ListaSequenzaParametri.Add(obj_Where.RifParametro.ToUpper.Trim)
                End If
            End If
        Next

        '** GROUP BY
        str_GroupBy = ""
        For Each obj_GroupBy As cls_myGroupBy In _ListaGroupBY.Items
            'For Each o As String In _ListaCampi
            If str_GroupBy.Length > 0 Then
                str_GroupBy = str_GroupBy & ","
            End If
            If obj_GroupBy.TabellaRiferimento.Length > 0 Then
                str_GroupBy = str_GroupBy & obj_GroupBy.TabellaRiferimento.Trim & "."
            End If
            str_GroupBy = str_GroupBy & obj_GroupBy.NomeCampo
        Next

        '** HAVING
        str_Having = ""
        For Each obj_having As cls_myHaving In _ListaHaving.Items
            If str_Having.Length > 0 Then
                str_Having = str_Having & " AND "
            End If
            str_Having = str_Having & " (" & obj_having.Condizione & ")"

            If obj_having.RifParametro.ToUpper.Trim.Length > 0 Then
                If Not (_ListaSequenzaParametri.Contains(obj_having.RifParametro.ToUpper.Trim)) Then
                    _ListaSequenzaParametri.Add(obj_having.RifParametro.ToUpper.Trim)
                End If
            End If
        Next

        '** ORDER BY
        str_OrderBy = ""
        For Each obj_OrderBy As cls_myOrderBy In _ListaOrderBy.Items
            If str_OrderBy.Length > 0 Then
                str_OrderBy = str_OrderBy & " , "
            End If

            If obj_OrderBy.TabellaRiferimento.Length > 0 Then
                str_OrderBy = str_OrderBy & obj_OrderBy.TabellaRiferimento.Trim & "."
            End If
            str_OrderBy = str_OrderBy & obj_OrderBy.NomeCampo

            If obj_OrderBy.Direzione = cls_myOrderBy.en_Direzione.Decrescente Then
                str_OrderBy = str_OrderBy & " DESC"
            End If

        Next

        '** JOIN
        str_Join = _NomeTabella
        If _ListaTabelleAddizionali.Count > 0 Then

            str_Join = ""
            For i As Int32 = 0 To _ListaTabelleAddizionali.Count - 1
                str_Join = str_Join & "("
            Next
            str_Join = str_Join & _NomeTabella

            For Each obj_Join As cls_myAdditionalTables In _ListaTabelleAddizionali.ListOf
                'If str_Join.Length > 0 Then
                '    str_Join = str_Join & " , "
                'End If

                Select Case obj_Join.TipoJoin
                    Case cls_myAdditionalTables.en_TipoJoinTable.INNER
                        str_Join = str_Join & " INNER JOIN "
                    Case cls_myAdditionalTables.en_TipoJoinTable.LEFT
                        str_Join = str_Join & " LEFT JOIN "
                    Case cls_myAdditionalTables.en_TipoJoinTable.RIGHT
                        str_Join = str_Join & " RIGHT JOIN "
                End Select

                str_Join = str_Join & obj_Join.NomeTabella & " ON("
                Dim int_temp As Int32 = 0
                For Each obj_cc As cls_myJoinCondition In obj_Join.Joins.ListOf
                    'str_Join = str_Join & String.Format("{%0}.{%1}={%2}.{%3}", Tabella, obj_cc.NomeCampo_TabellaBase, obj_Join.NomeTabella, obj_cc.NomeCampo_TabellaAggiuntiva)
                    If int_temp > 0 Then
                        str_Join = str_Join & " AND "
                    End If
                    str_Join = str_Join & String.Concat(Tabella, ".", obj_cc.NomeCampo_TabellaBase, "=", obj_Join.NomeTabella, ".", obj_cc.NomeCampo_TabellaAggiuntiva)
                    int_temp += 1
                Next
                str_Join = str_Join & ")"

                'If _ListaTabelleAddizionali.Count > 1 Then
                str_Join = str_Join & ")"
                'End If
            Next
            'Throw New System.Exception("myEasyQuery - Join di tabelle non ancora implementato!!!")
        End If


        '* Componi Query
        str_Query = "SELECT " & str_Fields & " FROM " '& _NomeTabella

        If str_Join.Length > 0 Then
            str_Query = str_Query & str_Join
        End If

        If str_Where.Length > 0 Then
            str_Query = str_Query & " WHERE " & str_Where
        End If

        If str_GroupBy.Length > 0 Then
            str_Query = str_Query & " GROUP BY " & str_GroupBy
        End If

        If str_Having.Length > 0 Then
            str_Query = str_Query & " HAVING " & str_Having
        End If

        If str_OrderBy.Length > 0 Then
            str_Query = str_Query & " ORDER BY " & str_OrderBy
        End If

        Return str_Query
    End Function

    Private Function CreaQuery_INSERT() As String
        Dim str_Query As String = ""
        'Dim str_Fields As String = ""
        'Dim str_Where As String = ""
        'Dim str_Having As String = ""
        'Dim str_GroupBy As String = ""
        'Dim str_OrderBy As String = ""
        Dim str_ListaCampi As String = ""
        Dim str_ListaParametri As String = ""
        _ListaSequenzaParametri.Clear()


        For Each obj_x As System.Data.Common.DbParameter In _ListaParametri.Items

            If str_ListaCampi <> "" Then
                str_ListaCampi = str_ListaCampi & ","
            End If
            str_ListaCampi = str_ListaCampi & obj_x.ParameterName

            If str_ListaParametri <> "" Then
                str_ListaParametri = str_ListaParametri & ","
            End If
            str_ListaParametri = str_ListaParametri & "@" & obj_x.ParameterName

            If Not (_ListaSequenzaParametri.Contains(obj_x.ParameterName.ToUpper.Trim)) Then
                _ListaSequenzaParametri.Add(obj_x.ParameterName.ToUpper.Trim)
            End If


        Next
        str_Query = "INSERT INTO " & _NomeTabella & "(" & str_ListaCampi & ")"
        str_Query = str_Query & " VALUES (" & str_ListaParametri & ")"

        Return str_Query
    End Function

    Private Function CreaQuery_UPDATE() As String
        Dim str_Query As String = ""
        Dim str_Where As String = ""
        Dim str_ListaCampi As String = ""
        _ListaSequenzaParametri.Clear()

        'Inserisci Campi Parametrici
        For Each obj_x As System.Data.Common.DbParameter In _ListaParametri.Items
            If (_ListaCampi.ContainsKey(obj_x.ParameterName)) Then
                If str_ListaCampi <> "" Then
                    str_ListaCampi = str_ListaCampi & ","
                End If
                str_ListaCampi = str_ListaCampi & obj_x.ParameterName & "=@" & obj_x.ParameterName
                'If Not (_ListaSequenzaParametri.Contains(obj_x.ParameterName.ToUpper.Trim)) Then
                _ListaSequenzaParametri.Add(obj_x.ParameterName.ToUpper.Trim)
                'End If
            End If
        Next

        'Inserisci Campi Personalizzati
        For Each obj_x As cls_myUpdateField In _ListaCampiPersUpdarte.Items
            If str_ListaCampi <> "" Then
                str_ListaCampi = str_ListaCampi & ","
            End If
            str_ListaCampi = str_ListaCampi & obj_x.Espressione
        Next

        'Crea Condizione
        str_Where = ""
        For Each obj_Where As cls_myWhere In _ListaWhere.Items
            If str_Where.Length > 0 Then
                str_Where = str_Where & " AND "
            End If
            str_Where = str_Where & " (" & obj_Where.Condizione & ")"

            If obj_Where.RifParametro.ToUpper.Trim.Length > 0 Then
                'If Not (_ListaSequenzaParametri.Contains(obj_Where.RifParametro.ToUpper.Trim)) Then
                _ListaSequenzaParametri.Add(obj_Where.RifParametro.ToUpper.Trim)
                'End If
            End If
        Next
        If str_Where = "" Then
            Throw New System.Exception("Query di Update senza nessuna condizione indicata.")
        End If

        str_Query = "UPDATE " & _NomeTabella & " SET " & str_ListaCampi
        str_Query = str_Query & " WHERE " & str_Where


        Return str_Query
    End Function

    Private Function CreaQuery_DELETE() As String
        Dim str_Query As String = ""
        Dim str_Where As String = ""

        _ListaSequenzaParametri.Clear()

        str_Where = ""
        For Each obj_Where As cls_myWhere In _ListaWhere.Items
            If str_Where.Length > 0 Then
                str_Where = str_Where & " AND "
            End If
            str_Where = str_Where & " (" & obj_Where.Condizione & ")"

            If obj_Where.RifParametro.ToUpper.Trim.Length > 0 Then
                'If Not (_ListaSequenzaParametri.Contains(obj_Where.RifParametro.ToUpper.Trim)) Then
                _ListaSequenzaParametri.Add(obj_Where.RifParametro.ToUpper.Trim)
                'End If
            End If
        Next
        If str_Where = "" Then
            Throw New System.Exception("Query di Update senza nessuna condizione indicata.")
        End If
        Dim str_FDel As String = ""

        If _obj_myDatabaseClass IsNot Nothing Then
            Select Case MyBase._obj_myDatabaseClass.TipoDatabase
                Case cls_myDatabase.en_TipoDB.ACCESS
                    str_FDel = "*"
                Case cls_myDatabase.en_TipoDB.MSSQL
                    str_FDel = ""
            End Select
        Else

            If TypeOf MyBase.MyDBClass.ConnessioneDB Is System.Data.SqlClient.SqlConnection Then
                str_FDel = ""
            End If

            If TypeOf MyBase.MyDBClass.ConnessioneDB Is System.Data.OleDb.OleDbConnection Then
                str_FDel = "*"
            End If
        End If


        str_Query = "DELETE " & str_FDel & " From " & _NomeTabella
        str_Query = str_Query & " WHERE " & str_Where

        Return str_Query
    End Function

#End Region '"Crea Query Function"


#Region "WHERE class"

    Public Class cls_myWhereList
        Inherits cls_BaseList(Of cls_myWhere)

        Public Overrides Function Add(ByRef Valore As cls_myWhere) As cls_myWhere
            If Valore IsNot Nothing Then
                With Valore
                    '    Return Add(.Lingua, .RegolaComposizioneDescrizione, .DescrizioneBase)

                    With Valore
                        Return Add(.Condizione, .RifParametro, .Catena, .NomeCampo)
                    End With

                    'Return Nothing
                End With
            Else
                Return Nothing
            End If
        End Function

        Public Overloads Function Add(ByVal Condizione As String, Optional ByVal RifParametro As String = "", Optional ByVal CatenaCondizione As String = "STD",
                            Optional ByVal NomeCondizione As String = "") As cls_myWhere

            Dim bool_Return As Boolean = False
            Dim obj_RetValue As New cls_myWhere

            'Sistemo Parametri
            Condizione = Condizione.ToUpper.Trim
            RifParametro = RifParametro.ToUpper.Trim
            CatenaCondizione = CatenaCondizione.ToUpper.Trim
            NomeCondizione = NomeCondizione.ToUpper.Trim
            If NomeCondizione = "" Then
                NomeCondizione = Condizione
                'TODO
                'MsgBox("cls_myEasyQuery.Add (NomeCondizione): TODO!!!!")
            End If

            If MyBase._ListaValori.ContainsKey(NomeCondizione) Then
                Throw New System.Exception("Condizione già esistente")
                bool_Return = False
                obj_RetValue = Nothing
            Else
                'Creo Oggetto da inserire
                Try
                    'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                    With obj_RetValue
                        .NomeCampo = NomeCondizione
                        .Condizione = Condizione
                        .RifParametro = RifParametro
                        .Catena = CatenaCondizione
                    End With

                Catch ex As Exception
                    obj_RetValue = Nothing
                End Try

                'Aggiungo Oggetto alla lista
                'If InizializzaListaParametri() Then
                MyBase._ListaValori.Add(NomeCondizione, obj_RetValue)
                'Else
                'bool_Return = false
                'End If
            End If


            Return obj_RetValue
        End Function

    End Class

    Public Class cls_myWhere
        Public Property NomeCampo As String
        Public Property Condizione As String
        Public Property Catena As String
        Public Property RifParametro As String
    End Class

#End Region '"WHERE class"


#Region "FIELD class"

    Public Class cls_myFieldsList
        Inherits cls_BaseList(Of cls_myField)

        Public Overrides Function Add(ByRef Valore As cls_myField) As cls_myField
            If Valore IsNot Nothing Then
                With Valore
                    With Valore
                        Return Add(.NomeCampo, .TabellaRiferimento, .TipoCampo, .Etichetta)
                    End With
                End With
            Else
                Return Nothing
            End If
        End Function

        Public Overloads Function Add(ByVal Nome As String, Optional ByVal Etichetta As String = "") As cls_myField
            Return Add(Nome, "", Data.DbType.Object, Etichetta)
        End Function

        Public Overloads Function Add(ByVal Nome As String, ByVal Tabella As String, ByVal Tipo As System.Data.DbType,
                                      Optional ByVal Etichetta As String = "") As cls_myField
            Dim bool_Return As Boolean = False
            Dim str_ID As String = "2"
            Nome = Nome.ToUpper.Trim
            Dim obj_RetValue As New cls_myField

            Try
                'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                With obj_RetValue
                    .NomeCampo = Nome
                    .TabellaRiferimento = Tabella
                    .TipoCampo = Tipo
                    .Etichetta = Etichetta
                End With

            Catch ex As Exception
                obj_RetValue = Nothing
            End Try

            str_ID = ""
            If Tabella.Trim.Length > 0 Then
                str_ID = Tabella.Trim & "."
            End If
            str_ID = str_ID & Nome.Trim
            str_ID = str_ID.ToUpper

            'If InizializzaListaParametri() Then
            MyBase._ListaValori.Add(str_ID, obj_RetValue)
            'Else
            'bool_Return = false
            'End If

            Return obj_RetValue
        End Function
    End Class

    Public Class cls_myField
        Property NomeCampo As String
        Property Etichetta As String
        Property TipoCampo As System.Data.DbType
        Property TabellaRiferimento As String

        Public Sub New()
            NomeCampo = ""
            TipoCampo = Data.DbType.String
            TabellaRiferimento = ""
            Etichetta = ""
        End Sub

    End Class

#End Region '"FIELD class"


#Region "Pers_Update Class"

    Public Class cls_myUpdateFieldList
        Inherits cls_BaseList(Of cls_myUpdateField)

        Public Overrides Function Add(ByRef Valore As cls_myUpdateField) As cls_myUpdateField
            If Valore IsNot Nothing Then
                With Valore
                    With Valore
                        Return Add(.NomeCampo, .Espressione)
                    End With
                End With
            Else
                Return Nothing
            End If
        End Function

        Public Overloads Function Add(ByVal Nome As String, ByVal Espressione As String) As cls_myUpdateField
            Dim bool_Return As Boolean = False
            Dim str_ID As String = "2"
            Nome = Nome.ToUpper.Trim
            Dim obj_RetValue As New cls_myUpdateField

            Try
                'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                With obj_RetValue
                    .NomeCampo = Nome
                    .Espressione = Espressione

                End With

            Catch ex As Exception
                obj_RetValue = Nothing
            End Try

            str_ID = Nome.Trim.ToUpper

            'If InizializzaListaParametri() Then
            MyBase._ListaValori.Add(str_ID, obj_RetValue)
            'Else
            'bool_Return = false
            'End If

            Return obj_RetValue
        End Function
    End Class

    Public Class cls_myUpdateField
        Property NomeCampo As String
        Property Espressione As String


        Public Sub New()
            NomeCampo = ""
            Espressione = ""
        End Sub

    End Class

#End Region '"Pers_Update Class"

#Region "ORDER BY class"

    Public Class cls_myOrderByList
        Inherits cls_BaseList(Of cls_myOrderBy)

        Public Overrides Function Add(ByRef Valore As cls_myOrderBy) As cls_myOrderBy
            If Valore IsNot Nothing Then
                With Valore
                    With Valore
                        Return Add(.NomeCampo, .TabellaRiferimento, .Direzione)
                    End With
                End With
            Else
                Return Nothing
            End If
        End Function

        Public Overloads Function Add(ByVal Nome As String, Optional ByVal Tabella As String = "", Optional ByVal DirezioneOrdinamento As cls_myOrderBy.en_Direzione = cls_myOrderBy.en_Direzione.Crescente) As cls_myOrderBy
            Dim bool_Return As Boolean = False
            Nome = Nome.ToUpper.Trim
            Dim obj_RetValue As New cls_myOrderBy

            Try
                With obj_RetValue
                    .NomeCampo = Nome
                    .TabellaRiferimento = Tabella
                    .Direzione = DirezioneOrdinamento
                End With

            Catch ex As Exception
                obj_RetValue = Nothing
            End Try

            MyBase._ListaValori.Add(Nome, obj_RetValue)

            Return obj_RetValue
        End Function


    End Class

    Public Class cls_myOrderBy
        Property NomeCampo As String
        Property TabellaRiferimento As String
        Property Direzione As en_Direzione

        Public Enum en_Direzione
            Crescente = 1
            Decrescente = 2
        End Enum

        Public Sub New()
            NomeCampo = ""
            TabellaRiferimento = ""
            Direzione = en_Direzione.Crescente
        End Sub

    End Class

#End Region '"ORDER BY class"


#Region "GROUP BY class"

    Public Class cls_myGroupByList
        Inherits cls_BaseList(Of cls_myGroupBy)

        Public Overrides Function Add(ByRef Valore As cls_myGroupBy) As cls_myGroupBy
            If Valore IsNot Nothing Then
                With Valore
                    With Valore
                        Return Add(.NomeCampo, .TabellaRiferimento)
                    End With
                End With
            Else
                Return Nothing
            End If
        End Function

        Public Overloads Function Add(ByVal Nome As String, Optional ByVal Tabella As String = "") As cls_myGroupBy
            Dim bool_Return As Boolean = False
            Nome = Nome.ToUpper.Trim
            Dim obj_RetValue As New cls_myGroupBy

            Try
                'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                With obj_RetValue
                    .NomeCampo = Nome
                    .TabellaRiferimento = Tabella
                End With

            Catch ex As Exception
                obj_RetValue = Nothing
            End Try

            'If InizializzaListaParametri() Then
            MyBase._ListaValori.Add(Nome, obj_RetValue)
            'Else
            'bool_Return = false
            'End If

            Return obj_RetValue
        End Function

    End Class

    Public Class cls_myGroupBy
        Property NomeCampo As String
        Property TabellaRiferimento As String

        Public Sub New()
            NomeCampo = ""
            TabellaRiferimento = ""
        End Sub

    End Class

#End Region '"GROUP BY class"


#Region "HAVING class"

    Public Class cls_myHavingList
        Inherits cls_BaseList(Of cls_myHaving)


        Public Overrides Function Add(ByRef Valore As cls_myHaving) As cls_myHaving
            If Valore IsNot Nothing Then
                With Valore
                    '    Return Add(.Lingua, .RegolaComposizioneDescrizione, .DescrizioneBase)

                    With Valore
                        Return Add(.Condizione, .RifParametro, .Catena, .NomeCampo)
                    End With

                    'Return Nothing
                End With
            Else
                Return Nothing
            End If
        End Function

        Public Overloads Function Add(ByVal Condizione As String, Optional ByVal RifParametro As String = "", Optional ByVal CatenaCondizione As String = "STD",
                           Optional ByVal NomeCondizione As String = "") As cls_myHaving

            Dim bool_Return As Boolean = False
            Dim obj_RetValue As New cls_myHaving

            'Sistemo Parametri
            Condizione = Condizione.ToUpper.Trim
            RifParametro = RifParametro.ToUpper.Trim
            CatenaCondizione = CatenaCondizione.ToUpper.Trim
            NomeCondizione = NomeCondizione.ToUpper.Trim
            If NomeCondizione = "" Then
                NomeCondizione = Condizione
                'TODO
                'MsgBox("cls_myEasyQuery.Add (NomeCondizione): TODO!!!!")
            End If

            If _ListaValori.ContainsKey(NomeCondizione) Then
                Throw New System.Exception("Condizione già esistente")
                bool_Return = False
                obj_RetValue = Nothing
            Else
                'Creo Oggetto da inserire
                Try
                    'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                    With obj_RetValue
                        .NomeCampo = NomeCondizione
                        .Condizione = Condizione
                        .RifParametro = RifParametro
                        .Catena = CatenaCondizione
                    End With

                Catch ex As Exception
                    obj_RetValue = Nothing
                End Try

                'Aggiungo Oggetto alla lista
                'If InizializzaListaParametri() Then
                MyBase._ListaValori.Add(NomeCondizione, obj_RetValue)
                'Else
                'bool_Return = false
                'End If
            End If


            Return obj_RetValue
        End Function
    End Class

    Public Class cls_myHaving
        Public Property NomeCampo As String
        Public Property Condizione As String
        Public Property Catena As String
        Public Property RifParametro As String

        Public Sub New()
            NomeCampo = ""
            Condizione = ""
            Catena = ""
            RifParametro = ""
        End Sub

    End Class

#End Region '"HAVING class"


#Region "ADDITIONAL TABLES class"

#End Region
End Class
