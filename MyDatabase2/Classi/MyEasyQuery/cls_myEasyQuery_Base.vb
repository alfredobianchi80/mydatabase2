Public MustInherit Class cls_myEasyQuery_Base
    Inherits cls_myQuery

#Region "Enumerati e Strutture"

    Protected Structure tp_WhereField
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
    Protected _ListaCampi As cls_MySubClass.cls_myFieldsList
    Protected _ListaWhere_V1 As cls_MySubClass.cls_myWhereList
    Protected _ListaWhere_V2 As cls_MySubClass.cls_myBaseCondictionWhereLIST
    Protected _ListaOrderBy As cls_MySubClass.cls_myOrderByList
    Protected _ListaGroupBY As cls_MySubClass.cls_myGroupByList
    Protected _ListaHaving As cls_MySubClass.cls_myHavingList
    Protected _ListaTabelleAddizionali As cls_myAdditionalTablesList

    Protected _ListaCampiPersUpdarte As cls_MySubClass.cls_myUpdateFieldList

    Protected _RiordinaParametri As Boolean
    Protected _TipoQuerySelezionata As en_TipoQuery
    Protected _ListaSequenzaParametri As List(Of String)

    Protected _VersioneWhere As Int32 = 2

    Protected _OperatoreFraCatene As String = "OR"
    Protected _OperatoreINCatene As String = "AND"


    Property OperatoreFraCatene As String
        Get
            Return _OperatoreFraCatene
        End Get
        Set(value As String)
            _OperatoreFraCatene = value
        End Set
    End Property

    Property OperatoreINCatene As String
        Get
            Return _OperatoreINCatene
        End Get
        Set(value As String)
            _OperatoreINCatene = value
        End Set
    End Property


#End Region '"Attributi"

#Region "Property"

    Protected Property _TipoQuery As en_TipoQuery
        Get
            Return _TipoQuerySelezionata
        End Get
        Set(value As en_TipoQuery)
            _TipoQuerySelezionata = value
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

    Protected Property _Field As cls_MySubClass.cls_myFieldsList
        Get
            If _ListaCampi Is Nothing Then
                _ListaCampi = New cls_MySubClass.cls_myFieldsList
            End If
            Return _ListaCampi
        End Get
        Set(value As cls_MySubClass.cls_myFieldsList)
            _ListaCampi = value
        End Set
    End Property

    Protected Property _OtherUpdateField As cls_MySubClass.cls_myUpdateFieldList
        Get
            If _ListaCampiPersUpdarte Is Nothing Then
                _ListaCampiPersUpdarte = New cls_MySubClass.cls_myUpdateFieldList
            End If
            Return _ListaCampiPersUpdarte
        End Get
        Set(value As cls_MySubClass.cls_myUpdateFieldList)
            _ListaCampiPersUpdarte = value
        End Set
    End Property

    Protected Property _Where_V2 As cls_MySubClass.cls_myBaseCondictionWhereLIST
        Get
            If _ListaWhere_V2 Is Nothing Then
                _ListaWhere_V2 = New cls_MySubClass.cls_myBaseCondictionWhereLIST
            End If

            Return _ListaWhere_V2
        End Get
        Set(value As cls_MySubClass.cls_myBaseCondictionWhereLIST)

        End Set
    End Property

    Protected Property _Where_V1 As cls_MySubClass.cls_myWhereList
        Get
            If _ListaWhere_V1 Is Nothing Then
                _ListaWhere_V1 = New cls_MySubClass.cls_myWhereList
            End If

            Return _ListaWhere_V1
        End Get
        Set(value As cls_MySubClass.cls_myWhereList)

        End Set
    End Property

    Protected Property _OrderBy As cls_MySubClass.cls_myOrderByList
        Get
            If _ListaOrderBy Is Nothing Then
                _ListaOrderBy = New cls_MySubClass.cls_myOrderByList
            End If

            Return _ListaOrderBy
        End Get
        Set(value As cls_MySubClass.cls_myOrderByList)

        End Set
    End Property

    Protected Property _Having As cls_MySubClass.cls_myHavingList
        Get
            If _ListaHaving Is Nothing Then
                _ListaHaving = New cls_MySubClass.cls_myHavingList
            End If

            Return _ListaHaving
        End Get
        Set(value As cls_MySubClass.cls_myHavingList)

        End Set
    End Property

    Protected Property _GroupBy As cls_MySubClass.cls_myGroupByList
        Get
            If _ListaGroupBY Is Nothing Then
                _ListaGroupBY = New cls_MySubClass.cls_myGroupByList
            End If

            Return _ListaGroupBY
        End Get
        Set(value As cls_MySubClass.cls_myGroupByList)

        End Set
    End Property

    Protected Property _TabelleAddizionali As cls_myAdditionalTablesList
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
        _ListaCampi = New cls_MySubClass.cls_myFieldsList
        _ListaWhere_V1 = New cls_MySubClass.cls_myWhereList
        _ListaWhere_V2 = New cls_MySubClass.cls_myBaseCondictionWhereLIST

        '_ListaCateneWhere = New cls_MySubClass.cls_myCatenaWhereList

        _ListaGroupBY = New cls_MySubClass.cls_myGroupByList
        _ListaHaving = New cls_MySubClass.cls_myHavingList
        _ListaOrderBy = New cls_MySubClass.cls_myOrderByList
        _ListaTabelleAddizionali = New cls_myAdditionalTablesList
        _ListaCampiPersUpdarte = New cls_MySubClass.cls_myUpdateFieldList

        _NomeTabella = ""
        _ListaSequenzaParametri = New List(Of String)
        _RiordinaParametri = True
    End Sub

#End Region '"Costruttore"

#Region "Overrides Elements"


    Public Property Query_ID() As String
        Get
            Return ""
            'Dim str_Query As String
            '_ListaSequenzaParametri.Clear()

            'Select Case _TipoQuerySelezionata
            '    Case en_TipoQuery.myINSERT
            '        str_Query = CreaQuery_INSERT(True)


            '    Case Else
            '        Throw New System.Exception("Casistica non gestita!!!!")
            '        str_Query = ""
            'End Select
        End Get
        Set(value As String)

        End Set
    End Property


    Public Overrides Property Query() As String
        Get
            Dim str_Query As String = ""

            _ListaSequenzaParametri.Clear()

            Select Case _TipoQuerySelezionata
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

    Public Function eseguiQuery_ID(Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.DataTable) As Object
        Dim obj_RetValue As Object = Nothing
        Dim str_TestoQuery As String = Query_ID
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


    Protected Function _Get_WhereString_V1() As String
        Dim str_Result As String = ""

        str_Result = ""

        For Each obj_Where As cls_MySubClass.cls_myWhere In _ListaWhere_V1.Items

            If str_Result.Length > 0 Then
                str_Result = str_Result & " AND "
            End If
            str_Result = str_Result & " (" & obj_Where.Condizione & ")"

            If obj_Where.RifParametro.ToUpper.Trim.Length > 0 Then
                If Not (_ListaSequenzaParametri.Contains(obj_Where.RifParametro.ToUpper.Trim)) Then
                    _ListaSequenzaParametri.Add(obj_Where.RifParametro.ToUpper.Trim)
                End If
            End If
        Next

        Return str_Result
    End Function

    Protected Function _Get_WhereString_V2() As String
        Dim str_Result As String = ""
        Dim str_Result_Catena As String = ""

        'Throw New System.Exception("[_Get_WhereString_V2] Non ancora implementato.")

        str_Result = ""
        str_Result_Catena = ""

        'Cicla catene
        For Each obj_Catena As cls_MySubClass.cls_myBaseCondictionWhere In _ListaWhere_V2.Items
            str_Result_Catena = ""
            'Cicla Condizioni
            For Each obj_Where As cls_MySubClass.cls_myBaseWhere In obj_Catena.ListaCondizioni.ListOf
                If str_Result_Catena.Length > 0 Then
                    str_Result_Catena = str_Result_Catena & " " & _OperatoreINCatene.Trim.ToUpper & " "
                End If
                str_Result_Catena = str_Result_Catena & " (" & obj_Where.Condizione & ")"

                If obj_Where.RifParametro.ToUpper.Trim.Length > 0 Then
                    If Not (_ListaSequenzaParametri.Contains(obj_Where.RifParametro.ToUpper.Trim)) Then
                        _ListaSequenzaParametri.Add(obj_Where.RifParametro.ToUpper.Trim)
                    End If
                End If
            Next

            If str_Result.Length > 0 Then
                str_Result = str_Result & " " & _OperatoreFraCatene.Trim.ToUpper & " "
            End If
            str_Result = str_Result & "(" & str_Result_Catena & ")"
        Next

        Dim str_Filtri_Comuni As String = ""
        str_Filtri_Comuni = _Get_WhereString_V1()

        If str_Filtri_Comuni.Trim.Length > 0 Then
            If str_Result.Trim.Length > 0 Then
                str_Result = "(" & str_Result & ") AND "
            End If
            str_Result = str_Result & "(" & str_Filtri_Comuni & ")"
        End If

        Return str_Result
    End Function

    Protected Function _Get_GroupByString() As String
        Dim str_Result As String = ""

        str_Result = ""
        For Each obj_GroupBy As cls_MySubClass.cls_myGroupBy In _ListaGroupBY.Items

            If obj_GroupBy.NomeCampo.Length > 0 Then
                If str_Result.Length > 0 Then
                    str_Result = str_Result & ", "
                End If

                If obj_GroupBy.TabellaRiferimento.Length > 0 Then
                    str_Result = str_Result & obj_GroupBy.TabellaRiferimento.Trim & "."
                End If

                str_Result = str_Result & obj_GroupBy.NomeCampo
            End If
        Next

        Return str_Result
    End Function

    Protected Function _Get_HavingString() As String
        Dim str_Result As String = ""

        str_Result = ""
        For Each obj_having As cls_MySubClass.cls_myHaving In _ListaHaving.Items
            If str_Result.Length > 0 Then
                str_Result = str_Result & " AND "
            End If
            str_Result = str_Result & " (" & obj_having.Condizione & ")"

            If obj_having.RifParametro.ToUpper.Trim.Length > 0 Then
                If Not (_ListaSequenzaParametri.Contains(obj_having.RifParametro.ToUpper.Trim)) Then
                    _ListaSequenzaParametri.Add(obj_having.RifParametro.ToUpper.Trim)
                End If
            End If
        Next

        Return str_Result
    End Function

    Protected Function _Get_OrderByString() As String
        Dim str_Result As String = ""

        str_Result = ""
        For Each obj_OrderBy As cls_MySubClass.cls_myOrderBy In _ListaOrderBy.Items
            If str_Result.Length > 0 Then
                str_Result = str_Result & " , "
            End If

            If obj_OrderBy.TabellaRiferimento.Length > 0 Then
                str_Result = str_Result & obj_OrderBy.TabellaRiferimento.Trim & "."
            End If
            str_Result = str_Result & obj_OrderBy.NomeCampo

            If obj_OrderBy.Direzione = cls_MySubClass.cls_myOrderBy.en_Direzione.Decrescente Then
                str_Result = str_Result & " DESC"
            End If
        Next

        Return str_Result
    End Function

    Protected Function _Get_JOINString() As String
        Dim str_Result As String = ""

        str_Result = _NomeTabella

        If _ListaTabelleAddizionali.Count > 0 Then

            str_Result = ""
            For i As Int32 = 0 To _ListaTabelleAddizionali.Count - 1
                str_Result = str_Result & "("
            Next
            str_Result = str_Result & _NomeTabella

            For Each obj_Join As cls_myAdditionalTables In _ListaTabelleAddizionali.ListOf
                Select Case obj_Join.TipoJoin
                    Case cls_myAdditionalTables.en_TipoJoinTable.INNER
                        str_Result = str_Result & " INNER JOIN "
                    Case cls_myAdditionalTables.en_TipoJoinTable.LEFT
                        str_Result = str_Result & " LEFT JOIN "
                    Case cls_myAdditionalTables.en_TipoJoinTable.RIGHT
                        str_Result = str_Result & " RIGHT JOIN "
                End Select

                str_Result = str_Result & obj_Join.NomeTabella & " ON("
                Dim str_JoinFields As String = ""

                Dim int_temp As Int32 = 0
                For Each obj_cc As cls_myJoinCondition In obj_Join.Joins.ListOf
                    'str_Join = str_Join & String.Format("{%0}.{%1}={%2}.{%3}", Tabella, obj_cc.NomeCampo_TabellaBase, obj_Join.NomeTabella, obj_cc.NomeCampo_TabellaAggiuntiva)
                    If str_JoinFields.Trim.Length > 0 Then
                        str_JoinFields = str_JoinFields & " AND "
                    End If
                    str_JoinFields = str_JoinFields & String.Concat(Tabella, ".", obj_cc.NomeCampo_TabellaBase, "=", obj_Join.NomeTabella, ".", obj_cc.NomeCampo_TabellaAggiuntiva)
                    int_temp += 1
                Next


                For Each obj_cc As cls_myJoinCondition In obj_Join.JoinsManual.ListOf
                    'str_Join = str_Join & String.Format("{%0}.{%1}={%2}.{%3}", Tabella, obj_cc.NomeCampo_TabellaBase, obj_Join.NomeTabella, obj_cc.NomeCampo_TabellaAggiuntiva)
                    If str_JoinFields.Trim.Length > 0 Then
                        str_JoinFields = str_JoinFields & " AND "
                    End If
                    str_JoinFields = str_JoinFields & String.Concat(obj_cc.NomeCampo_TabellaBase, "=", obj_cc.NomeCampo_TabellaAggiuntiva)
                    int_temp += 1
                Next
                str_Result = str_Result & str_JoinFields
                str_Result = str_Result & ")"

                'If _ListaTabelleAddizionali.Count > 1 Then
                str_Result = str_Result & ")"
                'End If
            Next

        End If




        Return str_Result
    End Function


    Protected Function CreaQuery_SELECT() As String
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
        For Each o As cls_MySubClass.cls_myField In _ListaCampi.Items
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
        Select Case _VersioneWhere
            Case 1
                str_Where = _Get_WhereString_V1()
            Case 2
                str_Where = _Get_WhereString_V2()
        End Select

        '** GROUP BY
        str_GroupBy = _Get_GroupByString()

        '** HAVING
        str_Having = _Get_HavingString()

        '** ORDER BY
        str_OrderBy = _Get_OrderByString()

        '** JOIN
        str_Join = _Get_JOINString()


        '* Componi Query
        str_Query = "SELECT " & str_Fields & " FROM "

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

    Protected Function CreaQuery_INSERT(Optional ByVal RestiuisciID As Boolean = False) As String
        Dim str_Query As String = ""
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


        If RestiuisciID Then
            str_Query = "SET NOCOUNT ON;" & str_Query & "; SELECT SCOPE_IDENTITY() as NewID;"
        End If

        Return str_Query
    End Function

    Protected Function CreaQuery_UPDATE() As String
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
        For Each obj_x As cls_MySubClass.cls_myUpdateField In _ListaCampiPersUpdarte.Items
            If str_ListaCampi <> "" Then
                str_ListaCampi = str_ListaCampi & ","
            End If
            str_ListaCampi = str_ListaCampi & obj_x.Espressione
        Next

        'Crea Condizione
        Select Case _VersioneWhere
            Case 1
                str_Where = _Get_WhereString_V1()
            Case 2
                str_Where = _Get_WhereString_V2()
        End Select
       
        If str_Where = "" Then
            Throw New System.Exception("Query di Update senza nessuna condizione indicata.")
        End If

        str_Query = "UPDATE " & _NomeTabella & " SET " & str_ListaCampi
        str_Query = str_Query & " WHERE " & str_Where


        Return str_Query
    End Function

    Protected Function CreaQuery_DELETE() As String
        Dim str_Query As String = ""
        Dim str_Where As String = ""

        _ListaSequenzaParametri.Clear()

        Select Case _VersioneWhere
            Case 1
                str_Where = _Get_WhereString_V1()
            Case 2
                str_Where = _Get_WhereString_V2()
        End Select
        If str_Where = "" Then
            Throw New System.Exception("Query di Update senza nessuna condizione indicata.")
        End If


        str_Query = "DELETE " & Get_DelFildPar() & " From " & _NomeTabella
        str_Query = str_Query & " WHERE " & str_Where

        Return str_Query
    End Function

#End Region '"Crea Query Function"


#Region "funzioni Protette"

    Protected Function Get_DelFildPar() As String
        Dim str_Result As String = ""


        If _obj_myDatabaseClass IsNot Nothing Then
            Select Case MyBase._obj_myDatabaseClass.TipoDatabase
                Case cls_myDatabase.en_TipoDB.ACCESS
                    str_Result = "*"
                Case cls_myDatabase.en_TipoDB.MSSQL
                    str_Result = ""
                Case cls_myDatabase.en_TipoDB.DB2
                    str_Result = ""
                Case cls_myDatabase.en_TipoDB.DBF
                    str_Result = ""
                Case cls_myDatabase.en_TipoDB.MYSQL
                    str_Result = ""
                Case Else

            End Select
        Else

            If TypeOf MyBase.MyDBClass.ConnessioneDB Is System.Data.SqlClient.SqlConnection Then
                str_Result = ""
            End If

            If TypeOf MyBase.MyDBClass.ConnessioneDB Is System.Data.OleDb.OleDbConnection Then
                str_Result = "*"
            End If
        End If


        Return str_Result
    End Function

#End Region



#Region "Sub Class"

    Public Class cls_MySubClass

#Region "WHERE class"


        'Public Class cls_myWhere
        '    Private _ListaCondizioni As Dictionary(Of String, cls_myBaseWhere)


        '    Sub New()
        '        _ListaCondizioni = New Dictionary(Of String, cls_myBaseWhere)
        '    End Sub


        'End Class

        'Public Class cls_myBaseWhere
        '    Public Property NomeCampo As String = ""
        '    Public Property Condizione As String = ""
        '    Public Property RifParametro As String = ""

        '    Sub New(ByVal NomeCampo As String, ByVal Condizione As String, ByVal RifParametro As String)
        '        Me.NomeCampo = NomeCampo
        '        Me.Condizione = Condizione
        '        Me.RifParametro = RifParametro
        '    End Sub
        'End Class


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



        '    Protected _ListaWhere As cls_MySubClass.cls_myWhereList


        Public Class cls_myBaseCondictionWhereLIST
            Inherits cls_BaseList(Of cls_myBaseCondictionWhere)

            Public Overrides Function Add(ByRef Valore As cls_myBaseCondictionWhere) As cls_myBaseCondictionWhere
                Try
                    MyBase._ListaValori.Add(Valore.NomeCatena, Valore)
                    Return MyBase._ListaValori.Item(Valore.NomeCatena)
                Catch ex As Exception
                    Return Nothing
                End Try

            End Function



            Public Overloads Function Add(ByVal Condizione As String, Optional ByVal RifParametro As String = "", Optional ByVal CatenaCondizione As String = "STD",
                                Optional ByVal NomeCondizione As String = "") As cls_myBaseCondictionWhere

                Dim bool_Return As Boolean = False

                'Sistemo Parametri
                CatenaCondizione = CatenaCondizione.ToUpper.Trim

                'Fase 1: Ricavo Catena Condizioni su cui agire
                Dim obj_OggettoCatena As cls_myBaseCondictionWhere = Nothing
                If MyBase._ListaValori.ContainsKey(CatenaCondizione) Then
                    obj_OggettoCatena = MyBase._ListaValori.Item(CatenaCondizione)
                Else
                    obj_OggettoCatena = New cls_myBaseCondictionWhere(CatenaCondizione)
                End If


                'Fase 2: Verifico se, all'interno della catena già c'è la condizione
                Dim obj_RetValue As New cls_myBaseWhere

                If obj_OggettoCatena.ListaCondizioni.ContainsKey(NomeCondizione) Then
                    Throw New System.Exception("Condizione già esistente")
                    obj_RetValue = Nothing
                Else
                    'Creo Oggetto da inserire
                    Try
                        'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                        With obj_RetValue
                            .NomeCampo = NomeCondizione
                            .Condizione = Condizione
                            .RifParametro = RifParametro
                        End With

                    Catch ex As Exception
                        obj_RetValue = Nothing
                    End Try

                    'Aggiungo Oggetto alla lista

                    obj_OggettoCatena.ListaCondizioni.Add(obj_RetValue)
                    
                End If

                MyBase._ListaValori.Item(CatenaCondizione) = obj_OggettoCatena

                '----------------

                'Condizione = Condizione.ToUpper.Trim
                'RifParametro = RifParametro.ToUpper.Trim

                'NomeCondizione = NomeCondizione.ToUpper.Trim
                'If NomeCondizione = "" Then
                '    NomeCondizione = Condizione
                '    'TODO
                '    'MsgBox("cls_myEasyQuery.Add (NomeCondizione): TODO!!!!")
                'End If




                Return obj_OggettoCatena
            End Function
        End Class

        Public Class cls_myBaseCondictionWhere
            Property NomeCatena As String = ""
            Property ListaCondizioni As cls_myBaseWhereList = Nothing

            Sub New()
                ListaCondizioni = New cls_myBaseWhereList
            End Sub

            Sub New(ByVal NomeCatena As String)
                Me.NomeCatena = NomeCatena.ToUpper.Trim
                ListaCondizioni = New cls_myBaseWhereList
            End Sub
        End Class

        Public Class cls_myBaseWhereList
            Inherits cls_BaseList(Of cls_myBaseWhere)

            Public Overrides Function Add(ByRef Valore As cls_myBaseWhere) As cls_myBaseWhere
                If Valore IsNot Nothing Then
                    Return Add(Valore.Condizione, Valore.RifParametro, Valore.NomeCampo)
                Else
                    Return Nothing
                End If
            End Function

            Public Overloads Function Add(ByVal Condizione As String, Optional ByVal RifParametro As String = "", Optional ByVal NomeCondizione As String = "") As cls_myBaseWhere
                Dim obj_RetValue As New cls_myBaseWhere
                'Sistemo Parametri
                Condizione = Condizione.ToUpper.Trim
                RifParametro = RifParametro.ToUpper.Trim

                NomeCondizione = NomeCondizione.ToUpper.Trim
                If NomeCondizione = "" Then
                    NomeCondizione = Condizione
                End If

                If MyBase._ListaValori.ContainsKey(NomeCondizione) Then
                    Throw New System.Exception("Condizione già esistente")
                    obj_RetValue = Nothing
                Else
                    'Creo Oggetto da inserire
                    Try
                        'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                        With obj_RetValue
                            .NomeCampo = NomeCondizione
                            .Condizione = Condizione
                            .RifParametro = RifParametro
                        End With

                    Catch ex As Exception
                        obj_RetValue = Nothing
                    End Try

                    'Aggiungo Oggetto alla lista
                    MyBase._ListaValori.Add(NomeCondizione, obj_RetValue)
                End If
                Return obj_RetValue
            End Function
        End Class

        Public Class cls_myBaseWhere
            Property NomeCampo As String
            Property Condizione As String
            Property RifParametro As String
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

    End Class

#End Region '"Sub Class"




End Class
