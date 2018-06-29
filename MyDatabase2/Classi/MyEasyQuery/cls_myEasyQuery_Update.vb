Public Class cls_myEasyQuery_Update
    Inherits cls_myEasyQuery_Base


#Region "Property"


#Region "campi"

    Public Property Field As cls_MySubClass.cls_myFieldsList
        Get
            Return MyBase._Field
        End Get
        Set(value As cls_MySubClass.cls_myFieldsList)
            MyBase._Field = value
        End Set
    End Property

    Public Property ManualField As cls_MySubClass.cls_myUpdateFieldList
        Get
            Return MyBase._OtherUpdateField
        End Get
        Set(value As cls_MySubClass.cls_myUpdateFieldList)
            MyBase._OtherUpdateField = value
        End Set
    End Property

    Public Property Where As cls_MySubClass.cls_myWhereList
        Get
            Return MyBase._Where_V1
        End Get
        Set(value As cls_MySubClass.cls_myWhereList)
            MyBase._Where_V1 = value
        End Set
    End Property

    Public Property Where_V2 As cls_MySubClass.cls_myBaseCondictionWhereLIST
        Get
            Return MyBase._Where_V2
        End Get
        Set(value As cls_MySubClass.cls_myBaseCondictionWhereLIST)
            MyBase._Where_V2 = value
        End Set
    End Property


    Public Property TabelleAddizionali As cls_myAdditionalTablesList
        Get
            Return MyBase._TabelleAddizionali
        End Get
        Set(value As cls_myAdditionalTablesList)
            MyBase._TabelleAddizionali = value
        End Set
    End Property

    'Public Property OrderBy As cls_MySubClass.cls_myOrderByList
    '    Get
    '        Return MyBase._OrderBy
    '    End Get
    '    Set(value As cls_MySubClass.cls_myOrderByList)
    '        MyBase._OrderBy = value
    '    End Set
    'End Property

    'Public Property Having As cls_MySubClass.cls_myHavingList
    '    Get
    '        Return MyBase._Havinhg
    '    End Get
    '    Set(value As cls_MySubClass.cls_myHavingList)
    '        MyBase._Havinhg = value
    '    End Set
    'End Property

    'Public Property GroupBy As cls_MySubClass.cls_myGroupByList
    '    Get
    '        Return MyBase._GroupBy
    '    End Get
    '    Set(value As cls_MySubClass.cls_myGroupByList)
    '        MyBase._GroupBy = value
    '    End Set
    'End Property

#End Region '"campi"

#End Region '"Property"

#Region "costruttore"

    Public Sub New()
        MyBase.New()
        MyBase._TipoQuery = en_TipoQuery.myUPDATE
    End Sub

#End Region '"costruttore"

#Region "Metodi"

    Public Overrides Property Query() As String
        Get
            Return MyBase.Query
        End Get
        Set(value As String)

        End Set
    End Property

    Public Overloads Function eseguiQuery(Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.CommandQuery) As Object
        Return MyBase.eseguiQuery(TipoRisultato)
    End Function

#End Region '"Metodi"


End Class
