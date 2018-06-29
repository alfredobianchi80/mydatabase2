Public Class cls_myEasyQuery_Delete
    Inherits cls_myEasyQuery_Base


#Region "Property"


#Region "campi"

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

#End Region '"campi"

#End Region '"Property"

#Region "costruttore"

    Public Sub New()
        MyBase.New()
        MyBase._TipoQuery = en_TipoQuery.myDELETE

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
