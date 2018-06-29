Public Class cls_myEasyQuery_Insert
    Inherits cls_myEasyQuery_Base


#Region "Property"


#End Region '"Property"


#Region "costruttore"

    Public Sub New()
        MyBase.New()
        MyBase._TipoQuery = en_TipoQuery.myINSERT
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

    Public Overloads Function eseguiQuery(Optional TipoRisultato As en_TipoRisultatoQuery = en_TipoRisultatoQuery.CommandQuery, Optional ByVal RestituisciID As Boolean = False) As Object
        Return MyBase.eseguiQuery(TipoRisultato)

    End Function

#End Region '"Metodi"

End Class
