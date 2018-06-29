Public Class cls_myAdditionalTables
    Enum en_TipoJoinTable
        INNER = 1
        LEFT = 2
        RIGHT = 4
    End Enum

    Private _Joinconditions As cls_myJoinConditionList
    Private _Joinconditions_Manual As cls_myJoinConditionList

    Public Property Joins As cls_myJoinConditionList
        Get
            Return _Joinconditions
        End Get
        Set(value As cls_myJoinConditionList)
            _Joinconditions = value
        End Set
    End Property


    Public Property JoinsManual As cls_myJoinConditionList
        Get
            Return _Joinconditions_Manual
        End Get
        Set(value As cls_myJoinConditionList)
            _Joinconditions_Manual = value
        End Set
    End Property

    Public Property NomeTabella As String
    Public Property TipoJoin As en_TipoJoinTable
    'Public Property CondizioneJoin As String

    Public Sub New()
        NomeTabella = ""
        TipoJoin = en_TipoJoinTable.INNER
        'CondizioneJoin = ""
        _Joinconditions = New cls_myJoinConditionList
        _Joinconditions_Manual = New cls_myJoinConditionList
    End Sub

End Class



Public Class cls_myJoinConditionList
    Inherits cls_BaseList(Of cls_myJoinCondition)

    Public Overrides Function Add(ByRef Valore As cls_myJoinCondition) As cls_myJoinCondition
        Dim str_IDCondizione As String = ""
        Try
            str_IDCondizione = String.Format("{0}={1}", Valore.NomeCampo_TabellaBase, Valore.NomeCampo_TabellaAggiuntiva)
            str_IDCondizione = str_IDCondizione.Trim.ToUpper

            MyBase._ListaValori.Add(str_IDCondizione, Valore)
            Return Valore
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Overloads Function Add(ByVal NomeCampoPrincipale As String, ByVal NomeCampoDipendente As String) As cls_myJoinCondition
        Try
            Return Add(New cls_myJoinCondition(NomeCampoPrincipale, NomeCampoDipendente))
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


End Class

Public Class cls_myJoinCondition
    Property NomeCampo_TabellaBase As String
    Property NomeCampo_TabellaAggiuntiva As String


    Sub New()
        NomeCampo_TabellaBase = ""
        NomeCampo_TabellaAggiuntiva = ""
    End Sub

    Sub New(ByVal NomeCampo_Principale As String, ByVal NomeCampo_Secondaria As String)
        NomeCampo_TabellaBase = NomeCampo_Principale
        NomeCampo_TabellaAggiuntiva = NomeCampo_Secondaria
    End Sub
End Class