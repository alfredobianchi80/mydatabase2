Public Class cls_myAdditionalTablesList
    Inherits cls_BaseList(Of cls_myAdditionalTables)

    Public Sub New()
        MyBase.New()
    End Sub

    Public Overrides Function Add(ByRef Valore As cls_myAdditionalTables) As cls_myAdditionalTables
        Try
            MyBase._ListaValori.Add(Valore.NomeTabella.ToUpper.Trim, Valore)
            Return Valore
        Catch ex As Exception
            Return Nothing
        End Try
    End Function



    Public Overloads Function Add(ByVal NomeTabella As String, ByVal TipoJoin As cls_myAdditionalTables.en_TipoJoinTable) As cls_myAdditionalTables

        Dim bool_Return As Boolean = False
        Dim obj_RetValue As New cls_myAdditionalTables

        'Sistemo Parametri
        NomeTabella = NomeTabella.ToUpper.Trim

        If MyBase._ListaValori.ContainsKey(NomeTabella) Then
            Throw New System.Exception("Tabella già esistente")
            bool_Return = False
        Else
            'Creo Oggetto da inserire
            Try
                'obj_RetValue = _DBConnection.CreateCommand.CreateParameter()
                With obj_RetValue
                    .NomeTabella = NomeTabella
                    .TipoJoin = TipoJoin
                End With

            Catch ex As Exception
                obj_RetValue = Nothing
            End Try

            'Aggiungo Oggetto alla lista
            If obj_RetValue IsNot Nothing Then
                Try
                    MyBase._ListaValori.Add(NomeTabella, obj_RetValue)
                Catch ex As Exception
                    obj_RetValue = Nothing
                End Try
            End If
        End If

            Return obj_RetValue
    End Function

   

   
End Class


