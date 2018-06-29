Public Class cls_Utility

    Shared Function Lookup(ByRef DBConnection As MyDatabase2.cls_myDatabase, ByVal NomeTabella As String, ByVal NomeCampo As String, ByVal Condizione As String) As Object
        Dim obj_RetValue As Object = Nothing

        Dim str_Query As String = ""

        str_Query = String.Format("SELECT {0} FROM {1} WHERE {2}", NomeCampo, NomeTabella, Condizione)

        Using obj_Query As New MyDatabase2.cls_myQuery
            With obj_Query
                .MyDBClass = DBConnection
                .Query = str_Query
                Try
                    obj_RetValue = .eseguiQuery(cls_myQuery.en_TipoRisultatoQuery.Scalare)
                Catch ex As Exception
                    obj_RetValue = Nothing
                End Try
            End With
        End Using

        Return obj_RetValue
    End Function

    Shared Function IncrementaValore(ByRef DBConnection As MyDatabase2.cls_myDatabase, ByVal NomeTabella As String, ByVal NomeCampo As String, ByVal Condizione As String, Optional ByVal Incremento As Integer = 1) As Boolean
        Dim bool_RetValue As Boolean = False

        Dim str_Query As String = ""
        Dim int_NumRes As Integer = 0

        str_Query = String.Format("UPDATE {0} SET {1}={1}+{2} WHERE {3}", NomeTabella, NomeCampo, Incremento, Condizione)

        Using obj_Query As New MyDatabase2.cls_myQuery
            With obj_Query
                .MyDBClass = DBConnection
                .Query = str_Query
                Try
                    int_NumRes = .eseguiQuery(cls_myQuery.en_TipoRisultatoQuery.CommandQuery)
                Catch ex As Exception
                    int_NumRes = -1
                End Try
            End With
        End Using

        Return bool_RetValue
    End Function

End Class
