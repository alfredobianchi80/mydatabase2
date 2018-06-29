Public Class info


    Public Shared Function Versione() As String
        Dim str_VersioneCorrente As String = ""
        str_VersioneCorrente = "0.0.0.0"
        Try
            str_VersioneCorrente = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major
            str_VersioneCorrente = str_VersioneCorrente & "." & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor
            str_VersioneCorrente = str_VersioneCorrente & "." & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build

            str_VersioneCorrente = str_VersioneCorrente & "." & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Revision
        Catch ex As Exception
            str_VersioneCorrente = "0.9.9.9"
        End Try


        Return str_VersioneCorrente
    End Function

    ''' <summary>
    ''' Questa funzione fa questo e quest'altro
    ''' </summary>
    ''' <param name="VersioneCheck">String - Versione con cui verificare la compatibilità</param>
    ''' <returns>0: Se non Compatibile ; 1: Se compatibile con limitazioni ; 2: Se compatibile (cambia fix) ; 3: Se pienamente compatibile</returns>
    ''' <remarks>Formato versione Versione.Revisione.Fix</remarks>
    Public Shared Function isCompatibleWith(ByVal VersioneCheck As String) As Int32


        Dim int_Result As Int32
        Dim bool_FormatoOk As Boolean
        Dim str_MessaggioNoCompatibile As String
        Dim arr_VersioneCheck As List(Of String)
        Dim arr_VersioneCurrent As List(Of String)

        'Fase 1: Verifico valore parametro «Versione» passato
        'Formato atteso: X.X.X o X.X.X.X
        bool_FormatoOk = True


        'Fase 2: Splitta versione
        arr_VersioneCheck = New List(Of String)(VersioneCheck.Split("."))
        arr_VersioneCurrent = New List(Of String)(Versione.Split("."))

        int_Result = 3
        If Convert.ToInt32(arr_VersioneCurrent(0)) <> Convert.ToInt32(arr_VersioneCheck(0)) Then
            int_Result = 0
            str_MessaggioNoCompatibile = "Versione non compatibile"

        Else
            If (Convert.ToInt32(arr_VersioneCurrent(1)) < Convert.ToInt32(arr_VersioneCheck(1))) Then
                int_Result = 1
                str_MessaggioNoCompatibile = "Compatibilità parziale"
            Else
                If (Convert.ToInt32(arr_VersioneCurrent(2)) < Convert.ToInt32(arr_VersioneCheck(2))) Then
                    int_Result = 2
                    str_MessaggioNoCompatibile = "Fix Correttive"
                End If
            End If
        End If

        Return int_Result
    End Function

End Class
