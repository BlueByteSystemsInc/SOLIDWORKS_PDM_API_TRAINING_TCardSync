Imports EPDM.Interop.epdm

Partial Public Class AddIn



    Private Sub HandlesTaskSetupButton(poCmd As EdmCmd)

        If poCmd.mbsComment = "OK" And Not currentSetupPage Is Nothing Then
            currentSetupPage.StoreData(poCmd)
        End If
        currentSetupPage = Nothing

    End Sub

End Class
