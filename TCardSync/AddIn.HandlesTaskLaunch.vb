Imports System.Text
Imports EPDM.Interop.epdm

Partial Public Class AddIn

    Private Sub HandlesTaskLaunch(poCmd As EdmCmd, data() As EdmCmdData)

        ' this is where you can to get information about the affected document
        ' and decide


        ' you have two abilities:
        ' cancel the task
        ' save data to the task 


        If (False) Then

            Dim instance As IEdmTaskInstance = poCmd.mpoExtra
            'cancel task
            instance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneCancelled, 0, "Reason...")


            Dim k As String = String.Empty
            Dim v As Object = Nothing

            'store data
            instance.SetValEx(k, v)


            'warning: You cannot store objects!
        End If

    End Sub

End Class
