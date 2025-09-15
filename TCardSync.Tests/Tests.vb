Imports EPDM.Interop.epdm

Module Tests

    Public Sub Main()

        Dim TCardSync As New AddIn

        Dim Vault As New EdmVault5

        Vault.LoginAuto("Assemblageddon", 0)
        Dim Folder As IEdmFolder5

        Dim File As IEdmFile5 = Vault.GetFileFromPath("C:\Assemblageddon\Speaker\Speaker.slddrw", Folder)

        Dim poCmd As New EdmCmd

        poCmd.meCmdType = EdmCmdType.EdmCmd_TaskRun

        poCmd.mpoVault = Vault
        poCmd.mpoExtra = New EdmTaskInstance

        Dim poData As EdmCmdData()

        poData = New EdmCmdData(0) {
            New EdmCmdData With {
                .mlObjectID1 = File.ID,
                .mlObjectID2 = Folder.ID
            }
        }



        TCardSync.OnCmd(poCmd, poData)

    End Sub

End Module


Public Class EdmTaskInstance
    Implements IEdmTaskInstance

    Public Sub SetProgressRange(lMax As Integer, lPos As Integer, bsDocStr As String) Implements IEdmTaskInstance.SetProgressRange
        Console.WriteLine($"Doc Count: {lMax} Message: {bsDocStr    }")
    End Sub

    Public Sub SetProgressPos(lPos As Integer, bsDocStr As String) Implements IEdmTaskInstance.SetProgressPos
        Console.WriteLine($"Index: {lPos} Message: {bsDocStr    }")

    End Sub

    Public Function GetStatus() As EdmTaskStatus Implements IEdmTaskInstance.GetStatus
        Return EdmTaskStatus.EdmTaskStat_Running
    End Function

    Public Sub SetStatus(eStatus As EdmTaskStatus, Optional lHRESULT As Integer = Nothing, Optional bsCustomMsg As String = Nothing, Optional oNotificationAttachments As Object = Nothing, Optional bsExtraNotificationMsg As String = Nothing) Implements IEdmTaskInstance.SetStatus
        Console.WriteLine($"Status: {eStatus} Message: {bsCustomMsg    }")
    End Sub

    Public Function GetVar(oVarIDorName As Object) As Object Implements IEdmTaskInstance.GetVar
        Throw New NotImplementedException()
    End Function

    Public Sub SetVar(oVarIDorName As Object, oValue As Object) Implements IEdmTaskInstance.SetVar
        Throw New NotImplementedException()
    End Sub

    Public Function GetValEx(bsValName As String) As Object Implements IEdmTaskInstance.GetValEx
        Return "Description"
    End Function

    Public Sub SetValEx(bsValName As String, oValue As Object) Implements IEdmTaskInstance.SetValEx
        Throw New NotImplementedException()
    End Sub

    Public ReadOnly Property ID As Long Implements IEdmTaskInstance.ID
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property InstanceGUID As String Implements IEdmTaskInstance.InstanceGUID
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property TaskGUID As String Implements IEdmTaskInstance.TaskGUID
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property TaskName As String Implements IEdmTaskInstance.TaskName
        Get
            Throw New NotImplementedException()
        End Get
    End Property
End Class