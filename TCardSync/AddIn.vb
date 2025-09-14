Imports System.Runtime.InteropServices
Imports System.Text
Imports EPDM.Interop.epdm




<Guid("587DA78F-5973-4A3B-AFFE-567868CE8CBD")>
<ComVisible(True)>
Partial Public Class AddIn
    Implements IEdmAddIn5




    Public Const ADDIN_NAME As String = "TCardSync"
    Public Const STORAGEKEY As String = "VariableNames"
    Dim errorLogs As New StringBuilder()

    ' This method is called by the vault to get information about the add-in.
    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

        ' Set the add-in information.

        poInfo.mbsAddInName = ADDIN_NAME
        poInfo.mbsCompany = "BLUE BYTE SYSTEMS INC."
        poInfo.mbsDescription = "Task data card values between drawing and models"
        'suggest format yyyy-mm-dd for versioning 
        poInfo.mlAddInVersion = 1

        ' 18 corresponding to SolidWorks 2018
        ' more information here: https://github.com/BlueByteSystemsInc/SOLIDWORKS-PDM-API-SDK/blob/master/src/BlueByte.SOLIDWORKS.PDMProfessional.SDK/Enums/PDMProfessionalVersion_e.cs#L12
        poInfo.mlRequiredVersionMajor = 18
        ' this is the service pack number, 0 to 5
        poInfo.mlRequiredVersionMinor = 0

        'min task hooks
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup)
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton)

        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun)
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskLaunch)




    End Sub

    ' This method is called by the vault when a command associated with the add-in is executed.
    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd

        Dim vault As IEdmVault5 = poCmd.mpoVault
        Dim userMgr As IEdmUserMgr5 = vault
        Dim handle As Integer = poCmd.mlParentWnd
        Dim loggedInUser As IEdmUser5 = userMgr.GetLoggedInUser()


        Try

            Select Case poCmd.meCmdType
                Case EdmCmdType.EdmCmd_TaskLaunch

                    HandlesTaskLaunch(poCmd, ppoData)
                Case EdmCmdType.EdmCmd_TaskSetup

                    HandlesTaskSetup(poCmd)
                Case EdmCmdType.EdmCmd_TaskSetupButton

                    HandlesTaskSetupButton(poCmd)
                Case EdmCmdType.EdmCmd_TaskRun

                    HandlesTaskRun(poCmd, ppoData)
            End Select


        Catch ex As Exception

            ' todo report task cancellation or failure 

        End Try
    End Sub

End Class



