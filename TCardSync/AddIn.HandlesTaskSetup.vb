Imports System.Text
Imports EPDM.Interop.epdm

Partial Public Class AddIn

    Dim currentSetupPage As SetupPage
    ' this gets triggered when you choose the task from the dropdown in the new task dialog
    Private Sub HandlesTaskSetup(poCmd As EdmCmd)

        Dim props As IEdmTaskProperties
        props = poCmd.mpoExtra

        'Turn on some properties, e.g., the task can be launched during a
        'state change, can extend the details page, is called when the
        'task is launched, and supports scheduling
        props.TaskFlags = EdmTaskFlag.EdmTask_SupportsChangeState + EdmTaskFlag.EdmTask_SupportsDetails + EdmTaskFlag.EdmTask_SupportsInitExec + EdmTaskFlag.EdmTask_SupportsScheduling

        'Set menu commands that launch this task from File Explorer
        Dim cmds(0) As EdmTaskMenuCmd
        cmds(0).mbsMenuString = $"Tasks\\{props.TaskName}"
        cmds(0).mbsStatusBarHelp = $"Run {props.TaskName}"
        cmds(0).mlCmdID = 148
        cmds(0).mlEdmMenuFlags = EdmMenuFlags.EdmMenu_OnlyFiles
        props.SetMenuCmds(cmds)

        'Add a custom setup page; SetupPage is a user control with an
        'edit box; SetupPage::LoadData populates the edit box from a
        'variable in IEdmTaskProperties; saving of properties is handled
        'by OnTaskSetupButton 
        currentSetupPage = New SetupPage
        currentSetupPage.CreateControl()

        currentSetupPage.LoadData(poCmd)

        Dim pages(0) As EdmTaskSetupPage
        pages(0).mbsPageName = "Variables"
        pages(0).mlPageHwnd = currentSetupPage.Handle.ToInt32
        pages(0).mpoPageImpl = currentSetupPage

        props.SetSetupPages(pages)


    End Sub

End Class
