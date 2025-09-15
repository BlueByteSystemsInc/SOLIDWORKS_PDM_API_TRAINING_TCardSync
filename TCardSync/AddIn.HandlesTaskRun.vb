Imports System.Reflection.Emit
Imports EPDM.Interop.epdm

Partial Public Class AddIn



    Private Sub HandlesTaskRun(poCmd As EdmCmd, ppoData() As EdmCmdData)

        Dim handle = poCmd.mlParentWnd
        Dim vault As IEdmVault5 = poCmd.mpoVault
        Dim instance As IEdmTaskInstance = poCmd.mpoExtra
        Dim userMgr As IEdmUserMgr5 = vault
        Dim loggedInUser As IEdmUser5 = userMgr.GetLoggedInUser()
        'only get the drawings
        Dim drawings As New Dictionary(Of Tuple(Of IEdmFile5, IEdmFolder5), Tuple(Of IEdmFile5, IEdmFolder5))

        For Each item As EdmCmdData In ppoData

            Dim file As IEdmFile5 = vault.GetObject(EdmObjectType.EdmObject_File, item.mlObjectID1)
            Dim folder As IEdmFolder5 = vault.GetObject(EdmObjectType.EdmObject_Folder, item.mlObjectID2)

            If file.Name.ToLower().EndsWith(".slddrw") = False Then
                Continue For
            End If

            Dim associatedModel = GetAssociatedDocument(file, folder)

            If (associatedModel.Item1 Is Nothing) Then
                Continue For
            End If

            drawings.Add(New Tuple(Of IEdmFile5, IEdmFolder5)(file, folder), New Tuple(Of IEdmFile5, IEdmFolder5)(associatedModel.Item1, associatedModel.Item2))

        Next


        If (drawings.Count = 0) Then
            Throw New Exception("No found in the affected items by this task")
        End If


        Dim variableNames As String()

        Dim variableNamesStrRaw = instance.GetValEx(AddIn.STORAGEKEY)

        If String.IsNullOrWhiteSpace(variableNamesStrRaw) Then
            variableNamesStrRaw = String.Empty
        End If

        variableNames = variableNamesStrRaw.ToString().Split(";")


        If (variableNames Is Nothing Or variableNames.Length = 0) Then
            Throw New Exception("No variables defined in the task.")
        End If

        errorLogs.Clear()

        Dim index As Integer = 0
        Dim count As Integer = drawings.Count


        instance.SetProgressRange(count, 0, String.Empty)

        For Each item In drawings

            index = index + 1

            Try
                Dim drawingFile As IEdmFile5 = item.Key.Item1
                Dim drawingFolder As IEdmFolder5 = item.Key.Item2
                Dim drawingAssociatedModelFile As IEdmFile5 = item.Value.Item1
                Dim drawingAssociatedModelFolder As IEdmFolder5 = item.Value.Item2

                HandleCancellationRequest(instance)
                HandleSuspensionRequest(instance)

                ReportProgress(instance, index, $"Processing {drawingFile.Name}")



                Dim poVariables As Dictionary(Of String, String) = New Dictionary(Of String, String)

                For Each variableName In variableNames

                    poVariables.Add(variableName, Nothing)

                Next
                'pull variables from associated document
                GetVariables(drawingAssociatedModelFile, drawingAssociatedModelFolder, poVariables, errorLogs)


                Dim areThereAnyNonEmptyValues = poVariables.Values.Any(Function(x As String)
                                                                           Return String.IsNullOrWhiteSpace(x) = False
                                                                       End Function)
                If areThereAnyNonEmptyValues = False Then
                    Throw New Exception($"{drawingFile.Name}: No changes peformed.")
                End If

                ' check if we can check out the drawing
                If drawingFile.IsLocked Then
                    If drawingFile.LockedByUser.ID <> loggedInUser.ID Then
                        Throw New Exception($"{drawingFile.Name}: File is locked by {drawingFile.LockedByUser.Name}")
                    End If
                End If

                If drawingFile.IsLocked Then
                    If drawingFile.LockedOnComputer <> Environment.MachineName Then
                        Throw New Exception($"{drawingFile.Name}: File is locked on another computer {drawingFile.LockedOnComputer}")
                    End If
                End If

                If drawingFile.IsLocked = False Then
                    'locally cache the drawing and its references
                    drawingFile.GetFileCopy(handle, Nothing, Nothing, EdmGetFlag.EdmGet_Refs + EdmGetFlag.EdmGet_RefsVerLatest)
                    'check out drawing
                    drawingFile.LockFile(drawingFolder.ID, handle)
                End If

                SetVariables(drawingFile, drawingFolder, poVariables, errorLogs)
                'check drawing back into the vault

                'commit changes
                drawingFile.UnlockFile(handle, "Checked in by task")

            Catch ex As Exception
                errorLogs.AppendLine(ex.Message)
            End Try

        Next



    End Sub

    Private Sub ReportProgress(instance As IEdmTaskInstance, index As Integer, message As String)


        instance.SetProgressPos(index, message)
    End Sub

    Private Function GetAssociatedDocument(ByRef drawing As IEdmFile5, ByRef drawingFolder As IEdmFolder5) As Tuple(Of IEdmFile5, IEdmFolder5)

        Dim vault As IEdmVault5 = drawing.Vault

        Dim associatedModel As IEdmFile5 = Nothing
        Dim associatedModelFolder As IEdmFolder5 = Nothing

        Dim affectedDocumentRootReference As IEdmReference5 = drawing.GetReferenceTree(drawingFolder.ID)

        Dim position As IEdmPos5 = affectedDocumentRootReference.GetFirstChildPosition(ADDIN_NAME, True, False)

        Dim firstReference As IEdmReference5 = affectedDocumentRootReference.GetNextChild(position)

        If firstReference Is Nothing Then
            Return New Tuple(Of IEdmFile5, IEdmFolder5)(Nothing, Nothing)
        End If

        associatedModel = vault.GetObject(EdmObjectType.EdmObject_File, firstReference.FileID)

        associatedModelFolder = vault.GetObject(EdmObjectType.EdmObject_Folder, firstReference.FolderID)

        Return New Tuple(Of IEdmFile5, IEdmFolder5)(associatedModel, associatedModelFolder)

    End Function

    Private Sub HandleCancellationRequest(ByRef instance As IEdmTaskInstance)
        If instance.GetStatus() = EdmTaskStatus.EdmTaskStat_CancelPending Then
            instance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneCancelled)
            Throw New CancellationException("Task cancelled by user.")
        End If



    End Sub

    Private Sub HandleSuspensionRequest(ByRef instance As IEdmTaskInstance)
        'Handle temporary suspension of the task
        If instance.GetStatus() = EdmTaskStatus.EdmTaskStat_SuspensionPending Then
            instance.SetStatus(EdmTaskStatus.EdmTaskStat_Suspended)
            While instance.GetStatus() = EdmTaskStatus.EdmTaskStat_Suspended
                System.Threading.Thread.Sleep(1000)
            End While
            If instance.GetStatus() = EdmTaskStatus.EdmTaskStat_ResumePending Then
                instance.SetStatus(EdmTaskStatus.EdmTaskStat_Running)
            End If
        End If
    End Sub

End Class
