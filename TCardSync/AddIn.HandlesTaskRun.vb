Imports EPDM.Interop.epdm

Partial Public Class AddIn



    Private Sub HandlesTaskRun(poCmd As EdmCmd, ppoData() As EdmCmdData)


        Dim vault As IEdmVault5 = poCmd.mpoVault
        Dim instance As IEdmTaskInstance = poCmd.mpoExtra
        'only get the drawings
        Dim drawings As New Dictionary(Of IEdmFile5, Tuple(Of IEdmFile5, IEdmFile5))

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

            drawings.Add(file, New Tuple(Of IEdmFile5, IEdmFile5)(associatedModel.Item1, associatedModel.Item2))

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

        For Each item In drawings

            Try
                Dim drawingFile As IEdmFile5 = item.Key
                Dim drawingAssociatedModelFile As IEdmFile5 = item.Value.Item1
                Dim drawingAssociatedModelFolder As IEdmFolder5 = item.Value.Item2

                HandleCancellationRequest(instance)
                HandleSuspensionRequest(instance)



                Dim poVariables As Dictionary(Of String, String) = Nothing

                'pull variables from associated document
                GetVariables(drawingAssociatedModelFile, drawingAssociatedModelFolder, poVariables, errorLogs)


                Dim areThereAnyNonEmptyValues = poVariables.Values.Any(Function(x As String)
                                                                           Return String.IsNullOrWhiteSpace(x) = False
                                                                       End Function)
                If areThereAnyNonEmptyValues = False Then
                    Throw New Exception($"{drawingFile.Name}: No changes peformed.")
                End If


                'check out drawing



                SetVariables(drawingFile, drawingFolder, poVariables, errorLogs)
                'check drawing back into the vault



            Catch ex As Exception
                errorLogs.AppendLine(ex.Message)
            End Try

        Next



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
            Exit Sub
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
