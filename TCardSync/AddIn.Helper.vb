Imports EPDM.Interop.epdm
Imports System.Text

Partial Public Class AddIn

    Public Function GetVariableNames(ByRef Vault As IEdmVault10) As String()

        Dim variableNamesStoredRaw As String = String.Empty

        Dim storage As IEdmDictionary5 = Vault.GetDictionary(ADDIN_NAME, True)

        storage.StringGetAt(AddIn.STORAGEKEY, variableNamesStoredRaw)

        Dim variableNames As New List(Of String)

        If Not String.IsNullOrWhiteSpace(variableNamesStoredRaw) Then

            variableNames.AddRange(variableNamesStoredRaw.Split(";").Select(Function(x) x.Trim()).Where(Function(x) Not String.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
        End If

        Return variableNames.ToArray()
    End Function

    Public Sub GetVariables(ByRef poFile As IEdmFile5, ByRef poFolder As IEdmFolder5, ByRef poVariables As Dictionary(Of String, String), ByRef _StringBuilder As StringBuilder)
        Dim variableNames As String() = poVariables.Keys.ToArray()

        Dim fileEnumerator As IEdmEnumeratorVariable8 = poFile.GetEnumeratorVariable()
        Dim variableName As String = Nothing




        For Each variableName In variableNames
            Try
                Dim variableValue As Object = Nothing

                'read the value directly from the database - the file does not need to be locally cached
                fileEnumerator.GetVarFromDb(variableName, "@", variableValue)

                If variableValue IsNot Nothing Then
                    poVariables(variableName) = variableValue.ToString()
                End If

            Catch ex As Exception

                _StringBuilder.AppendLine(poFile.Name & $"[{variableName}]" & ": " & ex.Message)
            End Try

        Next


            fileEnumerator.CloseFile(True)




    End Sub

    Public Sub SetVariables(ByRef poFile As IEdmFile5, ByRef poFolder As IEdmFolder5, ByRef poVariables As Dictionary(Of String, String), ByRef _StringBuilder As StringBuilder)
        Dim variableNames As String() = poVariables.Keys.ToArray()

        Dim fileEnumerator As IEdmEnumeratorVariable8 = poFile.GetEnumeratorVariable()
        Dim variableName As String = Nothing


        For Each variableName In variableNames
            Try
                Dim variableValue As Object = poVariables(variableName)

                fileEnumerator.SetVar(variableName, "@", variableValue)

            Catch ex As Exception
                _StringBuilder.AppendLine(poFile.Name & $"[{variableName}]" & ": " & ex.Message)
            End Try
        Next




        fileEnumerator.CloseFile(True)



    End Sub
End Class
