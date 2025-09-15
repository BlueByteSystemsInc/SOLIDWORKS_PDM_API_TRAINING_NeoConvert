Imports System.Reflection.Emit
Imports BlueByte.SOLIDWORKS.Extensions
Imports EPDM.Interop.epdm

Partial Public Class AddIn



    Private Sub HandlesTaskRun(poCmd As EdmCmd, ppoData() As EdmCmdData)

        Dim handle = poCmd.mlParentWnd
        Dim vault As IEdmVault5 = poCmd.mpoVault
        Dim instance As IEdmTaskInstance = poCmd.mpoExtra
        Dim userMgr As IEdmUserMgr5 = vault
        Dim loggedInUser As IEdmUser5 = userMgr.GetLoggedInUser()



        Dim extensionNames As String()

        Dim extensionNamesStrRaw = instance.GetValEx(AddIn.STORAGEKEY)

        If String.IsNullOrWhiteSpace(extensionNamesStrRaw) Then
            extensionNamesStrRaw = String.Empty
        End If

        extensionNames = extensionNamesStrRaw.ToString().Split(";")


        If (extensionNames Is Nothing Or extensionNames.Length = 0) Then
            Throw New Exception("No extensions defined in the task.")
        End If

        errorLogs.Clear()

        Dim index As Integer = 0
        Dim count As Integer = ppoData.Count


        Dim solidworksInstanceManager = New SOLIDWORKSInstanceManager()

        Dim Any = ppoData.ToList().Any(Function(x As EdmCmdData) x.mbsStrData1.ToLower().EndsWith(".sldprt"))

        If Any = False Then
            Throw New Exception("No parts selected. This task only processes parts.")
        End If


        Dim swApp = solidworksInstanceManager.GetNewInstance()

        swApp.Visible = True

        instance.SetProgressRange(count, 0, String.Empty)

        For Each item In ppoData

            index = index + 1

            Try

                Dim file As IEdmFile5 = vault.GetObject(EdmObjectType.EdmObject_File, item.mlObjectID1)

                Dim folder As IEdmFolder5 = vault.GetObject(EdmObjectType.EdmObject_Folder, item.mlObjectID2)

                HandleCancellationRequest(instance)

                HandleSuspensionRequest(instance)

                ReportProgress(instance, index, $"Processing {file.Name}")


                ' we are only going to process parts and convert them to step and dxf 

                If file.Name.ToLower().EndsWith(".sldprt") = False Then
                    Continue For
                End If

                ' cache the part file 

                file.GetFileCopy(handle)

                Dim warnings As String() = Nothing
                Dim errors As String() = Nothing
                ' open solidworks and convert the file
                Dim modelRet = swApp.OpenDocument(file.GetLocalPath(folder.ID), errors, warnings)

                If modelRet.Item2 Is Nothing Then
                    swApp.CloseAllDocuments(True)
                    Continue For
                End If

                Dim model = modelRet.Item2

                'save document 
                For Each ext In extensionNames
                    If String.IsNullOrWhiteSpace(ext) Then
                        Continue For
                    End If
                    Try
                        ' we could create a new folder with the extension name 
                        Dim subFolder As IEdmFolder5 = Nothing
                        Try
                            subFolder = folder.GetSubFolder(ext.Trim("."))
                        Catch ex As Exception
                            'folder does not exist 
                        End Try

                        If subFolder Is Nothing Then
                            folder.AddFolder(handle, ext.Trim("."))
                            subFolder = folder.GetSubFolder(ext.Trim())
                        End If

                        Dim targetCompleteFileName As String

                        Dim temporaryFileName As String

                        targetCompleteFileName = $"{subFolder.LocalPath}\{System.IO.Path.ChangeExtension(file.Name, ext.Trim())}"

                        temporaryFileName = $"{System.IO.Path.GetTempPath()}\{System.IO.Path.ChangeExtension(file.Name, ext.Trim())}"

                        model.SaveAs3(temporaryFileName, 0, 0)

                        'check if file already exists in vault
                        Dim existingFile As IEdmFile5 = Nothing
                        Try
                            existingFile = vault.GetFileFromPath(targetCompleteFileName)
                        Catch ex As Exception
                            'file does not exist 
                        End Try

                        If existingFile IsNot Nothing Then
                            'file exists, we need to delete it first 
                            subFolder.DeleteFile(handle, existingFile.ID, True)
                        End If


                        'add the file to the vault
                        Dim id As Integer = subFolder.AddFile(handle, temporaryFileName, True)

                        Dim addedFile As IEdmFile5 = vault.GetObject(EdmObjectType.EdmObject_File, id)

                        'check in the file
                        addedFile.LockFile(handle, False, "Checked in by neoConvert")


                    Catch ex As Exception
                        errorLogs.AppendLine($"Error saving {file.Name} as {ext}: {ex.Message}")
                    End Try
                Next


            Catch ex As Exception

                errorLogs.AppendLine(ex.Message)

            End Try

        Next


        swApp.CloseAllDocuments(True)
        swApp.ExitApp()


    End Sub

    Private Sub ReportProgress(instance As IEdmTaskInstance, index As Integer, message As String)


        instance.SetProgressPos(index, message)
    End Sub



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
