Option Strict Off
Friend Class ExportFolder
    Private WithEvents FolderItems As Outlook.Items
    Private SaveDirectory As String
    Friend Sub SetFolder(ByVal Folder As Outlook.Folder)
        FolderItems = Folder.Items
    End Sub

    Friend Sub SetDirectory(ByVal folderPath As String)
        SaveDirectory = folderPath
    End Sub

    Friend Sub Init(ByVal Folder As Outlook.Folder, ByVal directory As String)
        SetFolder(Folder)
        SetDirectory(directory)
        If GetConfigValue("Global", "CreateDirectories").ToLower = "true" Then
            CreateDirectory(directory)
        End If
    End Sub

    Private Sub FolderItems_ItemAdd(ByVal Item As Object) Handles FolderItems.ItemAdd
        Dim sName As String = Trim(Item.Subject)
        Dim SavePath As String

        Dim DatestampFormat As String = GetConfigValue("Global", "DateStampFormat")

        If DatestampFormat <> "" Then
            sName = Format(Item.ReceivedTime, DatestampFormat) & sName
        End If

        sName = SanitizeInput(sName)
        sName &= ".msg"

        If FolderExists(SaveDirectory) Then
            SavePath = Path.Combine(SaveDirectory, sName)
            Try
                Item.SaveAs(SavePath, Outlook.OlSaveAsType.olMSGUnicode)
            Catch ex As Exception
                ErrorSave("Unable to save: " & SavePath & " (" & ex.Message & ")")
            End Try

            If FileExists(SavePath) Then
                'Save successful.
                'Delete the mail-item?
                'Item.Delete
            Else
                'Unsuccessful save
                ErrorSave("File was not saved: " & SavePath)
            End If
        Else
            ErrorSave("Directory doesn't exist: " & SaveDirectory)
        End If
        Exit Sub
    End Sub
End Class
