Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        CreateDirectory(AppDataFolder)

        Dim SearchPath As String
        For Each Drive As DriveInfo In DriveInfo.GetDrives()
            If Drive.DriveType = DriveType.Fixed Then
                SearchPath = Path.Combine(Drive.RootDirectory.ToString(), iniFileName)
                If FileExists(SearchPath) Then
                    iniFilePath = SearchPath
                    iniFileLoaded = True
                    Exit For
                End If
            End If
        Next

        'Dim searchPaths() As String = {SystemDrive, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Google Drive")}
        'Dim sPath As String
        'For Each searchPath As String In searchPaths
        '    sPath = Path.Combine(searchPath, "ExportFolders.ini")
        '    If FileExists(sPath) Then
        '        iniFile = sPath
        '        iniFileLoaded = True
        '        Exit For
        '    End If
        'Next

        If iniFileLoaded Then
            Dim ExternalINI As String = GetConfigValue("Global", "ExternalINI")
            If ExternalINI <> "" Then
                If FileExists(ExternalINI) Then
                    iniFilePath = ExternalINI
                Else
                    iniFileLoaded = False
                End If
            End If
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_MAPILogonComplete() Handles Application.MAPILogonComplete

        If Not iniFileLoaded Then
            ErrorSave("Configuration file could not be found.")
        Else
            Init(CType(Application.Session.Folders.Item(1), Outlook.Folder))
        End If

    End Sub
End Class
