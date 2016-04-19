Imports System.Text.RegularExpressions

Module Functions
    Friend Sub ErrorSave(ByVal message As String)
        Dim ErrorMsg As String = "[" & Now & "]" & " [error] " & message
        My.Computer.FileSystem.WriteAllText(Path.Combine(AppDataFolder, "Error.log"), ErrorMsg & vbCrLf, True)
        Debug.Print(ErrorMsg)
    End Sub

    Friend Function FileExists(ByVal path As String) As Boolean
        Return My.Computer.FileSystem.FileExists(path)
    End Function

    Friend Function FolderExists(ByVal path As String) As Boolean
        Return My.Computer.FileSystem.DirectoryExists(path)
    End Function

    Friend Sub CreateDirectory(ByVal directory As String)
        Try
            My.Computer.FileSystem.CreateDirectory(directory)
        Catch ex As Exception
            ErrorSave("Unable to create directory: " & directory & " (" & ex.Message & ")")
        End Try
    End Sub

    Friend Function GetConfigValue(ByVal section As String, ByVal key As String) As String

        Dim retBuff As String = Space(1024)
        Dim reBuffSize As Integer = Len(retBuff)
        Dim retVal As Integer = NativeMethods.GetPrivateProfileString(section, key, "", retBuff, reBuffSize, iniFilePath)

        Return Trim(Left(retBuff, retVal))
    End Function

    Friend Function SanitizeInput(ByVal Input As String) As String
        Dim sChr As String = GetConfigValue("Global", "ReplaceCharacter")

        For Each InvalidChar As Char In Path.GetInvalidFileNameChars()
            Input = Replace(Input, InvalidChar, sChr)
        Next

        Return Input
    End Function

    Friend Function OutlookFolderExists(ByVal FolderName As String, ByVal Parent As Outlook.Folder) As Boolean
        For Each Folder As Outlook.Folder In Parent.Folders
            If Folder.Name = FolderName Then
                Return True
            End If
        Next
        Return False
    End Function

    Friend Function AddOutlookFolder(ByVal FolderName As String, ByVal Parent As Outlook.Folder) As Outlook.Folder
        If Not OutlookFolderExists(FolderName, parent) Then
            Parent.Folders.Add(FolderName)
        End If
        Return CType(Parent.Folders.Item(FolderName), Outlook.Folder)
    End Function

    Friend Sub DeleteOutlookFolder(ByVal FolderName As String, ByVal Parent As Outlook.Folder)
        If OutlookFolderExists(FolderName, Parent) Then
            Parent.Folders.Item(FolderName).Delete()
        End If
    End Sub

    Friend Sub Init(ByVal ParentOlFolder As Outlook.Folder, Optional ByVal Section As String = "", Optional sPath As String = "", Optional ParentSectionPath As String = "", Optional ByVal Level As Integer = 0, Optional ByVal PathFoundAt As Integer = 0)

        Dim TempSection As String
        If Section = "" Then
            TempSection = "/"
        Else
            TempSection = Section
        End If

        Dim SubFolders() As String = Split(GetConfigValue(TempSection, "Folders"), "|")
        Dim SectionPath As String = GetConfigValue(TempSection, "Path")

        If SectionPath <> "" Then
            'Found a Path, note down which level it was found at.
            PathFoundAt = Level
        End If

        If SectionPath = "" And ParentSectionPath <> "" Then
            'No Path was found. However, a Path was found in the Parent
            SectionPath = ParentSectionPath
        End If

        If SubFolders(0) <> "" Then
            For Each SubFolder As String In SubFolders
                SubFolder = Trim(SanitizeInput(SubFolder))

                Init(AddOutlookFolder(SubFolder, ParentOlFolder), Section & "/" & SubFolder, Path.Combine(sPath, SubFolder), SectionPath, Level + 1, PathFoundAt)
            Next
        ElseIf sPath <> "" Then
            If SectionPath = "" Then
                'No Path found, use the System Drive (i.e. C:\)
                SectionPath = SystemDrive
            End If

            Dim Match As String = Regex.Match(SectionPath, "^[A-Z]:", RegularExpressions.RegexOptions.IgnoreCase).ToString

            If Match <> "" And Not Regex.Match(SectionPath, "^[A-Z]:\\", RegularExpressions.RegexOptions.IgnoreCase).Success Then
                'A match for something like C:directory was found. Make it C:\directory
                SectionPath = Replace(SectionPath, Match, Match & Path.DirectorySeparatorChar, , 1)
            End If

            If PathFoundAt > 0 Then
                Dim TempArray() As String = Split(sPath, Path.DirectorySeparatorChar)
                Dim TempPath As String = ""
                Dim i As Integer = 1
                For Each Directory As String In TempArray
                    If i > PathFoundAt Then
                        TempPath = Path.Combine(TempPath, Directory)
                    End If
                    i += 1
                Next
                sPath = TempPath
            End If

            Dim FolderNumber As Integer = UBound(ExportFolders)
            ExportFolders(FolderNumber) = New ExportFolder
            ExportFolders(FolderNumber).Init(ParentOlFolder, Path.Combine(SectionPath, sPath))
            ReDim Preserve ExportFolders(FolderNumber + 1)
        End If
    End Sub
End Module
