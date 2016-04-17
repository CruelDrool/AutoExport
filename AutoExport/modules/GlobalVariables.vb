Module GlobalVariables
    Friend iniFileName As String = "AutoExport.ini"
    Friend iniFilePath As String
    Friend iniFileLoaded As Boolean = False
    Friend AppDataFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), My.Application.Info.CompanyName, My.Application.Info.Title)
    Friend SystemDrive As String = Path.GetPathRoot(Environment.SystemDirectory)
    Friend ExportFolders(0) As ExportFolder
End Module
