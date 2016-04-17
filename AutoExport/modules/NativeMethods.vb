Friend NotInheritable Class NativeMethods
    Friend Declare Auto Function GetPrivateProfileString Lib "kernel32" ( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Integer, _
        ByVal lpFileName As String) As Integer

    Friend Declare Auto Function GetPrivateProfileInt Lib "kernel32" ( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As Integer, _
        ByVal lpFileName As String) As Integer
End Class
