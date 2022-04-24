Imports System.IO
Imports Microsoft.Win32

Public Class Stealth

    Private Shared _winPath As String = ShellApi.GetFolder(ShellApi.CSIDL.WINDOWS)
    Private Shared _sysPath As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
    Private Shared _installerPath As String = Path.Combine(_winPath, "Installer")
    Private Shared _appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Private Shared _sharedData As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)


    Private Shared Sub DeleteKey(ByVal keyName As String, ByVal throwOnError As Boolean)
        Try
            Registry.LocalMachine.DeleteSubKeyTree(keyName)
        Catch ex As Exception
            If throwOnError Then
                Throw ex
            End If
        End Try
    End Sub

End Class
