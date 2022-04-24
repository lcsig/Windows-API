Imports Microsoft.Win32

Public Class RegConstants

    Public Const Key_NTCurrentVersion As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    Public Const Key_Run As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    Public Const Key_GP_System As String = "SOFTWARE\Policies\Microsoft\Windows\System"
    Public Const Key_AppPaths As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
    Public Const Key_ProfileList As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Public Const Key_Winlogon As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Public Const Key_SpecialAccounts_UserList As String = Key_Winlogon + "\SpecialAccounts\UserList"
    Public Const Key_Policies_System As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system"
    Public Const Key_Explorer As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
    Public Const Key_InternetExplorer As String = "SOFTWARE\Microsoft\Internet Explorer"
    Public Const Key_GroupPolicy As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy"
    Public Const Key_SessionData As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\SessionData"
    Public Const Key_Lsa As String = "SYSTEM\CurrentControlSet\Control\Lsa"
    Public Const Key_SharedDlls As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDlls"
    Public Const Key_Services As String = "SYSTEM\CurrentControlSet\Services"
    Public Const Key_SystemControl As String = "SYSTEM\CurrentControlSet\Control"
    Public Const Key_ShellFolders As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

    Public Const Value_AllowFUS_xp As String = "AllowMultipleTSSessions"
    Public Const Value_HideFUS_vista As String = "HideFastUserSwitching"
    Public Const Value_LimitBlankPasswordUse As String = "LimitBlankPasswordUse"


    Public Shared Sub AddServicesPipeTimeout()
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(Key_SystemControl, True)
            key.SetValue("ServicesPipeTimeout", 60000, RegistryValueKind.DWord)
        End Using
    End Sub
End Class

