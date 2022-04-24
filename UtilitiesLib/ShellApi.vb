Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text
Imports System.Security.Principal

Public Class ShellApi

    Private Const S_OK As Integer = 0
    Private Const MAX_PATH As Integer = 260

    Private Enum SHERB
        SHERB_NOCONFIRMATION = &H1
        SHERB_NOPROGRESSUI = &H2
        SHERB_NOSOUND = &H4
    End Enum

    Public Enum CSIDL
        ADMINTOOLS = &H30
        ALTSTARTUP = &H1D
        APPDATA = &H1A
        BITBUCKET = &HA
        CDBURN_AREA = &H3B
        COMMON_ADMINTOOLS = &H2F
        COMMON_ALTSTARTUP = &H1E
        COMMON_APPDATA = &H23
        COMMON_DESKTOPDIRECTORY = &H19
        COMMON_DOCUMENTS = &H2E
        COMMON_FAVORITES = &H1F
        COMMON_MUSIC = &H35
        COMMON_OEM_LINKS = &H3A
        COMMON_PICTURES = &H36
        COMMON_PROGRAMS = &H17
        COMMON_STARTMENU = &H16
        COMMON_STARTUP = &H18
        COMMON_TEMPLATES = &H2D
        COMMON_VIDEO = &H37
        COMPUTERSNEARME = &H3D
        CONNECTIONS = &H31
        CONTROLS = &H3
        COOKIES = &H21
        DESKTOP = &H0
        DESKTOPDIRECTORY = &H10
        DRIVES = &H11
        FAVORITES = &H6
        FLAG_CREATE = &H8000
        FLAG_DONT_VERIFY = &H4000
        FLAG_MASK = &HFF00
        FLAG_NO_ALIAS = &H1000
        FLAG_PER_USER_INIT = &H800
        FONTS = &H14
        HISTORY = &H22
        INTERNET = &H1
        INTERNET_CACHE = &H20
        LOCAL_APPDATA = &H1C
        MYDOCUMENTS = &HC
        MYMUSIC = &HD
        MYPICTURES = &H27
        MYVIDEO = &HE
        NETHOOD = &H13
        NETWORK = &H12
        PERSONAL = &H5
        PRINTERS = &H4
        PRINTHOOD = &H1B
        PROFILE = &H28
        PROGRAM_FILES = &H26
        PROGRAM_FILES_COMMON = &H2B
        PROGRAM_FILES_COMMONX86 = &H2C
        PROGRAM_FILESX86 = &H2A
        PROGRAMS = &H2
        RECENT = &H8
        RESOURCES = &H38
        RESOURCES_LOCALIZED = &H39
        SENDTO = &H9
        STARTMENU = &HB
        STARTUP = &H7
        SYSTEM = &H25
        SYSTEMX86 = &H29
        TEMPLATES = &H15
        WINDOWS = &H24
    End Enum

    Private Structure SHItemID
        Dim cb As Integer
        Dim abID As Byte
    End Structure

    Private Structure ItemIDList
        Dim mkid As SHItemID
    End Structure

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SHEmptyRecycleBin( _
        ByVal hWnd As IntPtr, _
        ByVal pszRootPath As String, _
        ByVal dwFlags As SHERB) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SHGetFolderPath( _
        ByVal hWnd As IntPtr, _
        ByVal nFolder As CSIDL, _
        ByVal hToken As IntPtr, _
        ByVal dwFlags As Integer, _
        ByVal pszPath As StringBuilder) As Integer
    End Function

    <DllImport("shell32.dll")> _
    Private Shared Function SHGetSpecialFolderLocation( _
        ByVal hwndOwner As IntPtr, _
        ByVal nFolder As CSIDL, _
        ByRef pidlist As ItemIDList) As Integer
    End Function

    <DllImport("shell32.dll")> _
    Private Shared Function SHGetPathFromIDList( _
        ByVal pidlist As Integer, _
        ByVal lpBuffer As StringBuilder) As Integer
    End Function

    Public Shared Sub EmptyRecycleBin()
        Dim result As Integer
        result = SHEmptyRecycleBin( _
                        Nothing, _
                        Nothing, _
                        SHERB.SHERB_NOCONFIRMATION Or SHERB.SHERB_NOPROGRESSUI Or SHERB.SHERB_NOSOUND)
        If result = 0 Then
            'emptied
        ElseIf result = -2147418113 Then
            'already empty
        Else
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    Public Shared Function GetFolder(ByVal folder As CSIDL) As String
        Return GetFolder(folder, IntPtr.Zero)
    End Function

    Public Shared Function GetFolder(ByVal folder As CSIDL, ByVal hUserToken As IntPtr) As String
        Dim path As New StringBuilder(MAX_PATH)
        Dim result As Integer
        Dim pidList As ItemIDList

        result = ShellApi.SHGetSpecialFolderLocation(IntPtr.Zero, folder, pidList)
        If result <> S_OK Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        result = SHGetPathFromIDList(pidList.mkid.cb, path)
        If result = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return path.ToString()
    End Function

End Class
