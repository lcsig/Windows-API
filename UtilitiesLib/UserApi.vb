Imports System.Text
Imports System.Collections
Imports System.Management
Imports System.Security
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports Microsoft.Win32
Imports System.ComponentModel

Public Class UserApi

    Private Const SPI_SETSCREENSAVEACTIVE As Integer = 17
    Private Const KEYEVENTF_KEYUP As Integer = &H2
    Private Const VK_LWIN As Integer = &H5B

    Private Const WM_COMMAND As Integer = &H111
    Private Const MIN_ALL As Integer = 419
    Private Const MIN_ALL_UNDO As Integer = 416

    Public Const WM_GETTEXT As Integer = &HD
    Public Const WM_GETTEXTLENGTH As Integer = &HE

    Public Const SC_CLOSE As UInteger = 61536

    Public Enum MenuItemFlags As UInteger
        mfUnchecked = &H0    ' ... is not checked
        mfString = &H0    ' ... contains a string as label
        mfDisabled = &H2    ' ... is disabled
        mfGrayed = &H1    ' ... is grayed
        mfChecked = &H8    ' ... is checked
        mfPopop = &H10    ' ... Is a popup menu. Pass the menu handle of the popup menu into the ID parameter.
        mfBarBreak = &H20    ' ... is a bar break
        mfBreak = &H40    ' ... is a break
        mfByPosition = &H400    ' ... is identified by the position
        mfByCommand = &H0    ' ... is identified by its ID
        mfSeparator = &H800     ' ... is a seperator (String and ID parameters are ignored).
    End Enum

    Public Enum ExitOptions As Integer
        Logoff = &H0
        Shutdown = &H1
        Reboot = &H2
        Force = &H4
        PowerOff = &H8
        ForceIfHung = &H10
    End Enum

    Const SHTDN_REASON_MINOR_UNSTABLE As Integer = &H6
    Const SHTDN_REASON_MAJOR_APPLICATION As Integer = &H40000
    Const SHTDN_REASON_MINOR_OTHER As Integer = &H0
    Const SHTDN_REASON_MAJOR_OTHER As Integer = &H0

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure LASTINPUTINFO
        <MarshalAs(UnmanagedType.U4)> Public cbSize As Integer
        <MarshalAs(UnmanagedType.U4)> Public dwTime As Integer
    End Structure

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function SystemParametersInfo( _
        ByVal uAction As Integer, _
        ByVal uParam As Integer, _
        ByVal lpvParam As Integer, _
        ByVal fuWinIni As Integer) As Integer
    End Function

    Public Enum Cursors As Integer
        IDC_ARROW = 32512
        IDC_IBEAM = 32513
        IDC_WAIT = 32514
        IDC_CROSS = 32515
        IDC_UPARROW = 32516
        IDC_SIZENWSE = 32642
        IDC_SIZENESW = 32643
        IDC_SIZEWE = 32644
        IDC_SIZENS = 32645
        IDC_SIZEALL = 32646
        IDC_NO = 32648
        IDC_HAND = 32649
        IDC_APPSTARTING = 32650
        IDC_HELP = 32651
        IDC_ICON = 32641
        IDC_SIZE = 32640
    End Enum

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenDesktop(ByVal lpszDesktop As String, ByVal dwFlags As Integer, ByVal fInherit As Boolean, ByVal dwDesiredAccess As KernelApi.DesktopAccess) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function OpenDesktop(ByVal lpszDesktop As String, ByVal dwFlags As Integer, ByVal fInherit As Boolean, ByVal dwDesiredAccess As KernelApi.GenericAccess) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetThreadDesktop(ByVal hDesktop As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetDesktopWindow() As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function LoadCursor( _
        ByVal hInstance As Integer, _
        ByVal cursorName As Cursors) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function SetCursor( _
        ByVal hcursor As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Public Shared Function DestroyCursor( _
        ByVal hcursor As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function GetLastInputInfo(ByRef plii As LASTINPUTINFO) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function ExitWindowsEx( _
        ByVal flags As Integer, _
        ByVal reason As Integer) As Integer
    End Function

    Public Shared Sub ExitWindows(ByVal flags As ExitOptions)
        KernelApi.AddPrivilege(KernelApi.SE_SHUTDOWN_PRIVILEGE)
        Dim result As Integer
        result = ExitWindowsEx(flags, SHTDN_REASON_MAJOR_OTHER Or SHTDN_REASON_MINOR_OTHER)
        If result = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function LockWorkStation() As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Sub keybd_event( _
            ByVal bVk As Byte, _
            ByVal bScan As Byte, _
            ByVal dwFlags As Integer, _
            ByVal dwExtraInfo As Integer)
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function FindWindow( _
            ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SendMessage( _
            ByVal hwnd As Integer, _
            ByVal wMsg As Integer, _
            ByVal wParam As Integer, _
            ByVal lParam As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SendMessage( _
            ByVal hwnd As IntPtr, _
            ByVal wMsg As Integer, _
            ByVal wParam As IntPtr, _
            ByVal lParam As StringBuilder) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function GetSystemMenu( _
        ByVal hwnd As IntPtr, _
        ByVal bReset As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function ModifyMenu( _
        ByVal hMnu As IntPtr, _
        ByVal uPosition As UInteger, _
        ByVal uFlags As UInteger, _
        ByVal uIDNewItem As UIntPtr, _
        ByVal lpNewItem As String) As Boolean
    End Function

    Public Shared Function GetLastInputInSeconds() As Integer
        Dim idle As Integer = 0
        Dim info As UserApi.LASTINPUTINFO
        info.cbSize = Marshal.SizeOf(info)

        If GetLastInputInfo(info) Then
            idle = Environment.TickCount - info.dwTime
        End If

        If idle > 0 Then
            Return CInt(idle / 1000)
        Else
            Return 0
        End If
    End Function

    Public Shared Sub EnableScreenSaver(ByVal status As Boolean)
        Dim result As Integer
        result = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, CInt(status), 0, 0)
        If result = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    Public Shared Sub MinimizeAll()
        'keybd_event(VK_LWIN, 0, 0, 0)
        'keybd_event(Asc("M"), 0, 0, 0)
        'keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)

        Dim hWnd As Integer = FindWindow("Shell_TrayWnd", vbNullString)
        SendMessage(hWnd, WM_COMMAND, MIN_ALL, 0)
    End Sub

    Public Shared Function GetText(ByVal hWnd As Integer) As String
        Dim len As Integer
        len = SendMessage(hWnd, UserApi.WM_GETTEXTLENGTH, 0, 0) + 1

        Dim s As New StringBuilder(Space(len))
        len = SendMessage(New IntPtr(hWnd), UserApi.WM_GETTEXT, New IntPtr(len), s)

        Return Left(s.ToString(), len)
    End Function

    Public Shared Sub DisableClose(ByVal hWnd As IntPtr)
        Dim menu As IntPtr = UserApi.GetSystemMenu(hWnd, 0)
        UserApi.ModifyMenu(menu, UserApi.SC_CLOSE, MenuItemFlags.mfByCommand Or MenuItemFlags.mfGrayed, CType(0, UIntPtr), Nothing)
    End Sub
End Class

