Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class WindowsApi
    Public Const HWND_TOP As Integer = 0
    Public Const HWND_TOPMOST As Integer = -1
    Public Const HWND_NOTOPMOST As Integer = -2
    Public Const SWP_NOSIZE As Integer = &H1
    Public Const SWP_NOMOVE As Integer = &H2
    Public Const SWP_SHOWWINDOW As Integer = &H40

    Public Const GW_CHILD As Integer = 5

    Public Const INFINITE As Int32 = -1
    Public Const WAIT_TIMEOUT As Int32 = 258
    Public Declare Function WaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Int32, ByVal dwMilliseconds As Int32) As Int32

    Public Enum ShowWindowEnum
        SW_HIDE = 0
        SW_SHOWNORMAL = 1
        SW_NORMAL = 1
        SW_SHOWMINIMIZED = 2
        SW_SHOWMAXIMIZED = 3
        SW_MAXIMIZE = 3
        SW_SHOWNOACTIVATE = 4
        SW_SHOW = 5
        SW_MINIMIZE = 6
        SW_SHOWMINNOACTIVE = 7
        SW_SHOWNA = 8
        SW_RESTORE = 9
        SW_SHOWDEFAULT = 10
        SW_FORCEMINIMIZE = 11
        SW_MAX = 11
    End Enum

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetWindowPos( _
        ByVal hwnd As Integer, _
        ByVal hWndInsertAfter As Integer, _
        ByVal x As Integer, _
        ByVal y As Integer, _
        ByVal cx As Integer, _
        ByVal cy As Integer, _
        ByVal wFlags As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetForegroundWindow( _
        ByVal handle As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function ShowWindow( _
        ByVal handle As IntPtr, _
        ByVal nCmd As WindowsApi.ShowWindowEnum) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetWindow( _
        ByVal handle As IntPtr, _
        ByVal cmd As Integer) As IntPtr
    End Function

    Public Shared Sub SetAlwaysOnTop(ByVal hwnd As Integer, ByVal value As Boolean)
        If value Then
            SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
        Else
            SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
        End If
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function SetParent(ByVal hWnd As Integer, ByVal hWndParent As Integer) As Integer
    End Function

    Public Shared Sub SetParent(ByVal hwnd As IntPtr, ByVal hWndParent As IntPtr)
        Dim ret As Integer
        ret = SetParent(hwnd.ToInt32(), hWndParent.ToInt32())
        If ret = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function EnumWindows(ByVal lpEnumProc As EnumWindDelegate, _
        ByVal lParam As Int32) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function EnumChildWindows(ByVal hWnd As IntPtr, _
        ByVal lpEnumFunc As EnumChildWindDelegate, _
        ByRef lParam As IntPtr) As Int32
    End Function

    Delegate Function EnumWindDelegate(ByVal hWnd As Int32, _
        ByVal lParam As Int32) As Boolean

    Delegate Function EnumChildWindDelegate(ByVal hWnd As Int32, _
        ByVal lParam As Int32) As Boolean

    Private Shared enumChildren As String
    Private Shared enumParent As String
    Public Shared Function EnumChildCallback(ByVal hWnd As Int32, _
            ByVal lParam As Int32) As Boolean
        If enumChildren <> "" Then enumChildren += ", "
        enumChildren += CStr(hWnd) + " " + UserApi.GetText(hWnd)
        EnumChildCallback = True
    End Function

    Public Shared Function EnumWindowsCallback(ByVal hWnd As Integer, ByVal lParam As Integer) As Boolean
        enumChildren = ""

        Dim proc As New EnumChildWindDelegate(AddressOf EnumChildCallback)
        EnumChildWindows(New IntPtr(hWnd), proc, IntPtr.Zero)

        If enumParent <> "" Then enumParent += vbCrLf

        enumParent += "Parent: " + CStr(hWnd) + " "
        enumParent += UserApi.GetText(hWnd)
        If enumChildren <> "" Then enumParent += vbCrLf + vbTab + enumChildren

        Return True
    End Function

    Public Shared Function EnumAllWindows() As String

        enumParent = ""

        Dim proc As New EnumWindDelegate(AddressOf EnumWindowsCallback)
        EnumWindows(proc, 0)

        Return enumParent
    End Function

    Public Shared Function EnumChildWindows(ByVal hWnd As Integer) As String
        enumParent = ""
        enumChildren = ""

        Dim proc As New EnumChildWindDelegate(AddressOf EnumChildCallback)
        EnumChildWindows(New IntPtr(hWnd), proc, IntPtr.Zero)

        enumParent += "Parent: " + CStr(hWnd) + " "
        enumParent += UserApi.GetText(hWnd)
        enumParent += vbCrLf + vbTab + enumChildren

        Return enumParent
    End Function

    Public Shared Sub ShowDesktop()
        'Dim pman As IntPtr = New IntPtr(UserApi.FindWindow("ProgMan", Nothing))
        'Dim child As IntPtr = WindowsApi.GetWindow(pman, GW_CHILD)

        'WindowsApi.SetForegroundWindow(child)

        Dim clsidShell As New Guid("13709620-C279-11CE-A49E-444553540000")
        Dim shell As Object = Activator.CreateInstance(Type.GetTypeFromCLSID(clsidShell))

        shell.GetType().InvokeMember("ToggleDesktop", Reflection.BindingFlags.InvokeMethod, Nothing, shell, Nothing)
    End Sub

End Class
