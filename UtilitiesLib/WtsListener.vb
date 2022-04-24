Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class WtsListener
    Inherits NativeWindow
    Implements IDisposable

    Private Const WS_MINIMIZE As Integer = &H20000000
    Private Const WS_EX_NOACTIVATE As Integer = &H8000000L

    Private Const WM_WTSSESSION_CHANGE As Integer = &H2B1
    Private Const WM_DESTROY As Integer = &H4400

    Private Enum NotifyScope
        NOTIFY_FOR_THIS_SESSION = 0
        NOTIFY_FOR_ALL_SESSIONS
    End Enum

    Public Enum WtsMessages
        WTS_CONSOLE_CONNECT = 1
        WTS_CONSOLE_DISCONNECT = 2
        WTS_REMOTE_CONNECT = 3
        WTS_REMOTE_DISCONNECT = 4
        WTS_SESSION_LOGON = 5
        WTS_SESSION_LOGOFF = 6
        WTS_SESSION_LOCK = 7
        WTS_SESSION_UNLOCK = 8
        WTS_SESSION_REMOTE_CONTROL = 9
    End Enum

    <DllImport("WtsApi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function WTSRegisterSessionNotification( _
        ByVal hWnd As IntPtr, _
        ByVal dwFlags As Integer) As Boolean
    End Function

    <DllImport("WtsApi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function WTSUnRegisterSessionNotification( _
        ByVal hWnd As IntPtr) As Boolean
    End Function

    Public Event ConsoleConnect(ByVal sessionID As Integer)
    Public Event ConsoleDisconnect(ByVal sessionID As Integer)
    Public Event RemoteConnect(ByVal sessionID As Integer)
    Public Event RemoteDisconnect(ByVal sessionID As Integer)
    Public Event SessionLogon(ByVal sessionID As Integer)
    Public Event SessionLogoff(ByVal sessionID As Integer)
    Public Event SessionLock(ByVal sessionID As Integer)
    Public Event SessionUnlock(ByVal sessionID As Integer)
    Public Event SessionRemoteControl(ByVal sessionID As Integer)

    Public Sub New()

        Dim cp As New CreateParams()

        cp.Caption = "WtsMessageWindow"
        cp.ClassName = "STATIC"

        cp.X = 0
        cp.Y = 0
        cp.Height = 0
        cp.Width = 0

        cp.Style = WS_MINIMIZE
        cp.ExStyle = WS_EX_NOACTIVATE

        Me.CreateHandle(cp)

        Logger.WriteEntry("WTSRegisterSessionNotifications registering on window: " + CStr(Me.Handle))
        If Not WTSRegisterSessionNotification(Me.Handle, NotifyScope.NOTIFY_FOR_ALL_SESSIONS) Then
            Logger.WriteEntry("Registering WTS Notifications failed: " + New Win32Exception(Marshal.GetLastWin32Error()).Message, EventLogEntryType.Error)
        End If
    End Sub

    Protected _disposed As Boolean
    Public Overridable Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Me.Dispose(False)
        MyBase.Finalize()
    End Sub

    Private Sub Dispose(ByVal disposing As Boolean)
        If _disposed Then Exit Sub

        If disposing Then
            'call Dispose on all contained objects
        End If
        Logger.WriteEntry("WTSUnRegisterSessionNotification from window: " + CStr(Me.Handle))
        If Not WTSUnRegisterSessionNotification(Me.Handle) Then
            Logger.WriteEntry("WTS Unregistering notification failed: " + New Win32Exception(Marshal.GetLastWin32Error()).Message, EventLogEntryType.Warning)
        End If
        _disposed = True
    End Sub

    <System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")> _
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_WTSSESSION_CHANGE Then
            Dim sessionID As Integer = CInt(m.LParam)
            Select Case CType(m.WParam, WtsMessages)
                Case WtsMessages.WTS_CONSOLE_CONNECT
                    RaiseEvent ConsoleConnect(sessionID)
                Case WtsMessages.WTS_CONSOLE_DISCONNECT
                    RaiseEvent ConsoleDisconnect(sessionID)
                Case WtsMessages.WTS_SESSION_LOGON
                    RaiseEvent SessionLogon(sessionID)
                Case WtsMessages.WTS_SESSION_LOGOFF
                    RaiseEvent SessionLogoff(sessionID)
                Case WtsMessages.WTS_SESSION_LOCK
                    RaiseEvent SessionLock(sessionID)
                Case WtsMessages.WTS_SESSION_UNLOCK
                    RaiseEvent SessionUnlock(sessionID)
                Case WtsMessages.WTS_REMOTE_CONNECT
                    RaiseEvent RemoteConnect(sessionID)
                Case WtsMessages.WTS_REMOTE_DISCONNECT
                    RaiseEvent RemoteDisconnect(sessionID)
                Case WtsMessages.WTS_SESSION_REMOTE_CONTROL
                    RaiseEvent SessionRemoteControl(sessionID)
                Case Else
            End Select
        ElseIf m.Msg = WM_DESTROY Then
            Me.Dispose()
        End If
        MyBase.WndProc(m)
    End Sub

End Class
